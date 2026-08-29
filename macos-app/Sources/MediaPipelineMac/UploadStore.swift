import Foundation

enum UploadSourceRecovery {
    static func restoreRetainedSource(
        _ retainedURL: URL,
        preferredURL: URL,
        fileManager: FileManager = .default
    ) -> URL? {
        guard fileManager.fileExists(atPath: retainedURL.path) else { return nil }

        let directory = preferredURL.deletingLastPathComponent()
        let stem = preferredURL.deletingPathExtension().lastPathComponent
        let pathExtension = preferredURL.pathExtension
        for attempt in 0..<100 {
            let destination: URL
            if attempt == 0 {
                destination = preferredURL
            } else {
                let suffix = pathExtension.isEmpty ? "" : ".\(pathExtension)"
                destination = directory.appendingPathComponent(
                    "\(stem).upload-retry-\(attempt)\(suffix)"
                )
            }
            if fileManager.fileExists(atPath: destination.path) { continue }
            do {
                try fileManager.moveItem(at: retainedURL, to: destination)
                return destination
            } catch {
                if fileManager.fileExists(atPath: destination.path) { continue }
                return nil
            }
        }
        return nil
    }
}

@MainActor
final class UploadStore: ObservableObject {
    @Published private(set) var workspaces: [WorkspaceUploads] = []
    @Published private(set) var transfers: [UploadTransfer] = []
    @Published private(set) var isBusy = false
    @Published var errorMessage: String?

    private let worker: WorkerClient
    private var queueTask: Task<Void, Never>?

    init(worker: WorkerClient) {
        self.worker = worker
    }

    func refresh(pipelineRoot: URL, workspaceNames: [String]) {
        workspaces = workspaceNames.map { workspace in
            let directory = pipelineRoot
                .appendingPathComponent(workspace, isDirectory: true)
                .appendingPathComponent("sync", isDirectory: true)
            let keys: Set<URLResourceKey> = [.isRegularFileKey, .fileSizeKey]
            let urls = (try? FileManager.default.contentsOfDirectory(
                at: directory,
                includingPropertiesForKeys: Array(keys),
                options: [.skipsHiddenFiles]
            )) ?? []
            let files = urls.compactMap { url -> StagedFile? in
                guard let values = try? url.resourceValues(forKeys: keys),
                      values.isRegularFile == true else { return nil }
                return StagedFile(
                    url: url,
                    workspace: workspace,
                    bytes: Int64(values.fileSize ?? 0)
                )
            }.sorted { $0.name.localizedStandardCompare($1.name) == .orderedAscending }
            return WorkspaceUploads(name: workspace, files: files)
        }
    }

    func upload(_ files: [StagedFile]) {
        guard !files.isEmpty, queueTask == nil else { return }
        isBusy = true
        errorMessage = nil
        queueTask = Task { [weak self] in
            guard let self else { return }
            for file in files {
                if Task.isCancelled { break }
                let transferID = UUID()
                transfers.insert(UploadTransfer(
                    id: transferID,
                    fileURL: file.url,
                    workspace: file.workspace,
                    phase: "Queued",
                    chunksSent: 0,
                    chunks: 0,
                    bytesSent: 0,
                    bytes: file.bytes,
                    error: nil
                ), at: 0)
                do {
                    let terminal = try await worker.upload(file: file.url, workspace: file.workspace) {
                        [weak self] progress in
                        self?.apply(progress, transferID: transferID)
                    }
                    apply(terminal, transferID: transferID)
                    moveRetainedSourceToTrash(terminal, transferID: transferID)
                } catch {
                    updateTransfer(transferID) { transfer in
                        transfer.phase = "Failed"
                        transfer.error = error.localizedDescription
                    }
                    errorMessage = error.localizedDescription
                }
            }
            isBusy = false
            queueTask = nil
        }
    }

    func cancel() {
        queueTask?.cancel()
        worker.cancelUpload()
    }

    private func apply(_ progress: UploadProgressEvent, transferID: UUID) {
        updateTransfer(transferID) { transfer in
            transfer.phase = progress.phase
            transfer.chunksSent = progress.chunksSent
            transfer.chunks = progress.chunks
            transfer.bytesSent = progress.bytesSent
            transfer.bytes = progress.bytes
            transfer.error = progress.error
        }
    }

    private func moveRetainedSourceToTrash(
        _ terminal: UploadProgressEvent,
        transferID: UUID
    ) {
        guard terminal.phase.caseInsensitiveCompare("done") == .orderedSame,
              !terminal.sourceDeleted,
              let retainedSourcePath = terminal.retainedSourcePath else { return }
        let retainedURL = URL(fileURLWithPath: retainedSourcePath)
        let preferredURL = transfers.first(where: { $0.id == transferID })?.fileURL
        do {
            try FileManager.default.trashItem(
                at: retainedURL,
                resultingItemURL: nil
            )
        } catch {
            let restoredURL = preferredURL.flatMap {
                UploadSourceRecovery.restoreRetainedSource(retainedURL, preferredURL: $0)
            }
            updateTransfer(transferID) { transfer in
                let location = restoredURL?.path ?? retainedURL.path
                transfer.error = "Upload verified, but the local file could not be moved to Trash. The source was retained at \(location): \(error.localizedDescription)"
            }
        }
    }

    private func updateTransfer(_ id: UUID, change: (inout UploadTransfer) -> Void) {
        guard let index = transfers.firstIndex(where: { $0.id == id }) else { return }
        change(&transfers[index])
    }
}
