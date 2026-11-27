import AppKit
import Combine
import Darwin
import Foundation

enum SavedMediaCopyStatus {
    case preparing(preset: String, workspace: String)
    case ready(URL, copied: Bool)
    case failed(String)

    var isPreparing: Bool {
        if case .preparing = self { return true }
        return false
    }
}

struct SavedMediaCopy {
    let itemID: UUID
    var status: SavedMediaCopyStatus
}

@MainActor
final class AppModel: ObservableObject {
    @Published var selection: AppSection? = .activity
    @Published private(set) var status: WorkerStatus?
    @Published private(set) var workerRunning = false
    @Published private(set) var workerBusy = false
    @Published var errorMessage: String?
    @Published var notice: String?
    @Published private(set) var savedMediaCopy: SavedMediaCopy?

    let paths: AppPaths
    let configuration: ConfigurationStore
    let worker: WorkerClient
    let activity: ActivityStore
    let uploads: UploadStore
    let savedMedia: SavedMediaStore
    let loginItem = LoginItemController()
    let notifications = NotificationController()

    private var monitoringTask: Task<Void, Never>?
    private var savedMediaCopyTask: Task<Void, Never>?
    private let clipboard: NSPasteboard
    private var isShuttingDown = false
    private var failureStateActive = false
    private var activePipelineRoot: URL
    private var observations: Set<AnyCancellable> = []

    init(paths: AppPaths? = nil, workerExecutable: URL? = AppPaths.workerExecutable(), clipboard: NSPasteboard = .general) {
        do {
            let paths = try paths ?? AppPaths.current()
            let configuration = ConfigurationStore(configURL: paths.configFile)
            try configuration.load()
            let worker = WorkerClient(
                configURL: paths.configFile,
                executableURL: workerExecutable
            )
            self.paths = paths
            self.configuration = configuration
            self.worker = worker
            self.clipboard = clipboard
            let activity = ActivityStore()
            let uploads = UploadStore(worker: worker)
            self.activity = activity
            self.uploads = uploads
            self.savedMedia = try SavedMediaStore(
                directory: paths.applicationSupport.appendingPathComponent("saved-media", isDirectory: true)
            )
            activePipelineRoot = configuration.pipelineRoot

            // Child stores publish their fast-moving rows independently. Forward those changes
            // so the window and menu bar update immediately instead of waiting for a status tick.
            activity.objectWillChange
                .sink { [weak self] _ in self?.objectWillChange.send() }
                .store(in: &observations)
            uploads.objectWillChange
                .sink { [weak self] _ in self?.objectWillChange.send() }
                .store(in: &observations)
        } catch {
            fatalError("Media Pipeline could not prepare its application data: \(error)")
        }
    }

    var queuedCount: Int {
        status?.lanes.reduce(0) { $0 + $1.queued } ?? 0
    }

    var encoderName: String {
        status?.encoder ?? "Not selected"
    }

    var statusText: String {
        if workerBusy { return "Updating worker" }
        if workerRunning { return status?.pausedAll == true ? "Worker paused" : "Worker running" }
        return "Worker stopped"
    }

    var menuBarSymbol: String {
        if !workerRunning { return "circle.dotted" }
        if !activity.failures.isEmpty { return "exclamationmark.circle.fill" }
        if !activity.running.isEmpty { return "circle.inset.filled" }
        return "circle.fill"
    }

    func startMonitoring() {
        guard monitoringTask == nil, !isShuttingDown else { return }
        notifications.prepare()
        monitoringTask = Task { [weak self] in
            guard let self else { return }
            await refresh()
            guard !Task.isCancelled else { return }
            if !workerRunning,
               worker.hasExecutable,
               UserDefaults.standard.object(forKey: "startWorkerOnLaunch") as? Bool != false,
               ProcessInfo.processInfo.environment["XCODE_RUNNING_FOR_PREVIEWS"] != "1" {
                await startWorker()
            }

            while !Task.isCancelled {
                try? await Task.sleep(for: .seconds(2))
                guard !Task.isCancelled else { return }
                await refresh()
            }
        }
    }

    func stopMonitoring() {
        monitoringTask?.cancel()
        monitoringTask = nil
    }

    func shutdown() async {
        isShuttingDown = true
        stopMonitoring()
        savedMediaCopyTask?.cancel()
        uploads.cancel()
        await worker.shutdown(existingWorkerPID: status?.pid)
    }

    func refresh() async {
        let root = activePipelineRoot
        let statusURL = root
            .appendingPathComponent("status", isDirectory: true)
            .appendingPathComponent("watcher.json")
        if let data = try? Data(contentsOf: statusURL),
           let decoded = try? JSONDecoder().decode(WorkerStatus.self, from: data) {
            status = decoded
            // The status snapshot is intentionally only rewritten between queue scans. A long
            // encode can therefore make its timestamp old even though the worker is healthy.
            workerRunning = worker.isRunning(decoded)
        } else {
            status = nil
            workerRunning = worker.managedWorkerIsRunning
        }

        activity.refresh(pipelineRoot: root)
        activity.reconcile(workerRunning: workerRunning)
        if !failureStateActive, let failure = activity.failures.first {
            notifications.processingFailed(failure)
        }
        failureStateActive = !activity.failures.isEmpty
        uploads.refresh(
            pipelineRoot: root,
            workspaceNames: status?.workspaces ?? configuration.workspaces
        )
    }

    func startWorker() async {
        guard !workerBusy else { return }
        workerBusy = true
        defer { workerBusy = false }
        do {
            try await startWorkerOperation()
            notice = "Worker started."
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func stopWorker() async {
        guard !workerBusy else { return }
        workerBusy = true
        defer { workerBusy = false }
        do {
            try await stopWorkerOperation()
            notice = "Worker stopped."
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func restartWorker() async {
        guard !workerBusy else { return }
        workerBusy = true
        defer { workerBusy = false }
        do {
            if workerRunning || worker.managedWorkerIsRunning {
                try await stopWorkerOperation()
            }
            try await startWorkerOperation()
            notice = "Worker restarted."
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func togglePauseAll() async {
        do {
            let paused = !(status?.pausedAll ?? false)
            try await worker.setPaused(paused)
            await refresh()
            notice = paused ? "Worker paused." : "Worker resumed."
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func togglePauseLane(_ lane: LaneStatus) async {
        do {
            let paused = !lane.paused
            try await worker.setPaused(
                paused,
                preset: lane.preset,
                workspace: lane.workspace
            )
            await refresh()
            if paused {
                notice = "Paused \(lane.workspace) · \(lane.preset)."
            } else if status?.pausedAll == true {
                notice = "Lane resumed, but Pause All is still on."
            } else {
                notice = "Resumed \(lane.workspace) · \(lane.preset)."
            }
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func requeue(_ job: JobSnapshot) async {
        do {
            let result = try await worker.requeue(preset: job.preset, workspace: job.workspace)
            if !result.succeeded {
                throw WorkerClientError.commandFailed(result.standardError)
            }
            activity.dismissFailure(id: job.id)
            await refresh()
            notice = result.standardOutput.trimmingCharacters(in: .whitespacesAndNewlines)
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func queueSavedMedia(
        _ item: SavedMedia, copies: Int, preset: String, workspace: String, copyToClipboard: Bool = false
    ) async {
        guard !workerBusy, !copyToClipboard || savedMediaCopy?.status.isPreparing != true else { return }
        let clipboardVersion = copyToClipboard ? clipboard.changeCount : nil
        let pipelineRoot = activePipelineRoot
        if copyToClipboard {
            savedMediaCopy = SavedMediaCopy(itemID: item.id, status: .preparing(preset: preset, workspace: workspace))
        }
        workerBusy = true
        defer { workerBusy = false }
        do {
            let result = try await worker.queueMedia(
                file: savedMedia.sourceURL(for: item), copies: copyToClipboard ? 1 : copies,
                preset: preset, workspace: workspace
            )
            guard result.succeeded else { throw WorkerClientError.commandFailed(result.standardError) }
            if !copyToClipboard {
                let unit = copies == 1 ? "copy" : "copies"
                notice = "Queued \(copies) \(unit) of \(item.name). See Activity for progress."
            }
            await refresh()
            if !workerRunning {
                do { try await startWorkerOperation() }
                catch {
                    throw WorkerClientError.commandFailed(
                        "Your copies are queued. The worker could not start: \(error.localizedDescription)"
                    )
                }
            }
            await refresh()
            if let clipboardVersion {
                struct Receipt: Decodable { let id: String }
                guard let receipt = try? JSONDecoder().decode(Receipt.self, from: Data(result.standardOutput.utf8)) else {
                    throw WorkerClientError.commandFailed("The copy was queued, but its request ID could not be read.")
                }
                followSavedMediaCopy(receipt.id, pipelineRoot: pipelineRoot, clipboardVersion: clipboardVersion)
            }
        } catch {
            if copyToClipboard {
                savedMediaCopy?.status = .failed(error.localizedDescription)
            } else {
                errorMessage = error.localizedDescription
            }
        }
    }

    private func followSavedMediaCopy(_ id: String, pipelineRoot: URL, clipboardVersion: Int) {
        savedMediaCopyTask = Task { [weak self] in
            while !Task.isCancelled {
                if self?.finishSavedMediaCopy(id, pipelineRoot: pipelineRoot, clipboardVersion: clipboardVersion) != false {
                    return
                }
                try? await Task.sleep(for: .milliseconds(350))
            }
        }
    }

    private func finishSavedMediaCopy(_ id: String, pipelineRoot: URL, clipboardVersion: Int) -> Bool {
        guard pipelineRoot == activePipelineRoot else {
            savedMediaCopy?.status = .failed("The pipeline folder changed while the copy was queued.")
            return true
        }
        activity.refresh(pipelineRoot: pipelineRoot)
        guard let job = activity.job(forRequest: id), job.state == .done || job.state == .failed else {
            // A paused or stopped worker leaves the request in the queue for its next run.
            return false
        }
        let outputRoot = outputFolder(preset: job.preset, workspace: job.workspace)
        guard job.state == .done, job.outputCount == 1, let output = job.outputPaths.first else {
            savedMediaCopy?.status = .failed(job.error ?? "Processing did not create a copy.")
            return true
        }
        let url = outputRoot.appendingPathComponent(output).standardizedFileURL
        guard url.path.hasPrefix(outputRoot.path + "/"), FileManager.default.isReadableFile(atPath: url.path) else {
            savedMediaCopy?.status = .failed("The finished file could not be found.")
            return true
        }
        // A new copy made elsewhere while processing takes precedence over this request.
        savedMediaCopy?.status = clipboard.changeCount == clipboardVersion
            ? writeSavedMediaFile(url)
            : .ready(url, copied: false)
        return true
    }

    func copySavedMediaOutput(for itemID: UUID) {
        guard let copy = savedMediaCopy, copy.itemID == itemID, case .ready(let url, _) = copy.status else { return }
        savedMediaCopy?.status = writeSavedMediaFile(url)
    }

    private func writeSavedMediaFile(_ url: URL) -> SavedMediaCopyStatus {
        guard (try? url.resourceValues(forKeys: [.isRegularFileKey]).isRegularFile) == true else {
            return .failed("The finished file could not be found.")
        }
        clipboard.clearContents()
        return .ready(url, copied: clipboard.writeObjects([url as NSURL]))
    }

    func archive(_ job: JobSnapshot, thenUpload: Bool) async {
        let outputRoot = outputFolder(preset: job.preset, workspace: job.workspace)
        let prefix = outputRoot.path.hasSuffix("/") ? outputRoot.path : outputRoot.path + "/"
        let relativeOutputs = job.outputPaths.compactMap { output -> String? in
            let url = output.hasPrefix("/")
                ? URL(fileURLWithPath: output)
                : outputRoot.appendingPathComponent(output)
            let path = url.standardizedFileURL.path
            guard path.hasPrefix(prefix) else { return nil }
            return String(path.dropFirst(prefix.count))
        }

        do {
            let result = try await worker.archive(
                preset: job.preset,
                workspace: job.workspace,
                outputs: relativeOutputs,
                name: job.files.first.map { URL(fileURLWithPath: $0).deletingPathExtension().lastPathComponent }
            )
            guard result.succeeded else {
                throw WorkerClientError.commandFailed(result.standardError)
            }
            guard let data = result.standardOutput.data(using: .utf8),
                  let archive = try? JSONDecoder().decode(JobArchiveResult.self, from: data) else {
                throw WorkerClientError.commandFailed("The worker returned an unreadable archive result.")
            }
            notice = "Created \(URL(fileURLWithPath: archive.path).lastPathComponent)."
            if thenUpload {
                uploads.upload([
                    StagedFile(
                        url: URL(fileURLWithPath: archive.path),
                        workspace: job.workspace,
                        bytes: archive.bytes
                    ),
                ])
                selection = .uploads
            }
            await refresh()
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func saveConfigurationAndRestart() async {
        // Swift actors are reentrant at every await. Wait for an in-flight start or stop, then
        // retain the operation gate through stop, save, root switch, and replacement startup.
        while workerBusy {
            try? await Task.sleep(for: .milliseconds(100))
        }
        workerBusy = true
        defer { workerBusy = false }
        do {
            // Stop while the on-disk configuration still points at the worker's current root.
            // Otherwise a PipelineRoot edit would write the stop flag into the new, idle root.
            if workerRunning || worker.managedWorkerIsRunning {
                try await stopWorkerOperation()
            }
            try configuration.save()
            activePipelineRoot = configuration.pipelineRoot
            try await startWorkerOperation()
            notice = "Settings saved."
        } catch {
            errorMessage = error.localizedDescription
        }
    }

    func openPipelineFolder() {
        ensureDirectory(activePipelineRoot)
        NSWorkspace.shared.open(activePipelineRoot)
    }

    func openLogs() {
        let url = activePipelineRoot.appendingPathComponent("logs", isDirectory: true)
        ensureDirectory(url)
        NSWorkspace.shared.open(url)
    }

    func openConfig() {
        NSWorkspace.shared.open(configuration.configURL)
    }

    func openLane(_ job: JobSnapshot) {
        openLane(preset: job.preset, workspace: job.workspace)
    }

    func openLane(preset: String, workspace: String) {
        let url = activePipelineRoot
            .appendingPathComponent(workspace, isDirectory: true)
            .appendingPathComponent(preset, isDirectory: true)
        ensureDirectory(url)
        NSWorkspace.shared.open(url)
    }

    func openOutputFolder(preset: String, workspace: String) {
        let url = outputFolder(preset: preset, workspace: workspace)
        do {
            try FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
            if !NSWorkspace.shared.open(url) {
                errorMessage = "Could not open output folder at \(url.path)."
            }
        } catch {
            errorMessage = "Could not open output folder: \(error.localizedDescription)"
        }
    }

    private func outputFolder(preset: String, workspace: String) -> URL {
        activePipelineRoot
            .appendingPathComponent(workspace, isDirectory: true)
            .appendingPathComponent(preset, isDirectory: true)
            .appendingPathComponent("output", isDirectory: true)
            .standardizedFileURL
    }

    private func ensureDirectory(_ url: URL) {
        try? FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
    }

    private func startWorkerOperation() async throws {
        guard !isShuttingDown else { throw CancellationError() }
        let check = try await worker.check()
        guard check.succeeded else {
            throw WorkerClientError.commandFailed(
                check.standardError.isEmpty ? check.standardOutput : check.standardError
            )
        }
        try await worker.start()
        // Publish managed liveness immediately so every reentrant caller sees the new process.
        workerRunning = worker.managedWorkerIsRunning
        try? await Task.sleep(for: .milliseconds(700))
        await refresh()
    }

    private func stopWorkerOperation() async throws {
        try await worker.requestStop()
        let deadline = Date().addingTimeInterval(120)
        while Date() < deadline {
            guard !isShuttingDown else { throw CancellationError() }
            try? await Task.sleep(for: .milliseconds(250))
            await refresh()
            if !workerRunning && !worker.managedWorkerIsRunning {
                return
            }
        }
        throw WorkerClientError.commandFailed(
            "The worker did not stop within two minutes. It was left running."
        )
    }

}
