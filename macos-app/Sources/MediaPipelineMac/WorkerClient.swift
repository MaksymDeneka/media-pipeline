import Darwin
import Foundation

enum WorkerClientError: LocalizedError {
    case executableMissing
    case commandFailed(String)

    var errorDescription: String? {
        switch self {
        case .executableMissing:
            "The native worker is missing from the app bundle."
        case .commandFailed(let message):
            message
        }
    }
}

@MainActor
final class WorkerClient {
    private let configURL: URL
    private let executableURL: URL?
    private var workerProcess: Process?
    private var workerErrorURL: URL?
    private var workerErrorHandle: FileHandle?
    private var lastWorkerError = ""
    private var uploadProcess: Process?
    private var uploadBuffer = ""

    init(configURL: URL, executableURL: URL?) {
        self.configURL = configURL
        self.executableURL = executableURL
    }

    var hasExecutable: Bool { executableURL != nil }
    var managedWorkerIsRunning: Bool { workerProcess?.isRunning == true }

    func check() async throws -> CommandResult {
        try await run(["check", "--config", configURL.path])
    }

    func start() async throws {
        guard let executableURL else { throw WorkerClientError.executableMissing }
        if workerProcess?.isRunning == true { return }

        let process = Process()
        let errorURL = temporaryOutputURL(extension: "worker-errors")
        FileManager.default.createFile(atPath: errorURL.path, contents: nil)
        let errorHandle = try FileHandle(forWritingTo: errorURL)
        process.executableURL = executableURL
        process.arguments = ["run", "--config", configURL.path]
        process.currentDirectoryURL = executableURL.deletingLastPathComponent()
        process.standardOutput = FileHandle.nullDevice
        process.standardError = errorHandle
        process.terminationHandler = { [weak self] finished in
            Task { @MainActor in self?.recordWorkerExit(finished) }
        }
        do {
            try process.run()
        } catch {
            try? errorHandle.close()
            try? FileManager.default.removeItem(at: errorURL)
            throw error
        }
        workerProcess = process
        workerErrorURL = errorURL
        workerErrorHandle = errorHandle
        lastWorkerError = ""

        try? await Task.sleep(for: .milliseconds(750))
        if !process.isRunning {
            recordWorkerExit(process)
            throw WorkerClientError.commandFailed(
                lastWorkerError.isEmpty
                    ? "The worker exited during startup with code \(process.terminationStatus)."
                    : lastWorkerError
            )
        }
    }

    func requestStop() async throws {
        let result = try await run(["stop", "--config", configURL.path])
        if !result.succeeded {
            throw WorkerClientError.commandFailed(result.standardError)
        }
    }

    func setPaused(_ paused: Bool, preset: String? = nil, workspace: String? = nil) async throws {
        var arguments = [paused ? "pause" : "resume", "--config", configURL.path]
        if let preset {
            arguments.append(contentsOf: ["--preset", preset])
        }
        if let workspace {
            arguments.append(contentsOf: ["--workspace", workspace])
        }
        let result = try await run(arguments)
        if !result.succeeded {
            throw WorkerClientError.commandFailed(result.standardError)
        }
    }

    func requeue(preset: String, workspace: String) async throws -> CommandResult {
        try await run([
            "requeue",
            "--config", configURL.path,
            "--preset", preset,
            "--workspace", workspace,
        ])
    }

    func archive(
        preset: String,
        workspace: String,
        outputs: [String],
        name: String?
    ) async throws -> CommandResult {
        var arguments = [
            "archive",
            "--config", configURL.path,
            "--preset", preset,
            "--workspace", workspace,
        ]
        if let name, !name.isEmpty {
            arguments.append(contentsOf: ["--name", name])
        }
        for output in outputs {
            arguments.append(contentsOf: ["--output", output])
        }
        return try await run(arguments)
    }

    func recompress(preset: String? = nil, workspace: String? = nil) async throws -> CommandResult {
        var arguments = ["recompress", "--config", configURL.path]
        if let preset { arguments.append(contentsOf: ["--preset", preset]) }
        if let workspace { arguments.append(contentsOf: ["--workspace", workspace]) }
        return try await run(arguments)
    }

    func upload(
        file: URL,
        workspace: String,
        onProgress: @escaping @MainActor (UploadProgressEvent) -> Void
    ) async throws -> UploadProgressEvent {
        guard let executableURL else { throw WorkerClientError.executableMissing }
        guard uploadProcess == nil else {
            throw WorkerClientError.commandFailed("Another upload is already running.")
        }

        let process = Process()
        let output = Pipe()
        let errors = temporaryOutputURL(extension: "upload-errors")
        FileManager.default.createFile(atPath: errors.path, contents: nil)
        let errorHandle = try FileHandle(forWritingTo: errors)

        process.executableURL = executableURL
        process.arguments = [
            "upload",
            "--config", configURL.path,
            "--file", file.path,
            "--workspace", workspace,
            "--json",
        ]
        process.currentDirectoryURL = executableURL.deletingLastPathComponent()
        process.standardOutput = output
        process.standardError = errorHandle
        uploadBuffer = ""
        uploadProcess = process

        let (outputStream, streamContinuation) = AsyncStream<Data>.makeStream()
        let outputLock = NSLock()
        let progressTask: Task<UploadProgressEvent?, Never> = Task { @MainActor [weak self] in
            guard let self else { return nil }
            var lastProgress: UploadProgressEvent?
            for await data in outputStream {
                guard let text = String(data: data, encoding: .utf8) else { continue }
                for progress in consumeUploadOutput(text) {
                    onProgress(progress)
                    lastProgress = progress
                }
            }
            for progress in flushUploadBuffer() {
                onProgress(progress)
                lastProgress = progress
            }
            return lastProgress
        }

        output.fileHandleForReading.readabilityHandler = { handle in
            outputLock.withLock {
                let data = handle.availableData
                if !data.isEmpty {
                    streamContinuation.yield(data)
                }
            }
        }

        do {
            let exitCode = try await launchAndWait(process)
            output.fileHandleForReading.readabilityHandler = nil
            outputLock.withLock {
                let remaining = output.fileHandleForReading.readDataToEndOfFile()
                if !remaining.isEmpty {
                    streamContinuation.yield(remaining)
                }
                streamContinuation.finish()
            }
            let finalProgress = await progressTask.value
            try? errorHandle.close()
            uploadProcess = nil

            if exitCode != 0 && exitCode != 130 {
                let detail = (try? String(contentsOf: errors, encoding: .utf8)) ?? ""
                try? FileManager.default.removeItem(at: errors)
                throw WorkerClientError.commandFailed(
                    detail.isEmpty ? "Upload failed with exit code \(exitCode)." : detail
                )
            }
            try? FileManager.default.removeItem(at: errors)
            guard let finalProgress else {
                throw WorkerClientError.commandFailed(
                    "The upload worker exited without a terminal progress record."
                )
            }
            if exitCode == 0 &&
                finalProgress.phase.caseInsensitiveCompare("done") != .orderedSame {
                throw WorkerClientError.commandFailed(
                    "The upload worker exited successfully without confirming completion."
                )
            }
            if exitCode == 130 &&
                finalProgress.phase.caseInsensitiveCompare("cancelled") != .orderedSame {
                throw WorkerClientError.commandFailed(
                    "The cancelled upload exited without confirming cancellation."
                )
            }
            return finalProgress
        } catch {
            output.fileHandleForReading.readabilityHandler = nil
            outputLock.withLock {
                streamContinuation.finish()
            }
            _ = await progressTask.value
            try? errorHandle.close()
            uploadProcess = nil
            try? FileManager.default.removeItem(at: errors)
            throw error
        }
    }

    func cancelUpload() {
        guard let uploadProcess, uploadProcess.isRunning else { return }
        Darwin.kill(uploadProcess.processIdentifier, SIGINT)
    }

    func run(_ arguments: [String]) async throws -> CommandResult {
        guard let executableURL else { throw WorkerClientError.executableMissing }
        let outputURL = temporaryOutputURL(extension: "stdout")
        let errorURL = temporaryOutputURL(extension: "stderr")
        FileManager.default.createFile(atPath: outputURL.path, contents: nil)
        FileManager.default.createFile(atPath: errorURL.path, contents: nil)
        let outputHandle = try FileHandle(forWritingTo: outputURL)
        let errorHandle = try FileHandle(forWritingTo: errorURL)
        defer {
            try? outputHandle.close()
            try? errorHandle.close()
            try? FileManager.default.removeItem(at: outputURL)
            try? FileManager.default.removeItem(at: errorURL)
        }

        let process = Process()
        process.executableURL = executableURL
        process.arguments = arguments
        process.currentDirectoryURL = executableURL.deletingLastPathComponent()
        process.standardOutput = outputHandle
        process.standardError = errorHandle
        let exitCode = try await launchAndWait(process)
        try outputHandle.synchronize()
        try errorHandle.synchronize()

        return CommandResult(
            exitCode: exitCode,
            standardOutput: (try? String(contentsOf: outputURL, encoding: .utf8)) ?? "",
            standardError: (try? String(contentsOf: errorURL, encoding: .utf8)) ?? ""
        )
    }

    private func consumeUploadOutput(_ text: String) -> [UploadProgressEvent] {
        uploadBuffer.append(text)
        var lines = uploadBuffer.components(separatedBy: .newlines)
        uploadBuffer = lines.removeLast()
        return lines.compactMap(decodeProgress)
    }

    private func flushUploadBuffer() -> [UploadProgressEvent] {
        defer { uploadBuffer = "" }
        if !uploadBuffer.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty {
            return [uploadBuffer].compactMap(decodeProgress)
        }
        return []
    }

    private func decodeProgress(_ line: String) -> UploadProgressEvent? {
        guard let data = line.data(using: .utf8),
              let progress = try? JSONDecoder().decode(UploadProgressEvent.self, from: data) else {
            return nil
        }
        return progress
    }

    private func temporaryOutputURL(extension pathExtension: String) -> URL {
        FileManager.default.temporaryDirectory
            .appendingPathComponent("media-pipeline-\(UUID().uuidString)")
            .appendingPathExtension(pathExtension)
    }

    private func launchAndWait(_ process: Process) async throws -> Int32 {
        try await withCheckedThrowingContinuation { continuation in
            process.terminationHandler = { finished in
                continuation.resume(returning: finished.terminationStatus)
            }
            do {
                try process.run()
            } catch {
                process.terminationHandler = nil
                continuation.resume(throwing: error)
            }
        }
    }

    private func recordWorkerExit(_ process: Process) {
        guard workerProcess === process || workerProcess == nil else { return }
        try? workerErrorHandle?.close()
        workerErrorHandle = nil
        if let workerErrorURL {
            lastWorkerError = ((try? String(contentsOf: workerErrorURL, encoding: .utf8)) ?? "")
                .trimmingCharacters(in: .whitespacesAndNewlines)
            try? FileManager.default.removeItem(at: workerErrorURL)
        }
        self.workerErrorURL = nil
        if workerProcess === process { workerProcess = nil }
    }
}
