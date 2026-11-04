import Darwin
import Foundation

@main
struct QuitSelfTest {
    @MainActor
    static func main() async throws {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("media-pipeline-quit-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }
        let script = root.appendingPathComponent("mock-worker")
        try """
        #!/bin/sh
        cd "$(dirname "$0")"
        trap '' INT TERM
        echo $$ > "$1.pid"
        /bin/sh -c 'trap "" INT TERM; while :; do /bin/sleep 1; done' &
        echo $! > "$1.child.pid"
        while :; do /bin/sleep 1; done
        """.write(to: script, atomically: true, encoding: .utf8)
        try FileManager.default.setAttributes([.posixPermissions: 0o755], ofItemAtPath: script.path)
        let config = root.appendingPathComponent("config.ini")
        let worker = WorkerClient(configURL: config, executableURL: script)
        try await worker.start()
        let upload = Task {
            try? await worker.upload(file: root.appendingPathComponent("video.mp4"), workspace: "LC") { _ in }
        }
        let command = Task { try? await worker.run(["archive"]) }
        let deadline = Date().addingTimeInterval(5)
        while !FileManager.default.fileExists(atPath: root.appendingPathComponent("archive.child.pid").path) ||
            !FileManager.default.fileExists(atPath: root.appendingPathComponent("upload.child.pid").path) {
            precondition(Date() < deadline, "Concurrent commands did not start")
            try await Task.sleep(for: .milliseconds(20))
        }
        let pids = try ["run", "run.child", "upload", "upload.child", "archive", "archive.child"].map {
            Int32(try String(contentsOf: root.appendingPathComponent("\($0).pid"), encoding: .utf8)
                .trimmingCharacters(in: .whitespacesAndNewlines))!
        }
        defer { for pid in pids { Darwin.kill(pid, SIGKILL) } }
        await worker.shutdown()
        _ = await upload.value
        _ = await command.value
        precondition(!worker.managedWorkerIsRunning)
        for pid in pids {
            precondition(Darwin.kill(pid, 0) != 0, "Process \(pid) survived quitting")
        }
        print("PASS quit stops the worker, upload, commands, and stubborn child processes")
        await worker.shutdown()
        do {
            try await worker.start()
            fatalError("Worker restarted during shutdown")
        } catch is CancellationError {}
        do {
            _ = try await worker.run(["archive"])
            fatalError("Command launched during shutdown")
        } catch is CancellationError {}
        do {
            _ = try await worker.upload(file: root, workspace: "LC") { _ in }
            fatalError("Upload launched during shutdown")
        } catch is CancellationError {}
        print("PASS repeated quit is safe and new background work is rejected")

        // A stale status PID must never cause an unrelated process to be killed.
        let unrelated = Process()
        unrelated.executableURL = URL(fileURLWithPath: "/bin/sleep")
        unrelated.arguments = ["30"]
        try unrelated.run()
        defer { unrelated.terminate(); unrelated.waitUntilExit() }
        let recovered = WorkerClient(configURL: config, executableURL: script)
        await recovered.shutdown(existingWorkerPID: unrelated.processIdentifier)
        precondition(unrelated.isRunning)
        print("PASS stale status PID leaves unrelated processes running")

        print("QUIT SELF-TEST OK")
    }
}
