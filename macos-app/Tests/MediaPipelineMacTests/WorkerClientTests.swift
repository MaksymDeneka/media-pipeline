import Foundation
import XCTest
@testable import MediaPipelineMac

@MainActor
final class WorkerClientTests: XCTestCase {
    func testUploadAwaitsTerminalRecordWithoutTrailingNewline() async throws {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("worker-upload-stream-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }

        let source = root.appendingPathComponent("movie.mp4")
        let config = root.appendingPathComponent("config.ini")
        let workerScript = root.appendingPathComponent("mock-worker")
        try Data([1, 2, 3]).write(to: source)
        try "PipelineRoot = \(root.path)\n".write(
            to: config,
            atomically: true,
            encoding: .utf8
        )
        let terminalJSON = "{\"type\":\"upload.progress\",\"file\":\"movie.mp4\",\"workspace\":\"LC\",\"phase\":\"Done\",\"chunksSent\":1,\"chunks\":1,\"bytesSent\":3,\"bytes\":3,\"error\":null,\"sourceDeleted\":false,\"retainedSourcePath\":null}"
        try "#!/bin/sh\nprintf '%s' '\(terminalJSON)'\n".write(
            to: workerScript,
            atomically: true,
            encoding: .utf8
        )
        try FileManager.default.setAttributes(
            [.posixPermissions: 0o755],
            ofItemAtPath: workerScript.path
        )

        let worker = WorkerClient(configURL: config, executableURL: workerScript)
        for _ in 0..<20 {
            var observed: [UploadProgressEvent] = []
            let terminal = try await worker.upload(file: source, workspace: "LC") {
                observed.append($0)
            }
            XCTAssertEqual(terminal.phase, "Done")
            XCTAssertEqual(observed.last?.phase, "Done")
        }
    }
}
