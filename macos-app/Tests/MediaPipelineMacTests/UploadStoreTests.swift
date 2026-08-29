import Foundation
import XCTest
@testable import MediaPipelineMac

final class UploadStoreTests: XCTestCase {
    func testTrashFailureRecoveryNeverReplacesNewProducerOutput() throws {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("upload-trash-recovery-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }

        let original = root.appendingPathComponent("movie.mp4")
        let claim = root.appendingPathComponent(".movie.mp4.verified.claim")
        try Data("new producer output".utf8).write(to: original)
        try Data("verified upload source".utf8).write(to: claim)

        let restored = UploadSourceRecovery.restoreRetainedSource(
            claim,
            preferredURL: original
        )

        XCTAssertEqual(try Data(contentsOf: original), Data("new producer output".utf8))
        XCTAssertEqual(restored?.lastPathComponent, "movie.upload-retry-1.mp4")
        XCTAssertEqual(
            try Data(contentsOf: XCTUnwrap(restored)),
            Data("verified upload source".utf8)
        )
    }
}
