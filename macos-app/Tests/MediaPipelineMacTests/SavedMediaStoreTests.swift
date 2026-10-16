import Foundation
import XCTest
@testable import MediaPipelineMac

@MainActor
final class SavedMediaStoreTests: XCTestCase {
    func testSavedOriginalSurvivesSourceRemovalAndStoreReload() async throws {
        let root = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: root, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }
        let library = root.appendingPathComponent("library")
        let source = root.appendingPathComponent("My photo.JPG")
        let original = Data([1, 2, 3, 4])
        try original.write(to: source)

        let store = try SavedMediaStore(directory: library)
        try await store.save(source)
        // Saving the same filename twice must preserve both independent entries.
        try await store.save(source)
        try FileManager.default.removeItem(at: source)

        let reloaded = try SavedMediaStore(directory: library)
        XCTAssertEqual(reloaded.items.count, 2)
        XCTAssertEqual(Set(reloaded.items.map(\.id)).count, 2)
        for item in reloaded.items {
            XCTAssertEqual(item.name, "My photo.JPG")
            XCTAssertTrue(item.isImage)
            XCTAssertEqual(try Data(contentsOf: reloaded.sourceURL(for: item)), original)
        }
    }

    func testInvalidImportsAndCorruptEntriesDoNotHideGoodMedia() async throws {
        let root = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: root, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }
        let library = root.appendingPathComponent("library")
        let store = try SavedMediaStore(directory: library)
        let unsupported = root.appendingPathComponent("notes.txt")
        try Data([1]).write(to: unsupported)
        do {
            try await store.save(unsupported)
            XCTFail("Unsupported files must not be saved")
        } catch { }
        XCTAssertTrue(store.items.isEmpty)
        let movie = root.appendingPathComponent("clip.mp4")
        try Data([4, 5, 6]).write(to: movie)
        try await store.save(movie)
        let broken = library.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: broken, withIntermediateDirectories: true)
        try Data("broken".utf8).write(to: broken.appendingPathComponent("entry.json"))

        let reloaded = try SavedMediaStore(directory: library)
        XCTAssertEqual(reloaded.items.count, 1)
        XCTAssertFalse(try XCTUnwrap(reloaded.items.first).isImage)
        XCTAssertEqual(reloaded.loadWarnings.count, 1)
    }
}
