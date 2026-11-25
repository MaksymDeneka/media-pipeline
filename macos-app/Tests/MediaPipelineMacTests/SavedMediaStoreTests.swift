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
        let first = try await store.save(source)
        let repeated = try await store.save(source)
        XCTAssertTrue(first)
        XCTAssertFalse(repeated)
        let tab = try store.createCollection(named: "Campaign")
        let renamed = root.appendingPathComponent("Other.jpg")
        try original.write(to: renamed)
        let renamedCopy = try await store.save(renamed, in: tab.id)
        XCTAssertFalse(renamedCopy)
        let different = Data([5, 6, 7, 8])
        try different.write(to: source)
        let newContent = try await store.save(source, in: tab.id)
        XCTAssertTrue(newContent)
        try FileManager.default.removeItem(at: source)

        let reloaded = try SavedMediaStore(directory: library)
        XCTAssertEqual(reloaded.items.count, 2)
        let saved = try XCTUnwrap(reloaded.items(in: SavedMediaCollection.library.id).first)
        XCTAssertEqual(saved.name, "My photo.JPG")
        XCTAssertTrue(saved.isImage)
        XCTAssertEqual(try Data(contentsOf: reloaded.sourceURL(for: saved)), original)
        let other = try XCTUnwrap(reloaded.items(in: tab.id).first)
        XCTAssertEqual(other.name, saved.name)
        XCTAssertEqual(try Data(contentsOf: reloaded.sourceURL(for: other)), different)
        XCTAssertThrowsError(try reloaded.removeCollection(tab.id))
        try reloaded.move(other, to: SavedMediaCollection.library.id)
        try reloaded.renameCollection(tab.id, to: "Archive")
        let moved = try SavedMediaStore(directory: library)
        XCTAssertEqual(moved.items(in: SavedMediaCollection.library.id).count, 2)
        XCTAssertTrue(moved.items(in: tab.id).isEmpty)
        XCTAssertEqual(moved.collections.first(where: { $0.id == tab.id })?.name, "Archive")
        try moved.removeCollection(tab.id)
        XCTAssertEqual(try SavedMediaStore(directory: library).collections, [.library])
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
