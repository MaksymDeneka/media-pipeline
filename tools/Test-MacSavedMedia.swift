import Foundation

// This executable exercises the production store and worker client on Macs that
// have Command Line Tools but no XCTest. It does not open application windows.
@main
struct SavedMediaSelfTest {
    @MainActor
    static func main() async throws {
        let root = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: root, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }
        let library = root.appendingPathComponent("library")
        let source = root.appendingPathComponent("My photo.JPG")
        let original = Data([1, 2, 3, 4])
        try original.write(to: source)
        let store = try SavedMediaStore(directory: library)
        try await store.save(source)
        try await store.save(source)
        try FileManager.default.removeItem(at: source)
        let reloaded = try SavedMediaStore(directory: library)
        check(reloaded.items.count == 2, "library survives reload and source removal")
        check(Set(reloaded.items.map(\.id)).count == 2, "duplicate names have independent saved copies")
        for item in reloaded.items {
            check(item.name == "My photo.JPG" && item.isImage, "preserves source name and media kind")
            check(try Data(contentsOf: reloaded.sourceURL(for: item)) == original, "preserves original bytes")
        }

        let unsupported = root.appendingPathComponent("notes.txt")
        try Data([1]).write(to: unsupported)
        do {
            try await store.save(unsupported)
            fatalError("Unsupported media was accepted")
        } catch { print("PASS rejects unsupported media") }
        let movie = root.appendingPathComponent("clip.mp4")
        try Data([4, 5]).write(to: movie)
        try await store.save(movie)
        check(store.items.first?.isImage == false, "accepts video sources")
        let broken = library.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: broken, withIntermediateDirectories: true)
        try Data("broken".utf8).write(to: broken.appendingPathComponent("entry.json"))
        let recovered = try SavedMediaStore(directory: library)
        check(recovered.items.count == 3 && recovered.loadWarnings.count == 1,
              "corrupt entries do not hide valid media")

        let mock = root.appendingPathComponent("mock-worker")
        try "#!/bin/sh\nprintf '%s\\n' \"$@\"\n".write(to: mock, atomically: true, encoding: .utf8)
        try FileManager.default.setAttributes([.posixPermissions: 0o755], ofItemAtPath: mock.path)
        let config = root.appendingPathComponent("config.ini")
        let worker = WorkerClient(configURL: config, executableURL: mock)
        let saved = reloaded.items[0]
        let result = try await worker.queueMedia(
            file: reloaded.sourceURL(for: saved), copies: 3, preset: "bulk", workspace: "LC"
        )
        check(result.succeeded && result.standardOutput.split(separator: "\n").map(String.init) == [
            "queue-media", "--config", config.path, "--file", reloaded.sourceURL(for: saved).path,
            "--copies", "3", "--preset", "bulk", "--workspace", "LC",
        ], "submits exact count, scope, and paths with spaces to the worker")
        print("SAVED MEDIA SELF-TEST OK")
    }

    static func check(_ passed: Bool, _ label: String) {
        guard passed else { fatalError("FAIL \(label)") }
        print("PASS \(label)")
    }
}
