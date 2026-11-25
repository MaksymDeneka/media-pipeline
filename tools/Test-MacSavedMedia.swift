import AppKit
import Foundation
import ImageIO

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
        try testWorkerIdentity()
        try await testThumbnails(root: root)
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

    @MainActor
    static func testWorkerIdentity() throws {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/bin/sleep")
        process.arguments = ["30"]
        try process.run()
        defer {
            if process.isRunning { process.terminate() }
            process.waitUntilExit()
        }
        let formatter = ISO8601DateFormatter()
        formatter.formatOptions = [.withInternetDateTime, .withFractionalSeconds]
        let currentRun = formatter.string(from: Date())
        let oldRun = formatter.string(from: Date().addingTimeInterval(-86_400))
        check(WorkerProcessIdentity.isRunning(
            pid: process.processIdentifier, startedUtc: currentRun, executableURL: process.executableURL
        ), "recognizes a live process even when its status timestamp grows old")
        check(!WorkerProcessIdentity.isRunning(
            pid: process.processIdentifier, startedUtc: currentRun,
            executableURL: URL(fileURLWithPath: "/media-pipeline-worker")
        ), "rejects a stale worker PID reused by an unrelated executable")
        check(!WorkerProcessIdentity.isRunning(
            pid: process.processIdentifier, startedUtc: oldRun, executableURL: process.executableURL
        ), "rejects a stale worker PID reused by the same executable")
        process.terminate()
        process.waitUntilExit()
        check(!WorkerProcessIdentity.isRunning(
            pid: process.processIdentifier, startedUtc: currentRun, executableURL: process.executableURL
        ), "rejects an exited worker")
    }

    @MainActor
    static func testThumbnails(root: URL) async throws {
        let bitmap = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: 320, pixelsHigh: 160,
                                      bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true,
                                      isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
        let red = NSColor(deviceRed: 1, green: 0, blue: 0, alpha: 1)
        for y in 0..<160 { for x in 0..<320 { bitmap.setColor(red, atX: x, y: y) } }
        let imageURL = root.appendingPathComponent("thumbnail.png")
        try bitmap.representation(using: .png, properties: [:])!.write(to: imageURL)
        let loader = MediaThumbnailLoader()
        let image = await loader.image(for: imageURL)
        check(image?.size == NSSize(width: 144, height: 72),
              "real image previews retain aspect ratio at thumbnail size")
        let pixels = NSBitmapImageRep(cgImage: image!.cgImage(forProposedRect: nil, context: nil, hints: nil)!)
        let center = pixels.colorAt(x: 72, y: 36)!.usingColorSpace(.sRGB)!
        let source = CGImageSourceCreateWithURL(imageURL as CFURL, nil)!
        let sourcePixels = NSBitmapImageRep(cgImage: CGImageSourceCreateImageAtIndex(source, 0, nil)!)
        let expected = sourcePixels.colorAt(x: 160, y: 80)!.usingColorSpace(.sRGB)!
        check(abs(center.redComponent - expected.redComponent) < 0.02 &&
              abs(center.greenComponent - expected.greenComponent) < 0.02 &&
              abs(center.blueComponent - expected.blueComponent) < 0.02,
              "image preview preserves decoded source colors")
        let tile = await loader.image(for: imageURL, maxPixelSize: 520)
        check(tile?.size == NSSize(width: 320, height: 160),
              "tile previews do not reuse a lower-resolution list thumbnail")
        try FileManager.default.removeItem(at: imageURL)
        let cached = await loader.image(for: imageURL)
        check(image === cached, "cached previews survive view refresh without reloading")
        let cachedTile = await loader.image(for: imageURL, maxPixelSize: 520)
        check(tile === cachedTile, "tile and list previews have separate cache entries")
        if let path = ProcessInfo.processInfo.environment["MEDIA_THUMBNAIL_TEST_VIDEO"] {
            let frame = await loader.image(for: URL(fileURLWithPath: path))
            check(frame != nil && frame!.size.width > 0 && frame!.size.height <= 144 &&
                  frame!.size.width <= 144 && abs(frame!.size.height / frame!.size.width - 2) < 0.05,
                  "portrait video previews retain aspect ratio within the thumbnail bounds")
            let framePixels = NSBitmapImageRep(cgImage: frame!.cgImage(forProposedRect: nil, context: nil, hints: nil)!)
            let frameCenter = framePixels.colorAt(x: framePixels.pixelsWide / 2, y: framePixels.pixelsHigh / 2)!
                .usingColorSpace(.sRGB)!
            check(frameCenter.blueComponent > frameCenter.redComponent + 0.5 &&
                  frameCenter.blueComponent > frameCenter.greenComponent + 0.5,
                  "video preview displays the blue source frame")
        }
        let missing = await loader.image(for: root.appendingPathComponent("missing.mp4"))
        check(missing == nil, "unavailable media uses the row fallback without failing")
    }

    static func check(_ passed: Bool, _ label: String) {
        guard passed else { fatalError("FAIL \(label)") }
        print("PASS \(label)")
    }
}
