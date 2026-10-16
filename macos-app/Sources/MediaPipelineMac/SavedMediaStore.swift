import Combine
import Foundation

struct SavedMedia: Codable, Identifiable, Equatable {
    let id: UUID
    let name: String
    let fileExtension: String
    let bytes: Int64
    let savedAt: Date

    static let imageExtensions: Set<String> = ["jpg", "jpeg", "png", "webp", "heic", "heif", "avif", "apng"]
    static let videoExtensions: Set<String> = [
        "mp4", "mov", "m4v", "mkv", "webm", "avi", "flv", "m4a", "3gp", "3g2",
        "mpeg", "mpg", "mpegts", "ts", "ogg", "ogv", "mts",
    ]
    var isImage: Bool { Self.imageExtensions.contains(fileExtension) }
    var isSupported: Bool { isImage || Self.videoExtensions.contains(fileExtension) }
}

@MainActor
final class SavedMediaStore: ObservableObject {
    @Published private(set) var items: [SavedMedia] = []
    @Published var isBusy = false
    @Published private(set) var loadWarnings: [String] = []
    let directory: URL

    init(directory: URL) throws {
        self.directory = directory
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        for folder in try FileManager.default.contentsOfDirectory(
            at: directory, includingPropertiesForKeys: nil, options: [.skipsHiddenFiles]
        ) {
            do {
                let entry = folder.appendingPathComponent("entry.json")
                let item = try JSONDecoder().decode(SavedMedia.self, from: Data(contentsOf: entry))
                guard item.id.uuidString == folder.lastPathComponent, item.isSupported,
                      item.name == (item.name as NSString).lastPathComponent,
                      !item.name.contains("\\"),
                      URL(fileURLWithPath: item.name).pathExtension.lowercased() == item.fileExtension else {
                    throw WorkerClientError.commandFailed("Invalid saved media entry.")
                }
                items.append(item)
            } catch {
                loadWarnings.append("Could not read saved media \(folder.lastPathComponent): \(error.localizedDescription)")
            }
        }
        items.sort { $0.savedAt > $1.savedAt }
    }

    func sourceURL(for item: SavedMedia) -> URL {
        directory.appendingPathComponent(item.id.uuidString, isDirectory: true)
            .appendingPathComponent("source", isDirectory: true).appendingPathComponent(item.name)
    }

    func save(_ url: URL) async throws {
        let directory = directory
        let accessed = url.startAccessingSecurityScopedResource()
        defer { if accessed { url.stopAccessingSecurityScopedResource() } }
        let item = try await Task.detached(priority: .userInitiated) {
            let values = try url.resourceValues(forKeys: [.isRegularFileKey, .fileSizeKey])
            let item = SavedMedia(
                id: UUID(), name: url.lastPathComponent,
                fileExtension: url.pathExtension.lowercased(),
                bytes: Int64(values.fileSize ?? 0), savedAt: Date()
            )
            guard values.isRegularFile == true, item.isSupported, !item.name.contains("\\") else {
                throw WorkerClientError.commandFailed("Choose a supported image or video file.")
            }
            let staging = directory.appendingPathComponent(".import-\(item.id.uuidString)", isDirectory: true)
            let destination = directory.appendingPathComponent(item.id.uuidString, isDirectory: true)
            try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
            defer { try? FileManager.default.removeItem(at: staging) }
            let sourceDirectory = staging.appendingPathComponent("source", isDirectory: true)
            try FileManager.default.createDirectory(at: sourceDirectory, withIntermediateDirectories: true)
            try FileManager.default.copyItem(at: url, to: sourceDirectory.appendingPathComponent(item.name))
            try JSONEncoder().encode(item).write(to: staging.appendingPathComponent("entry.json"), options: .atomic)
            try FileManager.default.moveItem(at: staging, to: destination)
            return item
        }.value
        items.insert(item, at: 0)
    }

    func remove(_ item: SavedMedia) throws {
        try FileManager.default.trashItem(
            at: directory.appendingPathComponent(item.id.uuidString), resultingItemURL: nil
        )
        items.removeAll { $0.id == item.id }
    }
}
