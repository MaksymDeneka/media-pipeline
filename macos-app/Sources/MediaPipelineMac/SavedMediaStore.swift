import Combine
import Foundation

struct SavedMedia: Codable, Identifiable, Equatable {
    let id: UUID
    let name: String
    let fileExtension: String
    let bytes: Int64
    let savedAt: Date
    var collectionID: UUID? = nil

    static let imageExtensions: Set<String> = ["jpg", "jpeg", "png", "webp", "heic", "heif", "avif", "apng"]
    static let videoExtensions: Set<String> = [
        "mp4", "mov", "m4v", "mkv", "webm", "avi", "flv", "m4a", "3gp", "3g2",
        "mpeg", "mpg", "mpegts", "ts", "ogg", "ogv", "mts",
    ]
    var isImage: Bool { Self.imageExtensions.contains(fileExtension) }
    var isSupported: Bool { isImage || Self.videoExtensions.contains(fileExtension) }
}

struct SavedMediaCollection: Codable, Identifiable, Equatable {
    static let library = SavedMediaCollection(
        id: UUID(uuid: (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1)), name: "Library"
    )

    let id: UUID
    var name: String
}

@MainActor
final class SavedMediaStore: ObservableObject {
    @Published private(set) var items: [SavedMedia] = []
    @Published private(set) var collections: [SavedMediaCollection] = [.library]
    @Published var isBusy = false
    @Published private(set) var loadWarnings: [String] = []
    let directory: URL
    private var collectionsFile: URL { directory.appendingPathComponent(".collections.json") }
    private var canEditCollections = true

    init(directory: URL) throws {
        self.directory = directory
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        if FileManager.default.fileExists(atPath: collectionsFile.path) {
            do {
                let saved = try JSONDecoder().decode([SavedMediaCollection].self, from: Data(contentsOf: collectionsFile))
                let ids = Set(saved.map(\.id))
                guard ids.count == saved.count, ids.contains(SavedMediaCollection.library.id) else {
                    throw WorkerClientError.commandFailed("Invalid Saved Media tabs.")
                }
                collections = saved
            } catch {
                loadWarnings.append("Could not read Saved Media tabs. Media is available in Library.")
                canEditCollections = false
            }
        }
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
        folder(for: item)
            .appendingPathComponent("source", isDirectory: true).appendingPathComponent(item.name)
    }

    func items(in collectionID: UUID) -> [SavedMedia] {
        let knownIDs = Set(collections.map(\.id))
        return items.filter { item in
            let id = item.collectionID.flatMap { knownIDs.contains($0) ? $0 : nil }
            return (id ?? SavedMediaCollection.library.id) == collectionID
        }
    }

    @discardableResult
    func createCollection(named name: String) throws -> SavedMediaCollection {
        let collection = SavedMediaCollection(id: UUID(), name: try validatedName(name, excluding: nil))
        try saveCollections(collections + [collection])
        return collection
    }

    func renameCollection(_ id: UUID, to name: String) throws {
        guard let index = collections.firstIndex(where: { $0.id == id }) else { return }
        var updated = collections
        updated[index].name = try validatedName(name, excluding: id)
        try saveCollections(updated)
    }

    func removeCollection(_ id: UUID) throws {
        guard id != SavedMediaCollection.library.id, items(in: id).isEmpty else {
            throw WorkerClientError.commandFailed("Only empty tabs can be deleted.")
        }
        try saveCollections(collections.filter { $0.id != id })
    }

    func move(_ item: SavedMedia, to collectionID: UUID) throws {
        try requireCollection(collectionID)
        guard let index = items.firstIndex(where: { $0.id == item.id }) else { return }
        var updated = items[index]
        updated.collectionID = collectionID
        try JSONEncoder().encode(updated).write(to: folder(for: updated).appendingPathComponent("entry.json"), options: .atomic)
        items[index] = updated
    }

    @discardableResult
    func save(_ url: URL, in collectionID: UUID = SavedMediaCollection.library.id) async throws -> Bool {
        try requireCollection(collectionID)
        let directory = directory
        let id = UUID()
        let staging = directory.appendingPathComponent(".import-\(id.uuidString)", isDirectory: true)
        defer { try? FileManager.default.removeItem(at: staging) }
        let accessed = url.startAccessingSecurityScopedResource()
        defer { if accessed { url.stopAccessingSecurityScopedResource() } }
        let item = try await Task.detached(priority: .userInitiated) {
            let name = url.lastPathComponent
            let fileExtension = url.pathExtension.lowercased()
            let values = try url.resourceValues(forKeys: [.isRegularFileKey])
            guard values.isRegularFile == true, !name.contains("\\"),
                  SavedMedia.imageExtensions.contains(fileExtension) || SavedMedia.videoExtensions.contains(fileExtension) else {
                throw WorkerClientError.commandFailed("Choose a supported image or video file.")
            }
            let sourceDirectory = staging.appendingPathComponent("source", isDirectory: true)
            try FileManager.default.createDirectory(at: sourceDirectory, withIntermediateDirectories: true)
            let stagedSource = sourceDirectory.appendingPathComponent(name)
            try FileManager.default.copyItem(at: url, to: stagedSource)
            let bytes = try stagedSource.resourceValues(forKeys: [.fileSizeKey]).fileSize ?? 0
            return SavedMedia(id: id, name: name, fileExtension: fileExtension, bytes: Int64(bytes),
                              savedAt: Date(), collectionID: collectionID)
        }.value

        let stagedSource = staging.appendingPathComponent("source").appendingPathComponent(item.name)
        while true {
            let candidates = candidateURLs(bytes: item.bytes)
            let duplicate = await Task.detached(priority: .userInitiated) {
                candidates.contains { FileManager.default.contentsEqual(atPath: $0.path, andPath: stagedSource.path) }
            }.value
            // Another import can finish while the files are being compared.
            guard candidates == candidateURLs(bytes: item.bytes) else { continue }
            if duplicate { return false }
            try requireCollection(collectionID)
            try JSONEncoder().encode(item).write(to: staging.appendingPathComponent("entry.json"), options: .atomic)
            try FileManager.default.moveItem(at: staging, to: folder(for: item))
            items.insert(item, at: 0)
            return true
        }
    }

    func remove(_ item: SavedMedia) throws {
        try FileManager.default.trashItem(at: folder(for: item), resultingItemURL: nil)
        items.removeAll { $0.id == item.id }
    }

    private func folder(for item: SavedMedia) -> URL {
        directory.appendingPathComponent(item.id.uuidString, isDirectory: true)
    }

    private func candidateURLs(bytes: Int64) -> [URL] {
        items.filter { $0.bytes == bytes }.map { sourceURL(for: $0) }
    }

    private func requireCollection(_ id: UUID) throws {
        guard collections.contains(where: { $0.id == id }) else {
            throw WorkerClientError.commandFailed("This Saved Media tab no longer exists.")
        }
    }

    private func validatedName(_ name: String, excluding id: UUID?) throws -> String {
        let name = name.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !name.isEmpty else { throw WorkerClientError.commandFailed("Enter a tab name.") }
        guard !collections.contains(where: { $0.id != id && $0.name.localizedCaseInsensitiveCompare(name) == .orderedSame }) else {
            throw WorkerClientError.commandFailed("A tab with this name already exists.")
        }
        return name
    }

    private func saveCollections(_ updated: [SavedMediaCollection]) throws {
        guard canEditCollections else { throw WorkerClientError.commandFailed("Saved Media tabs could not be loaded.") }
        try JSONEncoder().encode(updated).write(to: collectionsFile, options: .atomic)
        collections = updated
    }
}
