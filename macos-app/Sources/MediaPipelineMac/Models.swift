import Foundation

enum AppSection: String, CaseIterable, Identifiable {
    case activity
    case uploads
    case presets
    case settings

    var id: String { rawValue }

    var title: String {
        switch self {
        case .activity: "Activity"
        case .uploads: "Uploads"
        case .presets: "Presets"
        case .settings: "Settings"
        }
    }

    var symbol: String {
        switch self {
        case .activity: "chart.bar.xaxis"
        case .uploads: "arrow.up.circle"
        case .presets: "slider.horizontal.3"
        case .settings: "gearshape"
        }
    }
}

struct WorkerStatus: Decodable {
    let schema: String
    let pid: Int32
    let startedUtc: String
    let updatedUtc: String
    let pipelineRoot: String
    let encoder: String
    let pollSeconds: Int
    let pausedAll: Bool
    let workspaces: [String]
    let presets: [PresetStatus]
    let lanes: [LaneStatus]
}

struct PresetStatus: Decodable, Identifiable {
    let name: String
    let videoCopies: Int
    let imageCopies: Int
    let grouping: String
    let setCount: Int
    let batch: String
    let segment: Bool
    let manifest: Bool
    let sizeCapMB: Double

    var id: String { name }
}

struct LaneStatus: Decodable, Identifiable {
    let preset: String
    let workspace: String
    let queued: Int
    let paused: Bool

    var id: String { "\(preset)/\(workspace)" }
}

struct PipelineEvent: Decodable {
    let timestamp: String
    let name: String
    let jobId: String?
    let preset: String?
    let workspace: String?
    let file: String?
    let files: [String]?
    let index: Int?
    let total: Int?
    let outputs: Int?
    let output: String?
    let bytes: Int64?
    let error: String?

    enum CodingKeys: String, CodingKey {
        case timestamp = "ts"
        case name = "ev"
        case jobId
        case preset
        case workspace
        case file
        case files
        case index = "n"
        case total
        case outputs
        case output
        case bytes
        case error
    }
}

enum JobState: String {
    case running
    case done
    case failed
    case cancelled
}

struct JobSnapshot: Identifiable {
    let id: String
    let preset: String
    let workspace: String
    let files: [String]
    let currentFile: String?
    let completed: Int
    let total: Int
    let outputCount: Int
    let outputPaths: [String]
    let bytes: Int64
    let state: JobState
    let startedAt: Date
    let endedAt: Date?
    let error: String?

    var lane: String { "\(workspace) · \(preset)" }

    var detail: String {
        if let currentFile, state == .running {
            return currentFile
        }

        if files.count == 1 {
            return files[0]
        }

        return "\(files.count) source files"
    }

    var fraction: Double {
        total > 0 ? min(1, Double(completed) / Double(total)) : 0
    }
}

struct UploadProgressEvent: Decodable {
    let type: String
    let file: String
    let workspace: String
    let phase: String
    let chunksSent: Int
    let chunks: Int
    let bytesSent: Int64
    let bytes: Int64
    let error: String?
    let sourceDeleted: Bool
    let retainedSourcePath: String?
}

struct UploadTransfer: Identifiable {
    let id: UUID
    let fileURL: URL
    let workspace: String
    var phase: String
    var chunksSent: Int
    var chunks: Int
    var bytesSent: Int64
    var bytes: Int64
    var error: String?

    var fraction: Double {
        bytes > 0 ? min(1, Double(bytesSent) / Double(bytes)) : 0
    }
}

struct StagedFile: Identifiable, Hashable {
    let url: URL
    let workspace: String
    let bytes: Int64

    var id: String { url.path }
    var name: String { url.lastPathComponent }
}

struct WorkspaceUploads: Identifiable {
    let name: String
    let files: [StagedFile]

    var id: String { name }
    var totalBytes: Int64 { files.reduce(0) { $0 + $1.bytes } }
}

struct CommandResult {
    let exitCode: Int32
    let standardOutput: String
    let standardError: String

    var succeeded: Bool { exitCode == 0 }
}

struct JobArchiveResult: Decodable {
    let path: String
    let fileCount: Int
    let bytes: Int64
    let missing: [String]
}

enum SettingKind: Equatable {
    case text
    case integer
    case decimal
    case boolean
    case choice([String])
}

struct SettingDefinition: Identifiable {
    let key: String
    let label: String
    let help: String
    let group: String
    let defaultValue: String
    let kind: SettingKind
    let presetScoped: Bool

    var id: String { key }
}

enum DateParser {
    private static let fractional: ISO8601DateFormatter = {
        let formatter = ISO8601DateFormatter()
        formatter.formatOptions = [.withInternetDateTime, .withFractionalSeconds]
        return formatter
    }()

    private static let regular = ISO8601DateFormatter()

    static func parse(_ value: String) -> Date? {
        fractional.date(from: value) ?? regular.date(from: value)
    }
}

enum ByteCount {
    static let formatter: ByteCountFormatter = {
        let formatter = ByteCountFormatter()
        formatter.allowedUnits = [.useMB, .useGB, .useTB]
        formatter.countStyle = .file
        return formatter
    }()

    static func string(_ bytes: Int64) -> String {
        formatter.string(fromByteCount: bytes)
    }
}
