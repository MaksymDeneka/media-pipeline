import Foundation
import SwiftUI

enum AppResources {
    static func url(forResource name: String, withExtension extensionName: String) -> URL? {
        // A signed application stores resources in Contents/Resources. SwiftPM tests and
        // development builds continue to use the generated module bundle as a fallback.
        Bundle.main.url(forResource: name, withExtension: extensionName)
            ?? Bundle.module.url(forResource: name, withExtension: extensionName)
    }
}

enum ConfigurationStoreError: LocalizedError {
    case invalidPresetName
    case duplicatePreset(String)

    var errorDescription: String? {
        switch self {
        case .invalidPresetName:
            "Enter a preset name without dots, colons, slashes, or backslashes."
        case .duplicatePreset(let name):
            "A preset named \"\(name)\" already exists."
        }
    }
}

struct IniDocument {
    private(set) var lines: [String]

    init(contents: String) {
        lines = contents.components(separatedBy: .newlines)
    }

    func globals() -> [String: String] {
        var values: [String: String] = [:]
        var preset: String?
        for line in lines {
            if let header = Self.header(line) {
                preset = Self.presetName(header)
                continue
            }
            if preset == nil, let pair = Self.pair(line) {
                values[pair.key] = pair.value
            }
        }
        return values
    }

    func presets() -> (order: [String], values: [String: [String: String]]) {
        var order: [String] = []
        var values: [String: [String: String]] = [:]
        var preset: String?
        for line in lines {
            if let header = Self.header(line) {
                preset = Self.presetName(header)
                if let preset, values[preset] == nil {
                    order.append(preset)
                    values[preset] = [:]
                }
                continue
            }
            if let preset, let pair = Self.pair(line) {
                values[preset, default: [:]][pair.key] = pair.value
            }
        }
        return (order, values)
    }

    mutating func setGlobal(_ key: String, value: String, section: String) {
        var currentPreset: String?
        for index in lines.indices {
            if let header = Self.header(lines[index]) {
                currentPreset = Self.presetName(header)
                continue
            }
            if currentPreset == nil,
               let pair = Self.pair(lines[index]),
               pair.key.caseInsensitiveCompare(key) == .orderedSame {
                lines[index] = Self.replacingValue(in: lines[index], with: value)
                return
            }
        }

        if let sectionIndex = lines.firstIndex(where: {
            Self.header($0)?.caseInsensitiveCompare(section) == .orderedSame
        }) {
            let nextHeader = lines[(sectionIndex + 1)...].firstIndex(where: {
                Self.header($0) != nil
            }) ?? lines.endIndex
            lines.insert("\(key) = \(Self.serializedValue(value))", at: nextHeader)
            return
        }

        let firstPreset = lines.firstIndex(where: {
            guard let header = Self.header($0) else { return false }
            return Self.presetName(header) != nil
        }) ?? lines.endIndex
        lines.insert(
            contentsOf: ["", "[\(section)]", "\(key) = \(Self.serializedValue(value))"],
            at: firstPreset
        )
    }

    mutating func setPreset(_ preset: String, key: String, value: String) {
        guard let section = presetRange(preset) else {
            addPreset(preset)
            setPreset(preset, key: key, value: value)
            return
        }

        for index in section.content {
            if let pair = Self.pair(lines[index]),
               pair.key.caseInsensitiveCompare(key) == .orderedSame {
                lines[index] = Self.replacingValue(in: lines[index], with: value)
                return
            }
        }
        lines.insert("\(key) = \(Self.serializedValue(value))", at: section.end)
    }

    mutating func removePresetValue(_ preset: String, key: String) {
        guard let section = presetRange(preset) else { return }
        if let index = section.content.first(where: {
            Self.pair(lines[$0])?.key.caseInsensitiveCompare(key) == .orderedSame
        }) {
            lines.remove(at: index)
        }
    }

    mutating func addPreset(_ name: String) {
        guard presetRange(name) == nil else { return }
        if lines.last?.isEmpty == false {
            lines.append("")
        }
        lines.append(contentsOf: [
            "[preset \(name)]",
            "VideoCopies = 1",
            "ImageCopies = 1",
        ])
    }

    mutating func removePreset(_ name: String) {
        guard let section = presetRange(name) else { return }
        var start = section.header
        while start > lines.startIndex && lines[start - 1].isEmpty {
            start -= 1
        }
        lines.removeSubrange(start..<section.end)
    }

    func serialized() -> String {
        var result = lines.joined(separator: "\n")
        if !result.hasSuffix("\n") {
            result.append("\n")
        }
        return result
    }

    private func presetRange(_ name: String) -> (header: Int, content: Range<Int>, end: Int)? {
        guard let headerIndex = lines.firstIndex(where: {
            guard let header = Self.header($0), let preset = Self.presetName(header) else {
                return false
            }
            return preset.caseInsensitiveCompare(name) == .orderedSame
        }) else {
            return nil
        }

        let end = lines[(headerIndex + 1)...].firstIndex(where: {
            Self.header($0) != nil
        }) ?? lines.endIndex
        return (headerIndex, (headerIndex + 1)..<end, end)
    }

    private static func header(_ line: String) -> String? {
        let trimmed = line.trimmingCharacters(in: .whitespaces)
        guard trimmed.hasPrefix("["), trimmed.hasSuffix("]") else { return nil }
        return String(trimmed.dropFirst().dropLast()).trimmingCharacters(in: .whitespaces)
    }

    private static func presetName(_ header: String) -> String? {
        guard header.lowercased().hasPrefix("preset ") else { return nil }
        let name = String(header.dropFirst(7)).trimmingCharacters(in: .whitespaces)
        return name.isEmpty ? nil : name
    }

    private static func pair(_ line: String) -> (key: String, value: String)? {
        let trimmed = line.trimmingCharacters(in: .whitespaces)
        guard !trimmed.isEmpty, !trimmed.hasPrefix(";"), !trimmed.hasPrefix("#"),
              let equals = trimmed.firstIndex(of: "=") else {
            return nil
        }

        let key = String(trimmed[..<equals]).trimmingCharacters(in: .whitespaces)
        var value = String(trimmed[trimmed.index(after: equals)...])
            .trimmingCharacters(in: .whitespaces)
        if let comment = inlineCommentIndex(in: value) {
            value = String(value[..<comment]).trimmingCharacters(in: .whitespaces)
        }
        if value.count >= 2,
           (value.hasPrefix("\"") && value.hasSuffix("\"") ||
            value.hasPrefix("'") && value.hasSuffix("'")) {
            value = String(value.dropFirst().dropLast())
        }
        return key.isEmpty ? nil : (key, value)
    }

    private static func replacingValue(in line: String, with value: String) -> String {
        guard let equals = line.firstIndex(of: "=") else { return line }
        let valueStart = line.index(after: equals)
        let tail = String(line[valueStart...])
        var commentSuffix = ""
        if let comment = inlineCommentIndex(in: tail) {
            var suffixStart = comment
            while suffixStart > tail.startIndex {
                let previous = tail.index(before: suffixStart)
                guard tail[previous].isWhitespace else { break }
                suffixStart = previous
            }
            commentSuffix = String(tail[suffixStart...])
        }

        return "\(line[...equals]) \(serializedValue(value))\(commentSuffix)"
    }

    private static func inlineCommentIndex(in value: String) -> String.Index? {
        var quote: Character?
        for index in value.indices {
            let character = value[index]
            if let activeQuote = quote {
                if character == activeQuote { quote = nil }
                continue
            }
            if character == "\"" || character == "'" {
                quote = character
                continue
            }
            if (character == ";" || character == "#"), index > value.startIndex {
                let previous = value.index(before: index)
                if value[previous].isWhitespace { return index }
            }
        }
        return nil
    }

    private static func serializedValue(_ value: String) -> String {
        guard inlineCommentIndex(in: value) != nil || value != value.trimmingCharacters(in: .whitespaces) else {
            return value
        }
        if !value.contains("\"") { return "\"\(value)\"" }
        if !value.contains("'") { return "'\(value)'" }
        return value
    }
}

@MainActor
final class ConfigurationStore: ObservableObject {
    @Published private(set) var globals: [String: String] = [:]
    @Published private(set) var presetOrder: [String] = []
    @Published private(set) var presets: [String: [String: String]] = [:]
    @Published var hasUnsavedChanges = false
    @Published var errorMessage: String?

    let configURL: URL
    private var document = IniDocument(contents: "")

    init(configURL: URL) {
        self.configURL = configURL
    }

    var pipelineRoot: URL {
        let configured = value(in: globals, key: "PipelineRoot") ?? "~/MediaPipeline"
        return AppPaths.expandPath(
            AppPaths.isWindowsAbsolutePath(configured) ? "~/MediaPipeline" : configured,
            relativeTo: configURL.deletingLastPathComponent()
        )
    }

    var workspaces: [String] { ["LC", "MD", "YL", "PL", "general"] }

    func load() throws {
        try installDefaultIfNeeded()
        document = IniDocument(contents: try String(contentsOf: configURL, encoding: .utf8))
        reloadPublishedValues()
        hasUnsavedChanges = false
        errorMessage = nil
    }

    func globalValue(_ definition: SettingDefinition) -> String {
        value(in: globals, key: definition.key) ?? definition.defaultValue
    }

    func presetValue(_ definition: SettingDefinition, preset: String) -> String {
        value(in: presets[preset] ?? [:], key: definition.key) ??
            value(in: globals, key: definition.key) ??
            definition.defaultValue
    }

    func hasOverride(_ definition: SettingDefinition, preset: String) -> Bool {
        value(in: presets[preset] ?? [:], key: definition.key) != nil
    }

    func setGlobal(_ definition: SettingDefinition, value: String) {
        document.setGlobal(definition.key, value: value, section: definition.group)
        removeValue(in: &globals, key: definition.key)
        globals[definition.key] = value
        hasUnsavedChanges = true
    }

    func setPreset(_ definition: SettingDefinition, preset: String, value: String) {
        document.setPreset(preset, key: definition.key, value: value)
        var values = presets[preset] ?? [:]
        removeValue(in: &values, key: definition.key)
        values[definition.key] = value
        presets[preset] = values
        hasUnsavedChanges = true
    }

    func resetPreset(_ definition: SettingDefinition, preset: String) {
        document.removePresetValue(preset, key: definition.key)
        var values = presets[preset] ?? [:]
        removeValue(in: &values, key: definition.key)
        presets[preset] = values
        hasUnsavedChanges = true
    }

    func addPreset(_ name: String) throws {
        let trimmed = name.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !trimmed.isEmpty,
              trimmed.rangeOfCharacter(from: CharacterSet(charactersIn: ".:\\/")) == nil else {
            throw ConfigurationStoreError.invalidPresetName
        }
        guard presets.keys.allSatisfy({ $0.caseInsensitiveCompare(trimmed) != .orderedSame }) else {
            throw ConfigurationStoreError.duplicatePreset(trimmed)
        }
        document.addPreset(trimmed)
        reloadPublishedValues()
        hasUnsavedChanges = true
    }

    func removePreset(_ name: String) {
        document.removePreset(name)
        reloadPublishedValues()
        hasUnsavedChanges = true
    }

    func save() throws {
        let temporary = configURL.appendingPathExtension("saving")
        try document.serialized().write(to: temporary, atomically: true, encoding: .utf8)
        _ = try FileManager.default.replaceItemAt(configURL, withItemAt: temporary)
        hasUnsavedChanges = false
        errorMessage = nil
    }

    private func reloadPublishedValues() {
        globals = document.globals()
        let loadedPresets = document.presets()
        presetOrder = loadedPresets.order
        presets = loadedPresets.values
    }

    private func installDefaultIfNeeded() throws {
        guard !FileManager.default.fileExists(atPath: configURL.path) else { return }
        guard let source = AppResources.url(
            forResource: "default-config",
            withExtension: "ini"
        ) else {
            throw CocoaError(.fileNoSuchFile)
        }
        try FileManager.default.copyItem(at: source, to: configURL)
    }

    private func value(in values: [String: String], key: String) -> String? {
        values.first(where: { $0.key.caseInsensitiveCompare(key) == .orderedSame })?.value
    }

    private func removeValue(in values: inout [String: String], key: String) {
        guard let existing = values.keys.first(where: {
            $0.caseInsensitiveCompare(key) == .orderedSame
        }) else { return }
        values.removeValue(forKey: existing)
    }
}

enum SettingCatalog {
    static let globals: [SettingDefinition] = [
        .init(key: "PipelineRoot", label: "Pipeline folder", help: "Input, output, logs, and status live here.", group: "General", defaultValue: "~/MediaPipeline", kind: .text, presetScoped: false),
        .init(key: "Crf", label: "CPU video quality", help: "Lower values produce larger, cleaner H.264 files.", group: "Video", defaultValue: "24", kind: .integer, presetScoped: true),
        .init(key: "X264Preset", label: "CPU encoding speed", help: "Trades encoding time for compression efficiency.", group: "Video", defaultValue: "medium", kind: .choice(["ultrafast", "fast", "medium", "slow", "veryslow"]), presetScoped: false),
        .init(key: "AudioBitrate", label: "Audio bitrate", help: "AAC bitrate used for generated videos.", group: "Video", defaultValue: "128k", kind: .text, presetScoped: true),
        .init(key: "MaxWidth", label: "Maximum width", help: "Wider videos are scaled down without upscaling smaller files.", group: "Video", defaultValue: "1080", kind: .integer, presetScoped: true),
        .init(key: "SizeCapMB", label: "Video size cap", help: "Zero disables the cap.", group: "Video", defaultValue: "8", kind: .decimal, presetScoped: true),
        .init(key: "SizeCapFallbackMaxWidth", label: "Fallback width", help: "Maximum width used when a video exceeds its size cap.", group: "Video", defaultValue: "720", kind: .integer, presetScoped: true),
        .init(key: "MinTrimMs", label: "Minimum trim", help: "Small trim variation in milliseconds.", group: "Video", defaultValue: "15", kind: .integer, presetScoped: true),
        .init(key: "MaxTrimMs", label: "Maximum trim", help: "Upper trim variation in milliseconds.", group: "Video", defaultValue: "95", kind: .integer, presetScoped: true),
        .init(key: "SegmentTargetSeconds", label: "Segment target", help: "Preferred long-video segment duration.", group: "Video", defaultValue: "15", kind: .integer, presetScoped: true),
        .init(key: "SegmentMinSeconds", label: "Minimum segment", help: "Short remainders are folded into the previous segment.", group: "Video", defaultValue: "11", kind: .integer, presetScoped: true),
        .init(key: "PreferVideoToolbox", label: "Use Apple hardware encoding", help: "Use VideoToolbox when its real test encode succeeds.", group: "Video", defaultValue: "true", kind: .boolean, presetScoped: false),
        .init(key: "VideoToolboxBitrateKbps", label: "VideoToolbox bitrate", help: "Default Apple hardware-encoder bitrate in kilobits per second.", group: "Video", defaultValue: "6000", kind: .integer, presetScoped: false),
        .init(key: "MaxrateScale", label: "Size-cap bitrate scale", help: "Advanced ceiling used to land below the size cap on the first encode.", group: "Video", defaultValue: "0.92", kind: .decimal, presetScoped: true),
        .init(key: "ImageProcessingConcurrency", label: "Files at once", help: "Use auto or a whole number up to six.", group: "Images", defaultValue: "auto", kind: .text, presetScoped: false),
        .init(key: "CropMinPermille", label: "Minimum crop", help: "Minimum crop from each image edge, in parts per thousand.", group: "Images", defaultValue: "5", kind: .integer, presetScoped: true),
        .init(key: "CropMaxPermille", label: "Maximum crop", help: "Maximum crop from each image edge, in parts per thousand.", group: "Images", defaultValue: "20", kind: .integer, presetScoped: true),
        .init(key: "JpegQuality", label: "JPEG quality", help: "FFmpeg quality scale. Lower is cleaner.", group: "Images", defaultValue: "4", kind: .integer, presetScoped: true),
        .init(key: "ConvertedJpegQuality", label: "Converted JPEG quality", help: "Quality for sources such as HEIC that require an initial decode.", group: "Images", defaultValue: "12", kind: .integer, presetScoped: true),
        .init(key: "PngCompressionLevel", label: "PNG compression", help: "Compression level from zero to nine for PNG output.", group: "Images", defaultValue: "6", kind: .integer, presetScoped: true),
        .init(key: "StableSeconds", label: "Stable-file delay", help: "A file must stop changing for this long before processing.", group: "Timing", defaultValue: "3", kind: .integer, presetScoped: false),
        .init(key: "TimeoutSeconds", label: "Arrival timeout", help: "Seconds to wait for a single file to finish arriving.", group: "Timing", defaultValue: "600", kind: .integer, presetScoped: false),
        .init(key: "PollSeconds", label: "Polling interval", help: "Delay between queue scans.", group: "Timing", defaultValue: "2", kind: .integer, presetScoped: false),
        .init(key: "ArchiveEnabled", label: "Archive old output", help: "Move old output out of active output folders.", group: "Archive", defaultValue: "true", kind: .boolean, presetScoped: false),
        .init(key: "ArchiveAgeHours", label: "Archive after", help: "Hours before completed output moves to archive.", group: "Archive", defaultValue: "15", kind: .decimal, presetScoped: false),
        .init(key: "ArchiveCheckIntervalMinutes", label: "Archive check", help: "Minutes between archive maintenance passes.", group: "Archive", defaultValue: "30", kind: .integer, presetScoped: false),
        .init(key: "AssetRetentionDays", label: "Retention", help: "Days to keep archives, originals, failures, and staged uploads.", group: "Archive", defaultValue: "5", kind: .integer, presetScoped: false),
        .init(key: "RemoteName", label: "rclone remote", help: "Configured rclone remote used for SFTP parts.", group: "Upload", defaultValue: "heatup-remote", kind: .text, presetScoped: false),
        .init(key: "RemoteSftpPartsRoot", label: "SFTP parts folder", help: "Remote path where resumable chunks are staged.", group: "Upload", defaultValue: "/D:/MediaPipeline/.sync-parts", kind: .text, presetScoped: false),
        .init(key: "RemotePartsRoot", label: "Assembly parts folder", help: "The same staging folder in the remote host's native path syntax.", group: "Upload", defaultValue: "D:\\MediaPipeline\\.sync-parts", kind: .text, presetScoped: false),
        .init(key: "RemoteDirectory", label: "Remote destination", help: "Windows directory where assembled uploads land.", group: "Upload", defaultValue: "D:\\MediaPipeline\\sync", kind: .text, presetScoped: false),
        .init(key: "RemoteSshHost", label: "SSH host", help: "Host used for remote assembly and verification.", group: "Upload", defaultValue: "heatup-remote", kind: .text, presetScoped: false),
        .init(key: "RemoteSshPort", label: "SSH port", help: "Remote SSH port.", group: "Upload", defaultValue: "2222", kind: .integer, presetScoped: false),
        .init(key: "RemoteSshKeyFile", label: "SSH key", help: "Private key used for batch-mode SSH.", group: "Upload", defaultValue: "~/.ssh/heatup_remote_debug_ed25519", kind: .text, presetScoped: false),
        .init(key: "DeleteAfterUpload", label: "Delete after upload", help: "Remove the local file only after remote verification.", group: "Upload", defaultValue: "false", kind: .boolean, presetScoped: false),
        .init(key: "ChunkSizeMB", label: "Chunk size", help: "Size of each resumable upload part.", group: "Upload", defaultValue: "256", kind: .integer, presetScoped: false),
        .init(key: "ParallelChunks", label: "Parallel chunks", help: "Number of parts sent at the same time.", group: "Upload", defaultValue: "4", kind: .integer, presetScoped: false),
    ]

    static let presetOnly: [SettingDefinition] = [
        .init(key: "Enabled", label: "Enabled", help: "Disabled presets keep their folders but are not watched.", group: "Preset", defaultValue: "true", kind: .boolean, presetScoped: true),
        .init(key: "VideoCopies", label: "Video copies", help: "Variants made for each video.", group: "Copies", defaultValue: "1", kind: .integer, presetScoped: true),
        .init(key: "ImageCopies", label: "Image copies", help: "Variants made for each image.", group: "Copies", defaultValue: "1", kind: .integer, presetScoped: true),
        .init(key: "CopiesAlternate", label: "Alternate count", help: "Optional count used on alternating entries.", group: "Copies", defaultValue: "0", kind: .integer, presetScoped: true),
        .init(key: "Grouping", label: "Output grouping", help: "Flat, one folder per source, or a batch containing set folders.", group: "Output", defaultValue: "Flat", kind: .choice(["Flat", "PerSource", "PerSet"]), presetScoped: true),
        .init(key: "SetCount", label: "Set count", help: "Number of set folders for PerSet output.", group: "Output", defaultValue: "1", kind: .integer, presetScoped: true),
        .init(key: "Batch", label: "Batch mode", help: "Process files independently or as one settled folder.", group: "Output", defaultValue: "PerFile", kind: .choice(["PerFile", "PerGroup"]), presetScoped: true),
        .init(key: "Segment", label: "Segment videos", help: "Split long videos before making variants.", group: "Output", defaultValue: "false", kind: .boolean, presetScoped: true),
        .init(key: "Manifest", label: "Write manifest", help: "Add one manifest record per generated variant.", group: "Output", defaultValue: "false", kind: .boolean, presetScoped: true),
        .init(key: "ManifestSchema", label: "Manifest schema", help: "Schema identifier written into generated manifests.", group: "Output", defaultValue: "heatup.assetStoreMediaManifest.v1", kind: .text, presetScoped: true),
        .init(key: "Normalize", label: "Normalize formats", help: "Convert MOV and HEIC inputs before making variants.", group: "Processing", defaultValue: "true", kind: .boolean, presetScoped: true),
        .init(key: "OnFailure", label: "Failure cleanup", help: "Choose what incomplete output the worker removes.", group: "Output", defaultValue: "PreservePartial", kind: .choice(["PreservePartial", "DeleteFiles", "DeleteContainer"]), presetScoped: true),
        .init(key: "Parallel", label: "Parallel mode", help: "Run over files, variants, or sequentially.", group: "Processing", defaultValue: "OverFiles", kind: .choice(["OverFiles", "OverVariants", "Sequential"]), presetScoped: true),
    ]

    static let presetDefinitions = presetOnly + globals.filter(\.presetScoped)

    static var globalGroups: [String] {
        unique(globals.map(\.group))
    }

    static var presetGroups: [String] {
        unique(presetDefinitions.map(\.group))
    }

    private static func unique(_ values: [String]) -> [String] {
        values.reduce(into: []) { result, value in
            if !result.contains(value) { result.append(value) }
        }
    }
}
