import Foundation

struct AppPaths {
    let applicationSupport: URL
    let configFile: URL

    static func current(fileManager: FileManager = .default) throws -> AppPaths {
        let supportRoot = try fileManager.url(
            for: .applicationSupportDirectory,
            in: .userDomainMask,
            appropriateFor: nil,
            create: true
        )
        let directory = supportRoot.appendingPathComponent("Media Pipeline", isDirectory: true)
        try fileManager.createDirectory(at: directory, withIntermediateDirectories: true)
        return AppPaths(
            applicationSupport: directory,
            configFile: directory.appendingPathComponent("config.ini")
        )
    }

    static func expandPath(_ value: String, relativeTo base: URL? = nil) -> URL {
        let expanded = NSString(string: value)
            .expandingTildeInPath
            .replacingOccurrences(of: "\\", with: "/")
        if NSString(string: expanded).isAbsolutePath || base == nil {
            return URL(fileURLWithPath: expanded, isDirectory: true).standardizedFileURL
        }

        return base!
            .appendingPathComponent(expanded, isDirectory: true)
            .standardizedFileURL
    }

    static func isWindowsAbsolutePath(_ value: String) -> Bool {
        value.range(of: "^[A-Za-z]:[\\\\/]", options: .regularExpression) != nil ||
            value.hasPrefix("\\\\")
    }

    static func workerExecutable(fileManager: FileManager = .default) -> URL? {
        let bundled = Bundle.main.bundleURL
            .appendingPathComponent("Contents/Helpers/media-pipeline-worker")
        if fileManager.isExecutableFile(atPath: bundled.path) {
            return bundled
        }

        var directory = URL(fileURLWithPath: fileManager.currentDirectoryPath, isDirectory: true)
        for _ in 0..<8 {
            let candidate = directory
                .appendingPathComponent("artifacts/native-worker/osx-arm64/media-pipeline-worker")
            if fileManager.isExecutableFile(atPath: candidate.path) {
                return candidate
            }
            directory.deleteLastPathComponent()
        }

        return nil
    }
}
