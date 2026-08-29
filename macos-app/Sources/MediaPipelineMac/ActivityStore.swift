import Foundation

@MainActor
final class ActivityStore: ObservableObject {
    @Published private(set) var jobs: [JobSnapshot] = []

    private struct MutableJob {
        let id: String
        var preset: String
        var workspace: String
        var files: [String]
        var currentFile: String?
        var completed: Int
        var total: Int
        var outputCount: Int
        var outputPaths: [String]
        var bytes: Int64
        var state: JobState
        var startedAt: Date
        var endedAt: Date?
        var error: String?

        var snapshot: JobSnapshot {
            JobSnapshot(
                id: id,
                preset: preset,
                workspace: workspace,
                files: files,
                currentFile: currentFile,
                completed: completed,
                total: total,
                outputCount: outputCount,
                outputPaths: outputPaths,
                bytes: bytes,
                state: state,
                startedAt: startedAt,
                endedAt: endedAt,
                error: error
            )
        }
    }

    private var mutableJobs: [String: MutableJob] = [:]
    private var currentPipelineRoot: URL?
    private var currentEventURL: URL?
    private var offset: UInt64 = 0
    private var partialLine = ""

    var running: [JobSnapshot] { jobs.filter { $0.state == .running } }
    var failures: [JobSnapshot] { jobs.filter { $0.state == .failed } }
    var recent: [JobSnapshot] { jobs.filter { $0.state != .running } }

    func refresh(pipelineRoot: URL, now: Date = Date()) {
        let normalizedRoot = pipelineRoot.standardizedFileURL
        if normalizedRoot != currentPipelineRoot {
            currentPipelineRoot = normalizedRoot
            currentEventURL = nil
            offset = 0
            partialLine = ""
            mutableJobs.removeAll()
            jobs = []
        }

        let formatter = DateFormatter()
        formatter.dateFormat = "yyyyMMdd"
        let eventURL = pipelineRoot
            .appendingPathComponent("logs", isDirectory: true)
            .appendingPathComponent("events-\(formatter.string(from: now)).jsonl")

        var requiresActiveOverlay = currentEventURL == nil
        if requiresActiveOverlay {
            let yesterday = Calendar.current.date(byAdding: .day, value: -1, to: now) ?? now
            let previousURL = pipelineRoot
                .appendingPathComponent("logs", isDirectory: true)
                .appendingPathComponent("events-\(formatter.string(from: yesterday)).jsonl")
            readWholeFile(previousURL)
        }
        if eventURL != currentEventURL {
            currentEventURL = eventURL
            offset = 0
            partialLine = ""
        }
        guard FileManager.default.fileExists(atPath: eventURL.path),
              let handle = try? FileHandle(forReadingFrom: eventURL) else {
            if requiresActiveOverlay {
                readActiveJobs(pipelineRoot, authoritative: true)
            }
            return
        }
        defer { try? handle.close() }

        let fileSize = (try? handle.seekToEnd()) ?? 0
        if fileSize < offset {
            offset = 0
            partialLine = ""
            mutableJobs.removeAll()
            requiresActiveOverlay = true
        }
        try? handle.seek(toOffset: offset)
        guard let data = try? handle.readToEnd(), !data.isEmpty else {
            if requiresActiveOverlay {
                readActiveJobs(pipelineRoot, authoritative: true)
            }
            return
        }
        offset += UInt64(data.count)
        guard let text = String(data: data, encoding: .utf8) else {
            if requiresActiveOverlay {
                readActiveJobs(pipelineRoot, authoritative: true)
            }
            return
        }

        partialLine.append(text)
        var lines = partialLine.components(separatedBy: .newlines)
        partialLine = lines.removeLast()
        for line in lines where !line.isEmpty {
            guard let eventData = line.data(using: .utf8),
                  let event = try? JSONDecoder().decode(PipelineEvent.self, from: eventData) else {
                continue
            }
            apply(event)
        }
        if requiresActiveOverlay {
            readActiveJobs(pipelineRoot, authoritative: true)
        } else {
            publish()
        }
    }

    func dismissFailure(id: String) {
        mutableJobs.removeValue(forKey: id)
        publish()
    }

    func reconcile(workerRunning: Bool) {
        guard !workerRunning else { return }
        var changed = false
        for id in mutableJobs.keys where mutableJobs[id]?.state == .running {
            mutableJobs[id]?.state = .cancelled
            mutableJobs[id]?.endedAt = Date()
            mutableJobs[id]?.error = "The worker stopped before this job completed."
            changed = true
        }
        if changed { publish() }
    }

    private func apply(_ event: PipelineEvent) {
        let timestamp = DateParser.parse(event.timestamp) ?? Date()
        if event.name == "watcher.start" || event.name == "watcher.stop" {
            for id in mutableJobs.keys where mutableJobs[id]?.state == .running {
                mutableJobs[id]?.state = .cancelled
                mutableJobs[id]?.endedAt = timestamp
                mutableJobs[id]?.error = "The worker stopped before this job completed."
            }
            return
        }

        guard let id = event.jobId else { return }
        switch event.name {
        case "job.start":
            mutableJobs[id] = MutableJob(
                id: id,
                preset: event.preset ?? "unknown",
                workspace: event.workspace ?? "general",
                files: event.files ?? [],
                currentFile: nil,
                completed: 0,
                total: 0,
                outputCount: 0,
                outputPaths: [],
                bytes: event.bytes ?? 0,
                state: .running,
                startedAt: timestamp,
                endedAt: nil,
                error: nil
            )
        case "job.variant":
            var job = mutableJobs[id] ?? reconstructedJob(from: event, id: id, timestamp: timestamp)
            job.currentFile = event.file
            job.completed = event.index ?? job.completed
            job.total = event.total ?? job.total
            if let output = event.output { job.outputPaths.append(output) }
            mutableJobs[id] = job
        case "job.done":
            var job = mutableJobs[id] ?? reconstructedJob(from: event, id: id, timestamp: timestamp)
            job.state = .done
            job.outputCount = event.outputs ?? job.outputPaths.count
            job.endedAt = timestamp
            mutableJobs[id] = job
        case "job.failed":
            var job = mutableJobs[id] ?? reconstructedJob(from: event, id: id, timestamp: timestamp)
            job.state = .failed
            job.error = event.error
            job.endedAt = timestamp
            mutableJobs[id] = job
        case "job.cancelled":
            var job = mutableJobs[id] ?? reconstructedJob(from: event, id: id, timestamp: timestamp)
            job.state = .cancelled
            job.error = event.error
            job.endedAt = timestamp
            mutableJobs[id] = job
        default:
            break
        }
    }

    private func publish() {
        let terminal = mutableJobs.values
            .filter { $0.state != .running }
            .sorted { ($0.endedAt ?? $0.startedAt) > ($1.endedAt ?? $1.startedAt) }
        if terminal.count > 500 {
            for job in terminal.dropFirst(500) { mutableJobs.removeValue(forKey: job.id) }
        }
        jobs = mutableJobs.values
            .map(\.snapshot)
            .sorted { left, right in
                let leftDate = left.endedAt ?? left.startedAt
                let rightDate = right.endedAt ?? right.startedAt
                return leftDate > rightDate
            }
            .prefix(150)
            .map { $0 }
    }

    private func reconstructedJob(from event: PipelineEvent, id: String, timestamp: Date) -> MutableJob {
        MutableJob(
            id: id,
            preset: event.preset ?? "unknown",
            workspace: event.workspace ?? "general",
            files: event.files ?? event.file.map { [$0] } ?? [],
            currentFile: event.file,
            completed: event.index ?? 0,
            total: event.total ?? 0,
            outputCount: 0,
            outputPaths: [],
            bytes: event.bytes ?? 0,
            state: .running,
            startedAt: timestamp,
            endedAt: nil,
            error: nil
        )
    }

    private func readWholeFile(_ url: URL) {
        guard let data = try? Data(contentsOf: url),
              let text = String(data: data, encoding: .utf8) else { return }
        for line in text.components(separatedBy: .newlines) where !line.isEmpty {
            guard let eventData = line.data(using: .utf8),
                  let event = try? JSONDecoder().decode(PipelineEvent.self, from: eventData) else {
                continue
            }
            apply(event)
        }
        publish()
    }

    private func readActiveJobs(_ pipelineRoot: URL, authoritative: Bool) {
        let url = pipelineRoot
            .appendingPathComponent("status", isDirectory: true)
            .appendingPathComponent("active-jobs.json")
        guard let data = try? Data(contentsOf: url),
              let active = try? JSONDecoder().decode([PipelineEvent].self, from: data) else {
            publish()
            return
        }
        if authoritative {
            let activeIDs = Set(active.compactMap(\.jobId))
            for id in mutableJobs.keys where
                mutableJobs[id]?.state == .running && !activeIDs.contains(id) {
                mutableJobs.removeValue(forKey: id)
            }
        }
        for event in active { apply(event) }
        publish()
    }
}
