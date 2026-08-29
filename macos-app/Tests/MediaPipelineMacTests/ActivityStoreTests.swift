import Foundation
import XCTest
@testable import MediaPipelineMac

@MainActor
final class ActivityStoreTests: XCTestCase {
    func testJobSurvivesDailyEventRollover() throws {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("activity-rollover-\(UUID().uuidString)")
        let logs = root.appendingPathComponent("logs", isDirectory: true)
        try FileManager.default.createDirectory(at: logs, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }

        let calendar = Calendar.current
        let today = Date()
        let yesterday = calendar.date(byAdding: .day, value: -1, to: today)!
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyyMMdd"
        try "{\"ts\":\"2026-08-28T23:59:00Z\",\"ev\":\"job.start\",\"jobId\":\"cross-day\",\"preset\":\"long\",\"workspace\":\"LC\",\"files\":[\"movie.mp4\"]}\n"
            .write(to: logs.appendingPathComponent("events-\(formatter.string(from: yesterday)).jsonl"), atomically: true, encoding: .utf8)
        try "{\"ts\":\"2026-08-29T00:02:00Z\",\"ev\":\"job.done\",\"jobId\":\"cross-day\",\"preset\":\"long\",\"workspace\":\"LC\",\"outputs\":3}\n"
            .write(to: logs.appendingPathComponent("events-\(formatter.string(from: today)).jsonl"), atomically: true, encoding: .utf8)

        let store = ActivityStore()
        store.refresh(pipelineRoot: root, now: today)

        XCTAssertEqual(store.jobs.first?.id, "cross-day")
        XCTAssertEqual(store.jobs.first?.state, .done)
        XCTAssertEqual(store.jobs.first?.outputCount, 3)
    }

    func testMissingStartIsReconstructedAndStoppedJobsAreCancelled() throws {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("activity-reconstruct-\(UUID().uuidString)")
        let logs = root.appendingPathComponent("logs", isDirectory: true)
        try FileManager.default.createDirectory(at: logs, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }

        let formatter = DateFormatter()
        formatter.dateFormat = "yyyyMMdd"
        let log = logs.appendingPathComponent("events-\(formatter.string(from: Date())).jsonl")
        try "{\"ts\":\"2026-08-29T00:02:00Z\",\"ev\":\"job.variant\",\"jobId\":\"missing-start\",\"preset\":\"long\",\"workspace\":\"LC\",\"file\":\"movie.mp4\",\"n\":1,\"total\":3}\n"
            .write(to: log, atomically: true, encoding: .utf8)

        let store = ActivityStore()
        store.refresh(pipelineRoot: root)
        store.reconcile(workerRunning: false)

        XCTAssertEqual(store.jobs.first?.state, .cancelled)
        XCTAssertEqual(store.jobs.first?.preset, "long")
    }

    func testActiveSnapshotSurvivesRelaunchAfterMultipleMidnights() throws {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("activity-active-snapshot-\(UUID().uuidString)")
        let status = root.appendingPathComponent("status", isDirectory: true)
        try FileManager.default.createDirectory(at: status, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }

        try "[{\"ts\":\"2026-08-26T23:59:00Z\",\"ev\":\"job.start\",\"jobId\":\"multi-day\",\"preset\":\"long\",\"workspace\":\"LC\",\"files\":[\"movie.mp4\"]}]"
            .write(
                to: status.appendingPathComponent("active-jobs.json"),
                atomically: true,
                encoding: .utf8
            )

        let store = ActivityStore()
        store.refresh(pipelineRoot: root, now: Date())

        XCTAssertEqual(store.jobs.first?.id, "multi-day")
        XCTAssertEqual(store.jobs.first?.state, .running)
        XCTAssertEqual(store.jobs.first?.files, ["movie.mp4"])
    }

    func testEmptyActiveSnapshotDoesNotResurrectYesterdayOrphan() throws {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("activity-stale-snapshot-\(UUID().uuidString)")
        let logs = root.appendingPathComponent("logs", isDirectory: true)
        let status = root.appendingPathComponent("status", isDirectory: true)
        try FileManager.default.createDirectory(at: logs, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: status, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }

        let now = Date()
        let yesterday = Calendar.current.date(byAdding: .day, value: -1, to: now)!
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyyMMdd"
        try "{\"ts\":\"2026-08-28T23:59:00Z\",\"ev\":\"job.start\",\"jobId\":\"orphan\",\"preset\":\"long\",\"workspace\":\"LC\",\"files\":[\"movie.mp4\"]}\n"
            .write(
                to: logs.appendingPathComponent(
                    "events-\(formatter.string(from: yesterday)).jsonl"
                ),
                atomically: true,
                encoding: .utf8
            )
        try "[]".write(
            to: status.appendingPathComponent("active-jobs.json"),
            atomically: true,
            encoding: .utf8
        )

        let store = ActivityStore()
        store.refresh(pipelineRoot: root, now: now)

        XCTAssertFalse(store.jobs.contains { $0.id == "orphan" })
    }

    func testEmptyActiveSnapshotDoesNotResurrectTodayOrphan() throws {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("activity-today-stale-snapshot-\(UUID().uuidString)")
        let logs = root.appendingPathComponent("logs", isDirectory: true)
        let status = root.appendingPathComponent("status", isDirectory: true)
        try FileManager.default.createDirectory(at: logs, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: status, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }

        let now = Date()
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyyMMdd"
        try "{\"ts\":\"2026-08-29T10:00:00Z\",\"ev\":\"job.start\",\"jobId\":\"today-orphan\",\"preset\":\"long\",\"workspace\":\"LC\",\"files\":[\"movie.mp4\"]}\n"
            .write(
                to: logs.appendingPathComponent(
                    "events-\(formatter.string(from: now)).jsonl"
                ),
                atomically: true,
                encoding: .utf8
            )
        try "[]".write(
            to: status.appendingPathComponent("active-jobs.json"),
            atomically: true,
            encoding: .utf8
        )

        let store = ActivityStore()
        store.refresh(pipelineRoot: root, now: now)

        XCTAssertFalse(store.jobs.contains { $0.id == "today-orphan" })
    }

    func testSwitchingPipelineRootClearsPriorActivity() throws {
        let base = FileManager.default.temporaryDirectory
            .appendingPathComponent("activity-root-switch-\(UUID().uuidString)")
        let firstRoot = base.appendingPathComponent("first", isDirectory: true)
        let secondRoot = base.appendingPathComponent("second", isDirectory: true)
        let firstLogs = firstRoot.appendingPathComponent("logs", isDirectory: true)
        try FileManager.default.createDirectory(at: firstLogs, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: secondRoot, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: base) }

        let now = Date()
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyyMMdd"
        try "{\"ts\":\"2026-08-29T10:00:00Z\",\"ev\":\"job.failed\",\"jobId\":\"first-root-failure\",\"preset\":\"long\",\"workspace\":\"LC\",\"error\":\"failed\"}\n"
            .write(
                to: firstLogs.appendingPathComponent(
                    "events-\(formatter.string(from: now)).jsonl"
                ),
                atomically: true,
                encoding: .utf8
            )

        let store = ActivityStore()
        store.refresh(pipelineRoot: firstRoot, now: now)
        XCTAssertEqual(store.jobs.first?.id, "first-root-failure")

        store.refresh(pipelineRoot: secondRoot, now: now)
        XCTAssertTrue(store.jobs.isEmpty)
    }
}
