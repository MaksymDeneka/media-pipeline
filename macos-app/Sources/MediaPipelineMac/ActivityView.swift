import SwiftUI

struct ActivityView: View {
    @EnvironmentObject private var model: AppModel

    var body: some View {
        ScrollView {
            LazyVStack(alignment: .leading, spacing: 0) {
                summary
                if !model.activity.failures.isEmpty {
                    sectionHeader("Failures", count: model.activity.failures.count)
                    ForEach(model.activity.failures) { job in
                        FailureJobRow(job: job)
                        Divider()
                    }
                }
                sectionHeader("In progress", count: model.activity.running.count)
                if model.activity.running.isEmpty {
                    EmptySection(text: "No jobs are running.")
                } else {
                    ForEach(model.activity.running) { job in
                        RunningJobRow(job: job)
                        Divider()
                    }
                }
                sectionHeader("Queued lanes", count: queuedLanes.count)
                if queuedLanes.isEmpty {
                    EmptySection(text: "No files are waiting.")
                } else {
                    ForEach(queuedLanes) { lane in
                        LaneRow(lane: lane)
                        Divider()
                    }
                }
                sectionHeader("Recent", count: model.activity.recent.count)
                ForEach(model.activity.recent.prefix(40)) { job in
                    CompletedJobRow(job: job)
                    Divider()
                }
            }
            .padding(.bottom, 28)
        }
        .navigationTitle("Activity")
    }

    private var queuedLanes: [LaneStatus] {
        (model.status?.lanes ?? []).filter { $0.queued > 0 }
    }

    private var summary: some View {
        HStack(spacing: 0) {
            Metric(value: "\(model.activity.running.count)", label: "Working")
            Divider()
            Metric(value: "\(model.queuedCount)", label: "Queued")
            Divider()
            Metric(
                value: "\(model.activity.recent.filter { $0.state == .done }.count)",
                label: "Finished today",
                color: .green
            )
            Divider()
            Metric(value: model.encoderName, label: "Encoder")
        }
        .frame(height: 84)
        .overlay(alignment: .bottom) { Divider() }
    }

    private func sectionHeader(_ title: String, count: Int) -> some View {
        HStack {
            Text(title.uppercased())
                .font(.caption2.weight(.semibold))
                .foregroundStyle(.secondary)
            Spacer()
            Text("\(count)")
                .font(.caption2.monospacedDigit())
                .foregroundStyle(.tertiary)
        }
        .padding(.horizontal, 22)
        .padding(.top, 22)
        .padding(.bottom, 7)
    }
}

private struct Metric: View {
    let value: String
    let label: String
    var color: Color = .primary

    var body: some View {
        VStack(alignment: .leading, spacing: 5) {
            Text(value)
                .font(.title3.weight(.semibold))
                .foregroundStyle(color)
                .lineLimit(1)
                .minimumScaleFactor(0.75)
            Text(label.uppercased())
                .font(.caption2.weight(.semibold))
                .foregroundStyle(.secondary)
        }
        .frame(maxWidth: .infinity, alignment: .leading)
        .padding(.horizontal, 20)
    }
}

private struct EmptySection: View {
    let text: String

    var body: some View {
        Text(text)
            .font(.callout)
            .foregroundStyle(.secondary)
            .padding(.horizontal, 22)
            .padding(.vertical, 13)
    }
}

private struct RunningJobRow: View {
    @EnvironmentObject private var model: AppModel
    let job: JobSnapshot

    var body: some View {
        HStack(spacing: 12) {
            Circle().fill(.primary).frame(width: 7, height: 7)
            Text(job.lane).font(.callout.weight(.semibold)).frame(width: 150, alignment: .leading)
            VStack(alignment: .leading, spacing: 7) {
                Text(job.detail).font(.caption.monospaced()).lineLimit(1)
                ProgressView(value: job.fraction).progressViewStyle(.linear).tint(.primary)
            }
            Text(job.total > 0 ? "\(job.completed) / \(job.total)" : "Starting")
                .font(.caption.monospacedDigit())
                .foregroundStyle(.secondary)
                .frame(width: 76, alignment: .trailing)
            Button("Open") { model.openLane(job) }
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 10)
    }
}

private struct FailureJobRow: View {
    @EnvironmentObject private var model: AppModel
    let job: JobSnapshot

    var body: some View {
        HStack(spacing: 12) {
            Circle().fill(Color.red).frame(width: 7, height: 7)
            Text(job.lane).font(.callout.weight(.semibold)).frame(width: 150, alignment: .leading)
            VStack(alignment: .leading, spacing: 3) {
                Text(job.detail).font(.caption.monospaced()).lineLimit(1)
                Text(job.error ?? "Processing failed")
                    .font(.caption)
                    .foregroundStyle(.red)
                    .lineLimit(2)
            }
            Spacer()
            Button("Open Folder") { model.openLane(job) }
            Button("Requeue") { Task { await model.requeue(job) } }
            Button("Dismiss") { model.activity.dismissFailure(id: job.id) }
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 10)
        .background(Color.red.opacity(0.055))
    }
}

private struct LaneRow: View {
    @EnvironmentObject private var model: AppModel
    let lane: LaneStatus

    var body: some View {
        HStack(spacing: 12) {
            Image(systemName: lane.paused ? "pause.circle.fill" : "circle")
                .foregroundStyle(.secondary)
                .frame(width: 9)
            Text("\(lane.workspace) · \(lane.preset)")
                .font(.callout.weight(.semibold))
                .frame(width: 180, alignment: .leading)
            Text("\(lane.queued) file\(lane.queued == 1 ? "" : "s") waiting")
                .font(.caption.monospaced())
                .foregroundStyle(.secondary)
            Spacer()
            Button("Open") {
                model.openLane(preset: lane.preset, workspace: lane.workspace)
            }
            Button(lane.paused ? "Resume" : "Pause") {
                Task { await model.togglePauseLane(lane) }
            }
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 9)
    }
}

private struct CompletedJobRow: View {
    @EnvironmentObject private var model: AppModel
    let job: JobSnapshot

    private var time: String {
        (job.endedAt ?? job.startedAt).formatted(date: .omitted, time: .shortened)
    }

    var body: some View {
        HStack(spacing: 12) {
            Circle()
                .fill(job.state == .done ? Color.green : job.state == .cancelled ? Color.orange : Color.red)
                .frame(width: 7, height: 7)
            Text(job.lane).font(.callout.weight(.semibold)).frame(width: 150, alignment: .leading)
            Text(job.detail).font(.caption.monospaced()).foregroundStyle(.secondary).lineLimit(1)
            Spacer()
            Text(job.state == .done ? "\(job.outputCount) outputs" : job.state == .cancelled ? "cancelled" : "failed")
                .font(.caption.monospacedDigit())
                .foregroundStyle(job.state == .done ? Color.green : job.state == .cancelled ? Color.orange : Color.red)
            Text(time).font(.caption).foregroundStyle(.secondary).frame(width: 62, alignment: .trailing)
            if job.state == .done && !job.outputPaths.isEmpty {
                Button("Zip") { Task { await model.archive(job, thenUpload: false) } }
                Button("Zip and Upload") { Task { await model.archive(job, thenUpload: true) } }
            }
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 8)
    }
}
