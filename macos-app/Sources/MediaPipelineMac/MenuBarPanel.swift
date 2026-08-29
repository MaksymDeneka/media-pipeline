import AppKit
import SwiftUI

struct MenuBarPanel: View {
    @Environment(\.openWindow) private var openWindow
    @EnvironmentObject private var model: AppModel

    var body: some View {
        VStack(spacing: 0) {
            header
            Divider()
            if model.activity.running.isEmpty {
                HStack {
                    Image(systemName: model.workerRunning ? "checkmark.circle" : "stop.circle")
                        .foregroundStyle(.secondary)
                    Text(model.workerRunning ? "No jobs are running." : "The worker is stopped.")
                        .font(.callout)
                    Spacer()
                }
                .padding(15)
            } else {
                ForEach(model.activity.running.prefix(3)) { job in
                    MenuJobRow(job: job)
                    Divider()
                }
            }
            queueSummary
            Divider()
            footer
        }
        .frame(width: 340)
    }

    private var header: some View {
        HStack(spacing: 10) {
            Circle()
                .fill(model.workerRunning ? Color.green : Color.secondary.opacity(0.55))
                .frame(width: 7, height: 7)
            VStack(alignment: .leading, spacing: 2) {
                Text("Media Pipeline").font(.headline)
                Text("\(model.statusText) · \(model.encoderName)")
                    .font(.caption2.monospaced())
                    .foregroundStyle(.secondary)
                    .lineLimit(1)
            }
            Spacer()
            if model.workerRunning {
                Button {
                    Task { await model.togglePauseAll() }
                } label: {
                    Image(systemName: model.status?.pausedAll == true ? "play.fill" : "pause.fill")
                }
                .help(model.status?.pausedAll == true ? "Resume all" : "Pause all")
                Button {
                    Task { await model.restartWorker() }
                } label: {
                    Image(systemName: "arrow.clockwise")
                }
                .help("Restart worker")
            } else {
                Button {
                    Task { await model.startWorker() }
                } label: {
                    Image(systemName: "play.fill")
                }
                .help("Start worker")
            }
        }
        .buttonStyle(.bordered)
        .padding(14)
    }

    private var queueSummary: some View {
        Button {
            showMainWindow(section: .activity)
        } label: {
            HStack {
                VStack(alignment: .leading, spacing: 3) {
                    Text("\(model.queuedCount) files queued")
                        .font(.callout.weight(.semibold))
                    Text("\(model.status?.workspaces.count ?? model.configuration.workspaces.count) workspaces")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
                Spacer()
                Image(systemName: "chevron.right").foregroundStyle(.tertiary)
            }
            .contentShape(Rectangle())
        }
        .buttonStyle(.plain)
        .padding(14)
    }

    private var footer: some View {
        HStack {
            Button("Open Media Pipeline") { showMainWindow(section: .activity) }
            Spacer()
            Button("Logs") { model.openLogs() }
            Button("Quit") { NSApp.terminate(nil) }
        }
        .buttonStyle(.plain)
        .font(.caption)
        .padding(.horizontal, 14)
        .padding(.vertical, 11)
        .background(.bar)
    }

    private func showMainWindow(section: AppSection) {
        model.selection = section
        openWindow(id: "main")
        NSApp.activate(ignoringOtherApps: true)
    }
}

private struct MenuJobRow: View {
    let job: JobSnapshot

    var body: some View {
        VStack(alignment: .leading, spacing: 7) {
            HStack {
                Text(job.lane).font(.callout.weight(.semibold))
                Spacer()
                Text(job.total > 0 ? "\(Int(job.fraction * 100))%" : "Starting")
                    .font(.caption.monospacedDigit())
            }
            Text(job.detail)
                .font(.caption2.monospaced())
                .foregroundStyle(.secondary)
                .lineLimit(1)
            ProgressView(value: job.fraction)
                .progressViewStyle(.linear)
                .tint(.primary)
        }
        .padding(.horizontal, 14)
        .padding(.vertical, 10)
    }
}
