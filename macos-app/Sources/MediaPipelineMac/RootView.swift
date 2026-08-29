import SwiftUI

struct RootView: View {
    @EnvironmentObject private var model: AppModel

    var body: some View {
        NavigationSplitView {
            List(selection: $model.selection) {
                Section("Monitor") {
                    sidebarRow(.activity, badge: model.activity.running.count)
                    sidebarRow(
                        .uploads,
                        badge: model.uploads.workspaces.reduce(0) { $0 + $1.files.count }
                    )
                }
                Section("Configure") {
                    sidebarRow(.presets)
                    sidebarRow(.settings)
                }
            }
            .navigationSplitViewColumnWidth(min: 195, ideal: 230, max: 280)
            .safeAreaInset(edge: .bottom) {
                SidebarWorkerStatus()
            }
        } detail: {
            switch model.selection ?? .activity {
            case .activity:
                ActivityView()
            case .uploads:
                UploadsView()
            case .presets:
                PresetsView(store: model.configuration)
            case .settings:
                SettingsView(store: model.configuration, loginItem: model.loginItem)
            }
        }
        .toolbar {
            ToolbarItemGroup(placement: .primaryAction) {
                if model.workerRunning {
                    Button(model.status?.pausedAll == true ? "Resume All" : "Pause All") {
                        Task { await model.togglePauseAll() }
                    }
                    Button {
                        Task { await model.restartWorker() }
                    } label: {
                        Label("Restart", systemImage: "arrow.clockwise")
                    }
                    Button {
                        Task { await model.stopWorker() }
                    } label: {
                        Label("Stop", systemImage: "stop.fill")
                    }
                } else {
                    Button {
                        Task { await model.startWorker() }
                    } label: {
                        Label("Start", systemImage: "play.fill")
                    }
                }
                Button {
                    model.openLogs()
                } label: {
                    Label("Logs", systemImage: "doc.text")
                }
            }
        }
        .disabled(model.workerBusy)
        .alert("Media Pipeline", isPresented: errorPresented) {
            Button("OK") { model.errorMessage = nil }
        } message: {
            Text(model.errorMessage ?? "Unknown error")
        }
    }

    @ViewBuilder
    private func sidebarRow(_ section: AppSection, badge: Int = 0) -> some View {
        HStack {
            Label(section.title, systemImage: section.symbol)
            Spacer()
            if badge > 0 {
                Text("\(badge)")
                    .font(.caption.monospacedDigit())
                    .foregroundStyle(.secondary)
            }
        }
        .tag(section)
    }

    private var errorPresented: Binding<Bool> {
        Binding(
            get: { model.errorMessage != nil },
            set: { if !$0 { model.errorMessage = nil } }
        )
    }
}

private struct SidebarWorkerStatus: View {
    @EnvironmentObject private var model: AppModel

    var body: some View {
        VStack(alignment: .leading, spacing: 9) {
            Divider()
            HStack(spacing: 8) {
                Circle()
                    .fill(model.workerRunning ? Color.green : Color.secondary.opacity(0.55))
                    .frame(width: 7, height: 7)
                VStack(alignment: .leading, spacing: 2) {
                    Text(model.statusText)
                        .font(.caption.weight(.semibold))
                    Text(model.encoderName)
                        .font(.caption2.monospaced())
                        .foregroundStyle(.secondary)
                }
            }
            Button {
                model.openPipelineFolder()
            } label: {
                Label("Open Pipeline Folder", systemImage: "folder")
                    .font(.caption)
            }
            .buttonStyle(.plain)
        }
        .padding(.horizontal, 13)
        .padding(.bottom, 11)
        .background(.bar)
    }
}
