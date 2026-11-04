import AppKit
import Darwin
import SwiftUI

@main
struct MediaPipelineApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) private var appDelegate
    @StateObject private var model: AppModel

    init() {
        if CommandLine.arguments.contains("--check-resources") {
            guard AppResources.url(
                forResource: "default-config",
                withExtension: "ini"
            ) != nil else {
                FileHandle.standardError.write(Data("Default configuration resource is missing.\n".utf8))
                Darwin.exit(1)
            }
            print("Swift resources resolved.")
            Darwin.exit(0)
        }
        _model = StateObject(wrappedValue: AppModel())
    }

    var body: some Scene {
        Window("Media Pipeline", id: "main") {
            RootView()
                .environmentObject(model)
                .frame(minWidth: 860, minHeight: 560)
                .onAppear {
                    appDelegate.model = model
                    model.startMonitoring()
                }
        }
        .defaultSize(width: 1120, height: 720)
        .commands {
            CommandGroup(after: .appInfo) {
                Button("Open Pipeline Folder") { model.openPipelineFolder() }
                    .keyboardShortcut("o", modifiers: [.command, .shift])
                Button("Open Logs") { model.openLogs() }
            }
            CommandMenu("Worker") {
                if model.workerRunning {
                    Button(model.status?.pausedAll == true ? "Resume All" : "Pause All") {
                        Task { await model.togglePauseAll() }
                    }
                    Button("Restart") { Task { await model.restartWorker() } }
                    Button("Stop") { Task { await model.stopWorker() } }
                } else {
                    Button("Start") { Task { await model.startWorker() } }
                }
            }
        }

        MenuBarExtra {
            MenuBarPanel()
                .environmentObject(model)
                .onAppear {
                    appDelegate.model = model
                    model.startMonitoring()
                }
        } label: {
            Image(systemName: model.menuBarSymbol)
                .accessibilityLabel("Media Pipeline")
        }
        .menuBarExtraStyle(.window)
    }
}

@MainActor
final class AppDelegate: NSObject, NSApplicationDelegate {
    weak var model: AppModel?
    private var terminationTask: Task<Void, Never>?

    func applicationShouldTerminate(_ sender: NSApplication) -> NSApplication.TerminateReply {
        guard let model else { return .terminateNow }
        if terminationTask == nil {
            terminationTask = Task {
                await model.shutdown()
                sender.reply(toApplicationShouldTerminate: true)
            }
        }
        return .terminateLater
    }

    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool {
        false
    }
}
