import SwiftUI

struct SettingsView: View {
    @EnvironmentObject private var model: AppModel
    @ObservedObject var store: ConfigurationStore
    @ObservedObject var loginItem: LoginItemController

    var body: some View {
        ScrollView {
            LazyVStack(alignment: .leading, spacing: 0) {
                HStack {
                    VStack(alignment: .leading, spacing: 4) {
                        Text("Global defaults").font(.headline)
                        Text("Presets inherit these values unless they override them.")
                            .font(.caption)
                            .foregroundStyle(.secondary)
                    }
                    Spacer()
                    Button("Open Config") { model.openConfig() }
                    Button("Save and Restart") {
                        Task { await model.saveConfigurationAndRestart() }
                    }
                    .disabled(!store.hasUnsavedChanges)
                }
                .padding(.horizontal, 22)
                .padding(.vertical, 15)
                .overlay(alignment: .bottom) { Divider() }

                sectionHeader("Application")
                HStack {
                    VStack(alignment: .leading, spacing: 3) {
                        Text("Start at login").font(.callout.weight(.semibold))
                        Text("Launch the app and worker when you sign in to this Mac.")
                            .font(.caption)
                            .foregroundStyle(.secondary)
                    }
                    Spacer()
                    Toggle("", isOn: Binding(
                        get: { loginItem.enabled },
                        set: { loginItem.setEnabled($0) }
                    )).labelsHidden()
                }
                .padding(.horizontal, 22)
                .padding(.vertical, 8)
                Divider().padding(.horizontal, 22)

                ForEach(SettingCatalog.globalGroups, id: \.self) { group in
                    sectionHeader(group)
                    ForEach(SettingCatalog.globals.filter { $0.group == group }) { definition in
                        SettingEditorRow(store: store, definition: definition, preset: nil)
                        Divider()
                    }
                    .padding(.horizontal, 22)
                }
            }
            .padding(.bottom, 28)
        }
        .navigationTitle("Settings")
        .alert("Start at login", isPresented: loginErrorPresented) {
            Button("OK") { loginItem.errorMessage = nil }
        } message: {
            Text(loginItem.errorMessage ?? "Unknown error")
        }
    }

    private func sectionHeader(_ title: String) -> some View {
        Text(title.uppercased())
            .font(.caption2.weight(.semibold))
            .foregroundStyle(.secondary)
            .padding(.horizontal, 22)
            .padding(.top, 21)
            .padding(.bottom, 7)
    }

    private var loginErrorPresented: Binding<Bool> {
        Binding(
            get: { loginItem.errorMessage != nil },
            set: { if !$0 { loginItem.errorMessage = nil } }
        )
    }
}
