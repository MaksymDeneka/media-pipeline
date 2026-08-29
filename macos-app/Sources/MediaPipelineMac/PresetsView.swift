import SwiftUI

struct PresetsView: View {
    @EnvironmentObject private var model: AppModel
    @ObservedObject var store: ConfigurationStore
    @State private var selectedPreset: String?
    @State private var addingPreset = false
    @State private var newPresetName = ""
    @State private var presetError: String?
    @State private var pendingRemoval: String?

    var body: some View {
        HSplitView {
            List(store.presetOrder, id: \.self, selection: $selectedPreset) { preset in
                Text(preset).tag(Optional(preset))
            }
            .frame(minWidth: 180, idealWidth: 210, maxWidth: 260)
            .safeAreaInset(edge: .bottom) {
                HStack {
                    Button { addingPreset = true } label: {
                        Image(systemName: "plus")
                    }
                    Button {
                        if let selectedPreset { pendingRemoval = selectedPreset }
                    } label: {
                        Image(systemName: "minus")
                    }
                    .disabled(selectedPreset == nil)
                    Spacer()
                }
                .padding(9)
                .background(.bar)
            }

            if let selectedPreset, store.presets[selectedPreset] != nil {
                presetEditor(selectedPreset)
            } else {
                VStack(spacing: 9) {
                    Image(systemName: "slider.horizontal.3")
                        .font(.largeTitle)
                        .foregroundStyle(.secondary)
                    Text("Select a preset").font(.headline)
                    Text("Choose a preset to inspect its inherited values and overrides.")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
                .frame(maxWidth: .infinity, maxHeight: .infinity)
            }
        }
        .navigationTitle("Presets")
        .toolbar {
            ToolbarItem(placement: .automatic) {
                Button("Save and Restart") {
                    Task { await model.saveConfigurationAndRestart() }
                }
                .disabled(!store.hasUnsavedChanges)
            }
        }
        .onAppear {
            if selectedPreset == nil { selectedPreset = store.presetOrder.first }
        }
        .sheet(isPresented: $addingPreset) {
            VStack(alignment: .leading, spacing: 15) {
                Text("New preset").font(.title2.weight(.semibold))
                TextField("Preset name", text: $newPresetName)
                    .textFieldStyle(.roundedBorder)
                HStack {
                    Spacer()
                    Button("Cancel") { addingPreset = false }
                    Button("Add") { addPreset() }
                        .keyboardShortcut(.defaultAction)
                        .disabled(newPresetName.trimmingCharacters(in: .whitespaces).isEmpty)
                }
            }
            .padding(22)
            .frame(width: 380)
        }
        .alert("Remove preset?", isPresented: removalPresented) {
            Button("Cancel", role: .cancel) { pendingRemoval = nil }
            Button("Remove", role: .destructive) {
                if let pendingRemoval {
                    store.removePreset(pendingRemoval)
                    selectedPreset = store.presetOrder.first
                    self.pendingRemoval = nil
                }
            }
        } message: {
            Text("Its folders and media will remain on disk.")
        }
        .alert("Preset", isPresented: presetErrorPresented) {
            Button("OK") { presetError = nil }
        } message: {
            Text(presetError ?? "Unknown error")
        }
    }

    private func presetEditor(_ preset: String) -> some View {
        ScrollView {
            LazyVStack(alignment: .leading, spacing: 0) {
                HStack {
                    VStack(alignment: .leading, spacing: 4) {
                        Text(preset).font(.title2.weight(.semibold))
                        Text("Editing an inherited value creates an override for this preset.")
                            .font(.caption)
                            .foregroundStyle(.secondary)
                    }
                    Spacer()
                    Button("Open Pipeline Folder") {
                        model.openPipelineFolder()
                    }
                }
                .padding(.horizontal, 22)
                .padding(.vertical, 15)
                .overlay(alignment: .bottom) { Divider() }

                ForEach(SettingCatalog.presetGroups, id: \.self) { group in
                    Text(group.uppercased())
                        .font(.caption2.weight(.semibold))
                        .foregroundStyle(.secondary)
                        .padding(.top, 21)
                        .padding(.bottom, 7)
                    ForEach(SettingCatalog.presetDefinitions.filter { $0.group == group }) { definition in
                        SettingEditorRow(store: store, definition: definition, preset: preset)
                        Divider()
                    }
                }
                .padding(.horizontal, 22)
            }
            .padding(.bottom, 28)
        }
    }

    private func addPreset() {
        do {
            try store.addPreset(newPresetName)
            selectedPreset = newPresetName.trimmingCharacters(in: .whitespacesAndNewlines)
            newPresetName = ""
            addingPreset = false
        } catch {
            presetError = "Use a unique name without periods or path separators."
        }
    }

    private var removalPresented: Binding<Bool> {
        Binding(
            get: { pendingRemoval != nil },
            set: { if !$0 { pendingRemoval = nil } }
        )
    }

    private var presetErrorPresented: Binding<Bool> {
        Binding(
            get: { presetError != nil },
            set: { if !$0 { presetError = nil } }
        )
    }
}
