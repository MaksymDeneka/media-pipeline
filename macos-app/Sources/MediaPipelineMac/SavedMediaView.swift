import AppKit
import SwiftUI
import UniformTypeIdentifiers

struct SavedMediaView: View {
    @EnvironmentObject private var model: AppModel
    @ObservedObject var store: SavedMediaStore
    @ObservedObject var configuration: ConfigurationStore
    @State private var selectedID: UUID?
    @State private var copiesText = "20"
    @AppStorage("savedMediaWorkspace") private var workspace = "LC"
    @AppStorage("savedMediaPreset") private var preset = "bulk"

    private var selected: SavedMedia? { store.items.first { $0.id == selectedID } }
    private var copies: Int? {
        guard let count = Int(copiesText), (1...9999).contains(count) else { return nil }
        return count
    }
    private var presetNames: [String] {
        configuration.presetOrder.filter {
            let value = configuration.presets[$0]?.first {
                $0.key.caseInsensitiveCompare("Enabled") == .orderedSame
            }?.value.lowercased() ?? "true"
            return !["false", "0", "no", "off"].contains(value)
        }
    }

    var body: some View {
        VStack(alignment: .leading, spacing: 0) {
            HStack {
                VStack(alignment: .leading, spacing: 4) {
                    Text("Keep media for future batches").font(.headline)
                    Text("Save an original once, then choose how many processed copies to make.")
                        .font(.caption).foregroundStyle(.secondary)
                }
                Spacer()
                if store.isBusy { ProgressView().controlSize(.small) }
                Button("Save Media…", action: importMedia)
            }
            .padding(22)
            Divider()
            ForEach(store.loadWarnings, id: \.self) { warning in
                Text(warning).font(.caption).foregroundStyle(.red).padding(.horizontal, 22)
            }
            if store.items.isEmpty {
                VStack(spacing: 12) {
                    Image(systemName: "photo.on.rectangle.angled").font(.system(size: 36))
                    Text("No saved media yet").font(.headline)
                    Text("Add an image or video to reuse it in later batches.").foregroundStyle(.secondary)
                    Button("Save Media…", action: importMedia)
                }
                .frame(maxWidth: .infinity, maxHeight: .infinity)
            } else {
                List(selection: $selectedID) {
                    ForEach(store.items) { item in
                        HStack(spacing: 12) {
                            Image(systemName: item.isImage ? "photo" : "film")
                                .font(.title2).frame(width: 30)
                            VStack(alignment: .leading, spacing: 3) {
                                Text(item.name).lineLimit(1)
                                Text("\(item.isImage ? "Image" : "Video") · \(ByteCount.string(item.bytes))")
                                    .font(.caption).foregroundStyle(.secondary)
                            }
                            Spacer()
                        }
                        .padding(.vertical, 6)
                        .tag(item.id)
                    }
                }
                Divider()
                if let selected {
                    generationControls(selected)
                } else {
                    Text("Select saved media to make copies.")
                        .foregroundStyle(.secondary).padding(22)
                }
            }
        }
        .navigationTitle("Saved Media")
        .disabled(store.isBusy)
        .onAppear { repairSelection() }
        .onChange(of: store.items) { _ in repairSelection() }
        .onChange(of: configuration.presetOrder) { _ in repairSelection() }
    }

    private func generationControls(_ item: SavedMedia) -> some View {
        VStack(alignment: .leading, spacing: 14) {
            HStack {
                Text(item.name).font(.headline).lineLimit(1)
                Spacer()
                Button("Preview") { NSWorkspace.shared.open(store.sourceURL(for: item)) }
                Button("Remove") {
                    do { try store.remove(item) }
                    catch { model.errorMessage = error.localizedDescription }
                }
            }
            HStack(alignment: .bottom, spacing: 18) {
                VStack(alignment: .leading) {
                    Text("Copies").font(.caption)
                    TextField("1 to 9999", text: $copiesText)
                        .textFieldStyle(.roundedBorder).frame(width: 100)
                        .accessibilityLabel("Number of copies")
                }
                VStack(alignment: .leading) {
                    Text("Workspace").font(.caption)
                    Picker("Workspace", selection: $workspace) {
                        ForEach(configuration.workspaces, id: \.self) { Text($0).tag($0) }
                    }.labelsHidden().frame(width: 120)
                }
                VStack(alignment: .leading) {
                    Text("Preset").font(.caption)
                    Picker("Preset", selection: $preset) {
                        ForEach(presetNames, id: \.self) { Text($0).tag($0) }
                    }.labelsHidden().frame(minWidth: 140)
                }
                Spacer()
                Button("Make Copies") {
                    guard let copies else { return }
                    Task {
                        store.isBusy = true
                        defer { store.isBusy = false }
                        await model.queueSavedMedia(item, copies: copies, preset: preset, workspace: workspace)
                    }
                }
                .buttonStyle(.borderedProminent)
                .disabled(copies == nil || !presetNames.contains(preset))
            }
            if copies == nil {
                Text("Enter a whole number from 1 to 9999.").font(.caption).foregroundStyle(.red)
            }
            if presetNames.isEmpty {
                Text("Enable a preset in Presets to make copies.").font(.caption).foregroundStyle(.secondary)
            }
            Text("Applies the preset's image or video modifications to each copy. Outputs go to \(workspace)/\(preset)/output. Set and alternating counts do not apply.")
                .font(.caption).foregroundStyle(.secondary)
            if configuration.hasUnsavedChanges {
                Text("Uses saved settings. Save changes in Presets or Settings before making copies.")
                    .font(.caption).foregroundStyle(.secondary)
            }
            if let notice = model.notice { Text(notice).font(.caption).foregroundStyle(.secondary) }
        }
        .padding(22)
    }

    private func repairSelection() {
        if selected == nil { selectedID = store.items.first?.id }
        if !configuration.workspaces.contains(workspace) { workspace = configuration.workspaces.first ?? "LC" }
        if !presetNames.contains(preset) { preset = presetNames.first ?? "" }
    }

    private func importMedia() {
        let panel = NSOpenPanel()
        panel.title = "Save media for reuse"
        panel.canChooseDirectories = false
        panel.allowsMultipleSelection = true
        panel.allowedContentTypes = (SavedMedia.imageExtensions.union(SavedMedia.videoExtensions))
            .sorted().compactMap { UTType(filenameExtension: $0) }
        panel.begin { response in
            guard response == .OK else { return }
            Task { @MainActor in
                store.isBusy = true
                defer { store.isBusy = false }
                for url in panel.urls {
                    do { try await store.save(url) }
                    catch { model.errorMessage = "Could not save \(url.lastPathComponent): \(error.localizedDescription)" }
                }
            }
        }
    }
}
