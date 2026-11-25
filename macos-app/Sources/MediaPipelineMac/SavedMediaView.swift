import AppKit
import SwiftUI
import UniformTypeIdentifiers

struct SavedMediaView: View {
    private enum Layout: String {
        case tiles, list
    }

    private enum MediaFilter: String, CaseIterable {
        case all = "All media"
        case images = "Images"
        case videos = "Videos"

        func includes(_ item: SavedMedia) -> Bool {
            switch self {
            case .all: return true
            case .images: return item.isImage
            case .videos: return !item.isImage
            }
        }
    }

    private enum SortOrder: String, CaseIterable {
        case newest, oldest, nameAscending, nameDescending, largest, smallest

        var title: String {
            switch self {
            case .newest: return "Newest first"
            case .oldest: return "Oldest first"
            case .nameAscending: return "Name (A–Z)"
            case .nameDescending: return "Name (Z–A)"
            case .largest: return "Largest first"
            case .smallest: return "Smallest first"
            }
        }

        func precedes(_ lhs: SavedMedia, _ rhs: SavedMedia) -> Bool {
            switch self {
            case .newest: return lhs.savedAt > rhs.savedAt
            case .oldest: return lhs.savedAt < rhs.savedAt
            case .nameAscending: return lhs.name.localizedStandardCompare(rhs.name) == .orderedAscending
            case .nameDescending: return lhs.name.localizedStandardCompare(rhs.name) == .orderedDescending
            case .largest: return lhs.bytes > rhs.bytes
            case .smallest: return lhs.bytes < rhs.bytes
            }
        }
    }

    @EnvironmentObject private var model: AppModel
    @ObservedObject var store: SavedMediaStore
    @ObservedObject var configuration: ConfigurationStore
    @State private var selectedID: UUID?
    @State private var copiesText = "20"
    @State private var search = ""
    @State private var mediaFilter = MediaFilter.all
    @State private var importNotice: String?
    @State private var editingTab: SavedMediaCollection?
    @State private var tabName = ""
    @State private var tabError: String?
    @State private var showsTabEditor = false
    @AppStorage("savedMediaTab") private var tabID = SavedMediaCollection.library.id.uuidString
    @AppStorage("savedMediaLayout") private var layout = Layout.tiles
    @AppStorage("savedMediaSortOrder") private var sortOrder = SortOrder.newest
    @AppStorage("savedMediaWorkspace") private var workspace = "LC"
    @AppStorage("savedMediaPreset") private var preset = "bulk"

    private var activeTab: SavedMediaCollection {
        store.collections.first { $0.id.uuidString == tabID }
            ?? store.collections.first { $0.id == SavedMediaCollection.library.id }
            ?? .library
    }

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
        let tab = activeTab
        let tabItems = store.items(in: tab.id)
        let query = search.trimmingCharacters(in: .whitespacesAndNewlines)
        let items = tabItems.filter {
            mediaFilter.includes($0) && (query.isEmpty || $0.name.localizedStandardContains(query))
        }.sorted(by: sortOrder.precedes)
        VStack(alignment: .leading, spacing: 0) {
            HStack {
                VStack(alignment: .leading, spacing: 4) {
                    Text("Keep media for future batches").font(.headline)
                    Text(importNotice ?? "Save an original once, then choose how many processed copies to make.")
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
            libraryToolbar
            Divider()
            tabBar
            Divider()
            if tabItems.isEmpty {
                VStack(spacing: 12) {
                    Image(systemName: "photo.on.rectangle.angled").font(.system(size: 36))
                    Text("No media in \(tab.name)").font(.headline)
                    Button("Save Media…", action: importMedia)
                }
                .frame(maxWidth: .infinity, maxHeight: .infinity)
            } else {
                if items.isEmpty {
                    VStack(spacing: 12) {
                        Image(systemName: "magnifyingglass").font(.system(size: 32))
                            .accessibilityHidden(true)
                        Text("No matching media").font(.headline)
                        Button("Clear Filters") {
                            search = ""
                            mediaFilter = .all
                        }
                    }
                    .frame(maxWidth: .infinity, maxHeight: .infinity)
                } else {
                    mediaBrowser(items).id(tab.id)
                    Divider()
                    if let selected = items.first(where: { $0.id == selectedID }) {
                        generationControls(selected)
                    } else {
                        Text("Select saved media to make copies.")
                            .foregroundStyle(.secondary)
                            .padding(.horizontal, 22)
                            .padding(.vertical, 12)
                    }
                }
            }
        }
        .navigationTitle("Saved Media")
        .disabled(store.isBusy)
        .onAppear {
            repairSelection(in: items)
            repairCopyOptions()
        }
        .onChange(of: items) { repairSelection(in: $0) }
        .onChange(of: configuration.presetOrder) { _ in repairCopyOptions() }
        .sheet(isPresented: $showsTabEditor) { tabEditor }
    }

    private var tabBar: some View {
        HStack(spacing: 10) {
            ViewThatFits(in: .horizontal) {
                tabButtons.fixedSize(horizontal: true, vertical: false)
                ScrollViewReader { proxy in
                    ScrollView(.horizontal, showsIndicators: false) { tabButtons }
                        .onChange(of: tabID) { proxy.scrollTo($0, anchor: .center) }
                        .onAppear { proxy.scrollTo(activeTab.id.uuidString, anchor: .center) }
                }
            }
            .frame(height: 24)
            Button { editTab() } label: { Image(systemName: "plus") }
                .accessibilityLabel("New tab")
                .help("New tab")
            Spacer(minLength: 0)
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 7)
    }

    private var tabButtons: some View {
        HStack(spacing: 6) {
            ForEach(store.collections) { tab in
                let isSelected = tab.id == activeTab.id
                Button { selectTab(tab.id) } label: {
                    Text(tab.name)
                        .lineLimit(1)
                        .fontWeight(isSelected ? .semibold : .regular)
                        .padding(.horizontal, 12)
                        .padding(.vertical, 4)
                        .frame(maxWidth: 180)
                        .background(isSelected ? Color.accentColor.opacity(0.20) : Color.primary.opacity(0.06),
                                    in: RoundedRectangle(cornerRadius: 5))
                        .contentShape(Rectangle())
                }
                .buttonStyle(.plain)
                .accessibilityAddTraits(isSelected ? .isSelected : [])
                .help(tab.name)
                .id(tab.id.uuidString)
                .contextMenu {
                    Button("Rename Tab…") { editTab(tab) }
                    if tab.id != SavedMediaCollection.library.id {
                        Button("Delete Empty Tab", role: .destructive) {
                            do {
                                try store.removeCollection(tab.id)
                                if tabID == tab.id.uuidString { selectTab(SavedMediaCollection.library.id) }
                            } catch { model.errorMessage = error.localizedDescription }
                        }
                        .disabled(!store.items(in: tab.id).isEmpty)
                    }
                }
            }
        }
    }

    private var tabEditor: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text(editingTab == nil ? "New Tab" : "Rename Tab").font(.headline)
            TextField("Tab name", text: $tabName).textFieldStyle(.roundedBorder)
            if let tabError { Text(tabError).font(.caption).foregroundStyle(.red) }
            HStack {
                Spacer()
                Button("Cancel") { showsTabEditor = false }.keyboardShortcut(.cancelAction)
                Button(editingTab == nil ? "Create" : "Rename") {
                    do {
                        if let editingTab {
                            try store.renameCollection(editingTab.id, to: tabName)
                        } else {
                            let tab = try store.createCollection(named: tabName)
                            selectTab(tab.id)
                        }
                        showsTabEditor = false
                    } catch { tabError = error.localizedDescription }
                }
                .keyboardShortcut(.defaultAction)
                .disabled(tabName.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty)
            }
        }
        .padding(20)
        .frame(width: 320)
    }

    private func editTab(_ tab: SavedMediaCollection? = nil) {
        editingTab = tab
        tabName = tab?.name ?? ""
        tabError = nil
        showsTabEditor = true
    }

    private func selectTab(_ id: UUID) {
        guard id.uuidString != tabID else { return }
        tabID = id.uuidString
        selectedID = nil
        search = ""
        mediaFilter = .all
        importNotice = nil
    }

    private var libraryToolbar: some View {
        HStack(spacing: 12) {
            HStack(spacing: 6) {
                Image(systemName: "magnifyingglass")
                    .foregroundStyle(.secondary)
                    .accessibilityHidden(true)
                TextField("Search by filename", text: $search)
                    .textFieldStyle(.plain)
                if !search.isEmpty {
                    Button {
                        search = ""
                    } label: {
                        Image(systemName: "xmark.circle.fill").foregroundStyle(.secondary)
                    }
                    .buttonStyle(.plain)
                    .accessibilityLabel("Clear search")
                }
            }
            .padding(.horizontal, 8)
            .frame(height: 25)
            .frame(minWidth: 125, maxWidth: 235)
            .background(Color.primary.opacity(0.06), in: RoundedRectangle(cornerRadius: 6))

            Picker("Media type", selection: $mediaFilter) {
                ForEach(MediaFilter.allCases, id: \.self) { Text($0.rawValue).tag($0) }
            }
            .labelsHidden()
            .fixedSize()

            Spacer(minLength: 0)

            Picker("Sort by", selection: $sortOrder) {
                ForEach(SortOrder.allCases, id: \.self) { Text($0.title).tag($0) }
            }
            .fixedSize()

            Picker("Layout", selection: $layout) {
                Label("Tiles", systemImage: "square.grid.2x2").tag(Layout.tiles)
                Label("List", systemImage: "list.bullet").tag(Layout.list)
            }
            .pickerStyle(.segmented)
            .labelsHidden()
            .fixedSize()
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 10)
    }

    private func mediaBrowser(_ items: [SavedMedia]) -> some View {
        ScrollViewReader { proxy in
            Group {
                if layout == .tiles {
                    ScrollView {
                        LazyVGrid(columns: [GridItem(.adaptive(minimum: 176), spacing: 14)], spacing: 14) {
                            ForEach(items) { tile($0) }
                        }
                        .padding(22)
                    }
                } else {
                    List(selection: $selectedID) {
                        ForEach(items) { item in
                            HStack(spacing: 12) {
                                MediaThumbnail(url: store.sourceURL(for: item), isImage: item.isImage)
                                mediaCaption(item)
                                Spacer()
                            }
                            .padding(.vertical, 6)
                            .tag(item.id)
                            .contextMenu { moveMenu(for: item) }
                        }
                    }
                }
            }
            .onChange(of: layout) { _ in
                if let selectedID { proxy.scrollTo(selectedID) }
            }
        }
    }

    private func tile(_ item: SavedMedia) -> some View {
        let isSelected = item.id == selectedID
        let shape = RoundedRectangle(cornerRadius: 10)
        return Button {
            selectedID = item.id
        } label: {
            VStack(alignment: .leading, spacing: 0) {
                MediaThumbnail(url: store.sourceURL(for: item), isImage: item.isImage, size: .tile)
                    .accessibilityHidden(true)
                mediaCaption(item)
                    .padding(.horizontal, 12)
                    .padding(.vertical, 9)
                    .frame(maxWidth: .infinity, alignment: .leading)
            }
            .background(isSelected ? Color.accentColor.opacity(0.17) : Color.primary.opacity(0.045))
            .clipShape(shape)
            .overlay { shape.strokeBorder(isSelected ? Color.accentColor : Color.primary.opacity(0.12), lineWidth: isSelected ? 2 : 1) }
            .contentShape(shape)
        }
        .buttonStyle(.plain)
        .accessibilityAddTraits(isSelected ? .isSelected : [])
        .contextMenu { moveMenu(for: item) }
    }

    private func mediaCaption(_ item: SavedMedia) -> some View {
        let size = ByteCountFormatter.string(fromByteCount: item.bytes, countStyle: .file)
        return VStack(alignment: .leading, spacing: 3) {
            Text(item.name).lineLimit(1).truncationMode(.middle)
            Text("\(item.isImage ? "Image" : "Video") · \(size)")
                .font(.caption).foregroundStyle(.secondary)
        }
    }

    private func generationControls(_ item: SavedMedia) -> some View {
        VStack(alignment: .leading, spacing: 8) {
            HStack(spacing: 8) {
                Text(item.name).font(.headline).lineLimit(1).truncationMode(.middle)
                    .help(item.name)
                Spacer(minLength: 8)
                moveMenu(for: item).fixedSize()
                Button("Preview") { NSWorkspace.shared.open(store.sourceURL(for: item)) }
                Button("Remove") {
                    do { try store.remove(item) }
                    catch { model.errorMessage = error.localizedDescription }
                }
            }
            HStack(alignment: .bottom, spacing: 8) {
                VStack(alignment: .leading, spacing: 3) {
                    Text("Copies").font(.caption)
                    TextField("1 to 9999", text: $copiesText)
                        .textFieldStyle(.roundedBorder)
                        .accessibilityLabel("Number of copies")
                }
                .frame(width: 60)

                VStack(alignment: .leading, spacing: 3) {
                    Text("Workspace").font(.caption)
                    Picker("Workspace", selection: $workspace) {
                        ForEach(configuration.workspaces, id: \.self) { Text($0).tag($0) }
                    }.labelsHidden()
                }
                .frame(minWidth: 100, maxWidth: 132)

                VStack(alignment: .leading, spacing: 3) {
                    Text("Preset").font(.caption)
                    Picker("Preset", selection: $preset) {
                        ForEach(presetNames, id: \.self) { Text($0).tag($0) }
                    }.labelsHidden()
                }
                .frame(minWidth: 105, maxWidth: 250)

                Spacer(minLength: 4)

                Button {
                    model.openOutputFolder(preset: preset, workspace: workspace)
                } label: {
                    Label("Output Folder", systemImage: "folder")
                }
                .fixedSize()
                .disabled(!configuration.workspaces.contains(workspace) || !presetNames.contains(preset))
                .help("Open \(workspace)/\(preset)/output in Finder")

                Button("Make Copies") {
                    guard let copies else { return }
                    let requestedPreset = preset
                    let requestedWorkspace = workspace
                    store.isBusy = true
                    Task {
                        defer { store.isBusy = false }
                        await model.queueSavedMedia(item, copies: copies, preset: requestedPreset, workspace: requestedWorkspace)
                    }
                }
                .buttonStyle(.borderedProminent)
                .fixedSize()
                .disabled(copies == nil || !presetNames.contains(preset) || model.workerBusy)
                .help("Apply the preset to each copy. Set and alternating counts don’t apply.")
            }
            if let feedback {
                Text(feedback)
                    .font(.caption)
                    .foregroundStyle(copies == nil ? Color.red : Color.secondary)
                    .lineLimit(3)
                    .fixedSize(horizontal: false, vertical: true)
            }
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 12)
    }

    private var feedback: String? {
        if copies == nil { return "Enter a whole number from 1 to 9999." }
        if presetNames.isEmpty { return "Enable a preset in Presets to make copies." }
        let warning = configuration.hasUnsavedChanges
            ? "Copies use saved settings. Save your changes in Presets or Settings first."
            : nil
        let messages = [warning, model.notice].compactMap { $0 }.filter { !$0.isEmpty }
        return messages.isEmpty ? nil : messages.joined(separator: "\n")
    }

    private func repairSelection(in items: [SavedMedia]) {
        if !items.contains(where: { $0.id == selectedID }) { selectedID = items.first?.id }
    }

    private func repairCopyOptions() {
        if !configuration.workspaces.contains(workspace) { workspace = configuration.workspaces.first ?? "LC" }
        if !presetNames.contains(preset) { preset = presetNames.first ?? "" }
    }

    @ViewBuilder
    private func moveMenu(for item: SavedMedia) -> some View {
        if store.collections.count > 1 {
            Menu("Move to Tab") {
                ForEach(store.collections.filter { $0.id != activeTab.id }) { tab in
                    Button(tab.name) {
                        do { try store.move(item, to: tab.id) }
                        catch { model.errorMessage = error.localizedDescription }
                    }
                }
            }
        }
    }

    private func importMedia() {
        let targetTab = activeTab.id
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
                var saved = 0
                var duplicates = 0
                for url in panel.urls {
                    do {
                        if try await store.save(url, in: targetTab) { saved += 1 }
                        else { duplicates += 1 }
                    } catch { model.errorMessage = "Could not save \(url.lastPathComponent): \(error.localizedDescription)" }
                }
                let skipped = duplicates == 1 ? "Skipped 1 duplicate." : "Skipped \(duplicates) duplicates."
                importNotice = duplicates == 0 ? nil : "Saved \(saved). \(skipped)"
            }
        }
    }
}
