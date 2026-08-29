import SwiftUI

struct UploadsView: View {
    @EnvironmentObject private var model: AppModel
    @State private var expandedWorkspaces: Set<String> = ["LC"]

    var body: some View {
        ScrollView {
            LazyVStack(alignment: .leading, spacing: 0) {
                header
                sectionHeader("Ready to upload")
                ForEach(model.uploads.workspaces) { workspace in
                    workspaceRow(workspace)
                    Divider()
                }
                sectionHeader("Transfers")
                if model.uploads.transfers.isEmpty {
                    Text("No transfers yet.")
                        .font(.callout)
                        .foregroundStyle(.secondary)
                        .padding(.horizontal, 22)
                        .padding(.vertical, 12)
                } else {
                    ForEach(model.uploads.transfers) { transfer in
                        TransferRow(transfer: transfer)
                        Divider()
                    }
                }
            }
            .padding(.bottom, 28)
        }
        .navigationTitle("Uploads")
        .alert("Upload failed", isPresented: uploadErrorPresented) {
            Button("OK") { model.uploads.errorMessage = nil }
        } message: {
            Text(model.uploads.errorMessage ?? "Unknown error")
        }
    }

    private var header: some View {
        HStack {
            VStack(alignment: .leading, spacing: 4) {
                Text("\(stagedCount) staged file\(stagedCount == 1 ? "" : "s")")
                    .font(.headline)
                Text("Cancelling keeps complete local chunks for the next attempt.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }
            Spacer()
            Button("Refresh") { Task { await model.refresh() } }
            if model.uploads.isBusy {
                Button("Cancel") { model.uploads.cancel() }
            } else if stagedCount > 0 {
                Button("Upload All") {
                    model.uploads.upload(model.uploads.workspaces.flatMap(\.files))
                }
            }
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 15)
        .overlay(alignment: .bottom) { Divider() }
    }

    private var stagedCount: Int {
        model.uploads.workspaces.reduce(0) { $0 + $1.files.count }
    }

    private func sectionHeader(_ title: String) -> some View {
        Text(title.uppercased())
            .font(.caption2.weight(.semibold))
            .foregroundStyle(.secondary)
            .padding(.horizontal, 22)
            .padding(.top, 21)
            .padding(.bottom, 7)
    }

    private func workspaceRow(_ workspace: WorkspaceUploads) -> some View {
        DisclosureGroup(
            isExpanded: Binding(
                get: { expandedWorkspaces.contains(workspace.name) },
                set: { expanded in
                    if expanded { expandedWorkspaces.insert(workspace.name) }
                    else { expandedWorkspaces.remove(workspace.name) }
                }
            )
        ) {
            if workspace.files.isEmpty {
                Text("No staged files.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
                    .padding(.leading, 26)
                    .padding(.vertical, 8)
            } else {
                ForEach(workspace.files) { file in
                    HStack {
                        Text(file.name).font(.caption.monospaced()).lineLimit(1)
                        Spacer()
                        Text(ByteCount.string(file.bytes))
                            .font(.caption.monospacedDigit())
                            .foregroundStyle(.secondary)
                            .frame(width: 90, alignment: .trailing)
                        Button("Upload") { model.uploads.upload([file]) }
                            .disabled(model.uploads.isBusy)
                    }
                    .padding(.leading, 26)
                    .padding(.vertical, 5)
                }
            }
        } label: {
            HStack {
                Text(workspace.name).font(.callout.weight(.semibold))
                Text("\(workspace.files.count) files · \(ByteCount.string(workspace.totalBytes))")
                    .font(.caption.monospacedDigit())
                    .foregroundStyle(.secondary)
                Spacer()
                if !workspace.files.isEmpty {
                    Button("Upload All") { model.uploads.upload(workspace.files) }
                        .disabled(model.uploads.isBusy)
                }
            }
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 8)
    }

    private var uploadErrorPresented: Binding<Bool> {
        Binding(
            get: { model.uploads.errorMessage != nil },
            set: { if !$0 { model.uploads.errorMessage = nil } }
        )
    }
}

private struct TransferRow: View {
    let transfer: UploadTransfer

    var body: some View {
        VStack(alignment: .leading, spacing: 8) {
            HStack {
                Text(transfer.workspace.uppercased())
                    .font(.caption2.weight(.semibold))
                    .foregroundStyle(.secondary)
                Text(transfer.fileURL.lastPathComponent)
                    .font(.caption.monospaced())
                    .lineLimit(1)
                Spacer()
                Text(transfer.phase)
                    .font(.caption2.weight(.semibold))
                    .foregroundStyle(transfer.phase == "Failed" ? Color.red : Color.secondary)
            }
            ProgressView(value: transfer.fraction).progressViewStyle(.linear).tint(.primary)
            HStack {
                Text("\(transfer.chunksSent) of \(transfer.chunks) chunks")
                Spacer()
                Text("\(ByteCount.string(transfer.bytesSent)) of \(ByteCount.string(transfer.bytes))")
            }
            .font(.caption.monospacedDigit())
            .foregroundStyle(.secondary)
            if let error = transfer.error {
                Text(error).font(.caption).foregroundStyle(.red)
            }
        }
        .padding(.horizontal, 22)
        .padding(.vertical, 11)
    }
}
