import SwiftUI

struct SettingEditorRow: View {
    @ObservedObject var store: ConfigurationStore
    let definition: SettingDefinition
    let preset: String?

    private var value: String {
        if let preset { return store.presetValue(definition, preset: preset) }
        return store.globalValue(definition)
    }

    private var overridden: Bool {
        guard let preset else { return false }
        return store.hasOverride(definition, preset: preset)
    }

    var body: some View {
        HStack(alignment: .center, spacing: 18) {
            VStack(alignment: .leading, spacing: 3) {
                Text(definition.label).font(.callout.weight(.semibold))
                Text(definition.help).font(.caption).foregroundStyle(.secondary)
                if preset != nil && definition.presetScoped {
                    Text(overridden ? "Preset override" : "Inherited")
                        .font(.caption2.monospaced())
                        .foregroundStyle(.tertiary)
                }
            }
            Spacer(minLength: 20)
            editor.frame(width: 180)
            if let preset, definition.presetScoped, overridden {
                Button("Reset") { store.resetPreset(definition, preset: preset) }
            }
        }
        .padding(.vertical, 8)
    }

    @ViewBuilder
    private var editor: some View {
        switch definition.kind {
        case .boolean:
            Toggle("", isOn: boolBinding)
                .labelsHidden()
                .frame(maxWidth: .infinity, alignment: .trailing)
        case .choice(let choices):
            Picker("", selection: textBinding) {
                ForEach(choices, id: \.self) { Text($0).tag($0) }
            }
            .labelsHidden()
        case .text, .integer, .decimal:
            TextField("", text: textBinding)
                .textFieldStyle(.roundedBorder)
                .font(.body.monospaced())
                .multilineTextAlignment(.trailing)
        }
    }

    private var textBinding: Binding<String> {
        Binding(
            get: { value },
            set: { newValue in
                if let preset { store.setPreset(definition, preset: preset, value: newValue) }
                else { store.setGlobal(definition, value: newValue) }
            }
        )
    }

    private var boolBinding: Binding<Bool> {
        Binding(
            get: { ["true", "1", "yes", "on"].contains(value.lowercased()) },
            set: { textBinding.wrappedValue = $0 ? "true" : "false" }
        )
    }
}
