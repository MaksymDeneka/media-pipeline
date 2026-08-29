import Foundation
import ServiceManagement

@MainActor
final class LoginItemController: ObservableObject {
    @Published private(set) var enabled = false
    @Published var errorMessage: String?

    init() {
        refresh()
    }

    func refresh() {
        enabled = SMAppService.mainApp.status == .enabled
    }

    func setEnabled(_ value: Bool) {
        do {
            if value {
                try SMAppService.mainApp.register()
            } else {
                try SMAppService.mainApp.unregister()
            }
            refresh()
            errorMessage = value && SMAppService.mainApp.status == .requiresApproval
                ? "Allow Media Pipeline in System Settings → General → Login Items."
                : nil
        } catch {
            refresh()
            errorMessage = error.localizedDescription
        }
    }
}
