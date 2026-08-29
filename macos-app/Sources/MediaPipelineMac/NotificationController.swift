import Foundation
import UserNotifications

@MainActor
final class NotificationController {
    private var prepared = false

    func prepare() {
        guard !prepared else { return }
        prepared = true
        UNUserNotificationCenter.current().requestAuthorization(options: [.alert, .sound]) { _, _ in }
    }

    func processingFailed(_ job: JobSnapshot) {
        let content = UNMutableNotificationContent()
        content.title = "Processing failed"
        content.body = "\(job.lane): \(job.detail)"
        content.sound = .default
        let request = UNNotificationRequest(
            identifier: "media-pipeline-failure-\(job.id)",
            content: content,
            trigger: nil
        )
        UNUserNotificationCenter.current().add(request)
    }
}
