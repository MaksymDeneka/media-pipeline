import AppKit
import AVFoundation
import ImageIO
import SwiftUI

@MainActor
final class MediaThumbnailLoader {
    static let shared = MediaThumbnailLoader()
    private let cache = NSCache<NSString, NSImage>()
    private var pending: [Request: Task<CGImage?, Never>] = [:]

    struct Request: Hashable, Sendable {
        let url: URL
        let maxPixelSize: Int

        var cacheKey: NSString { "\(url.absoluteString)#\(maxPixelSize)" as NSString }
    }

    init() {
        cache.countLimit = 256
        cache.totalCostLimit = 64 * 1024 * 1024
    }

    func image(for url: URL, maxPixelSize: Int = 144) async -> NSImage? {
        let request = Request(url: url, maxPixelSize: maxPixelSize)
        if let image = cache.object(forKey: request.cacheKey) { return image }
        guard let bitmap = await loadBitmap(for: request) else { return nil }
        if let image = cache.object(forKey: request.cacheKey) { return image }
        let image = NSImage(cgImage: bitmap, size: .zero)
        cache.setObject(image, forKey: request.cacheKey, cost: bitmap.bytesPerRow * bitmap.height)
        return image
    }

    private func loadBitmap(for request: Request) async -> CGImage? {
        if let task = pending[request] { return await task.value }
        let task = Task.detached(priority: .utility) { await Self.generate(request) }
        pending[request] = task
        defer { pending[request] = nil }
        return await task.value
    }

    nonisolated private static func generate(_ request: Request) async -> CGImage? {
        if let source = CGImageSourceCreateWithURL(request.url as CFURL, nil),
           let image = CGImageSourceCreateThumbnailAtIndex(source, 0, [
               kCGImageSourceCreateThumbnailFromImageAlways: true,
               kCGImageSourceCreateThumbnailWithTransform: true,
               kCGImageSourceThumbnailMaxPixelSize: request.maxPixelSize,
           ] as CFDictionary) {
            return image
        }
        let generator = AVAssetImageGenerator(asset: AVURLAsset(url: request.url))
        generator.appliesPreferredTrackTransform = true
        generator.maximumSize = CGSize(width: request.maxPixelSize, height: request.maxPixelSize)
        return await withCheckedContinuation { continuation in
            generator.generateCGImagesAsynchronously(forTimes: [NSValue(time: .zero)]) {
                _, image, _, _, _ in
                continuation.resume(returning: image)
            }
        }
    }
}

struct MediaThumbnail: View {
    enum Size {
        case list, tile
    }

    let url: URL
    let isImage: Bool
    var size: Size = .list
    @Environment(\.displayScale) private var displayScale
    @State private var image: NSImage?

    var body: some View {
        let maxPixelSize = size == .tile ? Int(260 * displayScale) : 144
        let request = MediaThumbnailLoader.Request(url: url, maxPixelSize: maxPixelSize)
        ZStack {
            Color.black.opacity(0.2)
            if let image {
                Image(nsImage: image).resizable().scaledToFit()
            } else {
                Image(systemName: isImage ? "photo" : "film")
                    .font(size == .tile ? .largeTitle : .body)
                    .foregroundStyle(.secondary)
            }
        }
        .frame(width: size == .list ? 72 : nil, height: size == .list ? 48 : 176)
        .clipShape(RoundedRectangle(cornerRadius: size == .list ? 5 : 0))
        .accessibilityLabel(isImage ? "Image preview" : "Video preview")
        .task(id: request) {
            let thumbnail = await MediaThumbnailLoader.shared.image(for: url, maxPixelSize: maxPixelSize)
            if !Task.isCancelled { image = thumbnail }
        }
    }
}
