#!/usr/bin/swift

import AppKit
import Darwin
import Foundation

guard CommandLine.arguments.count == 2 else {
    fputs("Usage: Generate-MacAppIcon.swift <AppIcon.iconset>\n", stderr)
    exit(2)
}

let output = URL(fileURLWithPath: CommandLine.arguments[1], isDirectory: true)
try FileManager.default.createDirectory(at: output, withIntermediateDirectories: true)

let variants: [(name: String, pixels: Int)] = [
    ("icon_16x16.png", 16),
    ("icon_16x16@2x.png", 32),
    ("icon_32x32.png", 32),
    ("icon_32x32@2x.png", 64),
    ("icon_128x128.png", 128),
    ("icon_128x128@2x.png", 256),
    ("icon_256x256.png", 256),
    ("icon_256x256@2x.png", 512),
    ("icon_512x512.png", 512),
    ("icon_512x512@2x.png", 1024),
]

func point(_ x: CGFloat, _ y: CGFloat, scale: CGFloat) -> NSPoint {
    NSPoint(x: x * scale, y: y * scale)
}

func drawIcon(pixels: Int) throws -> Data {
    guard let bitmap = NSBitmapImageRep(
        bitmapDataPlanes: nil,
        pixelsWide: pixels,
        pixelsHigh: pixels,
        bitsPerSample: 8,
        samplesPerPixel: 4,
        hasAlpha: true,
        isPlanar: false,
        colorSpaceName: .deviceRGB,
        bytesPerRow: 0,
        bitsPerPixel: 0
    ) else {
        throw CocoaError(.fileWriteUnknown)
    }

    let previous = NSGraphicsContext.current
    NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: bitmap)
    defer { NSGraphicsContext.current = previous }

    let scale = CGFloat(pixels) / 1024
    NSColor.clear.setFill()
    NSRect(x: 0, y: 0, width: CGFloat(pixels), height: CGFloat(pixels)).fill()

    NSColor(calibratedWhite: 0.12, alpha: 1).setFill()
    NSBezierPath(
        roundedRect: NSRect(x: 34 * scale, y: 34 * scale, width: 956 * scale, height: 956 * scale),
        xRadius: 214 * scale,
        yRadius: 214 * scale
    ).fill()

    let path = NSBezierPath()
    path.lineWidth = max(1.6, 66 * scale)
    path.lineCapStyle = .round
    path.lineJoinStyle = .round

    path.move(to: point(244, 690, scale: scale))
    path.line(to: point(470, 690, scale: scale))
    path.line(to: point(633, 520, scale: scale))
    path.line(to: point(765, 520, scale: scale))

    path.move(to: point(244, 520, scale: scale))
    path.line(to: point(765, 520, scale: scale))

    path.move(to: point(244, 350, scale: scale))
    path.line(to: point(470, 350, scale: scale))
    path.line(to: point(633, 520, scale: scale))

    NSColor(calibratedWhite: 0.96, alpha: 1).setStroke()
    path.stroke()

    NSColor(calibratedRed: 0.23, green: 0.78, blue: 0.42, alpha: 1).setFill()
    NSBezierPath(
        ovalIn: NSRect(x: 742 * scale, y: 467 * scale, width: 106 * scale, height: 106 * scale)
    ).fill()

    guard let png = bitmap.representation(using: .png, properties: [:]) else {
        throw CocoaError(.fileWriteUnknown)
    }
    return png
}

for variant in variants {
    try drawIcon(pixels: variant.pixels)
        .write(to: output.appendingPathComponent(variant.name), options: .atomic)
}
