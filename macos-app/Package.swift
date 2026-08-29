// swift-tools-version: 5.9

import PackageDescription

let package = Package(
    name: "MediaPipelineMac",
    platforms: [.macOS(.v13)],
    products: [
        .executable(name: "MediaPipelineMac", targets: ["MediaPipelineMac"]),
    ],
    targets: [
        .executableTarget(
            name: "MediaPipelineMac",
            resources: [.process("Resources")]
        ),
        .testTarget(
            name: "MediaPipelineMacTests",
            dependencies: ["MediaPipelineMac"]
        ),
    ]
)
