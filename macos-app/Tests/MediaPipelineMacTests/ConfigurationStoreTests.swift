import XCTest
@testable import MediaPipelineMac

final class ConfigurationStoreTests: XCTestCase {
    func testDecorativeSectionsRemainGlobal() {
        let document = IniDocument(contents: """
        [Video]
        Crf = 24
        [Images]
        JpegQuality = 4
        [preset photos]
        VideoCopies = 0
        ImageCopies = 12
        """)

        XCTAssertEqual(document.globals()["Crf"], "24")
        XCTAssertEqual(document.globals()["JpegQuality"], "4")
        XCTAssertNil(document.globals()["ImageCopies"])
        XCTAssertEqual(document.presets().values["photos"]?["ImageCopies"], "12")
    }

    func testPresetOverrideCanBeAddedAndRemoved() {
        var document = IniDocument(contents: """
        [Video]
        MaxWidth = 1080
        [preset clips]
        VideoCopies = 2
        """)

        document.setPreset("clips", key: "MaxWidth", value: "720")
        XCTAssertEqual(document.presets().values["clips"]?["MaxWidth"], "720")

        document.removePresetValue("clips", key: "MaxWidth")
        XCTAssertNil(document.presets().values["clips"]?["MaxWidth"])
        XCTAssertEqual(document.globals()["MaxWidth"], "1080")
    }

    func testGlobalValueIsInsertedIntoItsSection() {
        var document = IniDocument(contents: """
        [Video]
        Crf = 24
        [preset clips]
        VideoCopies = 2
        """)

        document.setGlobal("MaxWidth", value: "720", section: "Video")
        XCTAssertEqual(document.globals()["MaxWidth"], "720")
        XCTAssertNil(document.presets().values["clips"]?["MaxWidth"])
    }

    func testHomePathExpansion() {
        let url = AppPaths.expandPath("~/MediaPipeline")
        XCTAssertTrue(url.path.hasSuffix("/MediaPipeline"))
        XCTAssertFalse(url.path.contains("~"))
        XCTAssertTrue(AppPaths.isWindowsAbsolutePath("D:\\MediaPipeline"))
        XCTAssertFalse(AppPaths.isWindowsAbsolutePath("~/MediaPipeline"))
    }

    func testInlineCommentsAreNotPartOfValues() {
        var document = IniDocument(contents: """
        [Video]
        Crf = 24 ; normal quality
        AudioBitrate = 128k # AAC
        [preset clips]
        MaxWidth = 720 ; keep this note
        """)

        XCTAssertEqual(document.globals()["Crf"], "24")
        XCTAssertEqual(document.globals()["AudioBitrate"], "128k")
        document.setGlobal("Crf", value: "26", section: "Video")
        document.setPreset("clips", key: "MaxWidth", value: "480")
        XCTAssertTrue(document.serialized().contains("Crf = 26 ; normal quality"))
        XCTAssertTrue(document.serialized().contains("MaxWidth = 480 ; keep this note"))
    }

    func testQuotedCommentCharactersRemainPartOfEditedValues() {
        var document = IniDocument(contents: """
        [General]
        PipelineRoot = "/Users/me/Media #1" ; keep this note
        """)

        XCTAssertEqual(document.globals()["PipelineRoot"], "/Users/me/Media #1")
        document.setGlobal("PipelineRoot", value: "/Users/me/Media #2", section: "General")

        XCTAssertTrue(document.serialized().contains(
            "PipelineRoot = \"/Users/me/Media #2\" ; keep this note"
        ))
        XCTAssertEqual(document.globals()["PipelineRoot"], "/Users/me/Media #2")
    }

    func testRelativePipelineRootResolvesBesideConfig() throws {
        let base = URL(fileURLWithPath: "/tmp/media-pipeline-config", isDirectory: true)
        let store = ConfigurationStore(configURL: base.appendingPathComponent("config.ini"))
        let definition = SettingCatalog.globals.first { $0.key == "PipelineRoot" }!
        store.setGlobal(definition, value: "relative-root")

        XCTAssertEqual(store.pipelineRoot.path, base.appendingPathComponent("relative-root").path)
    }

    func testPresetLifecyclePreservesOtherSections() {
        var document = IniDocument(contents: """
        [General]
        PipelineRoot = ~/MediaPipeline
        [preset existing]
        VideoCopies = 1
        [Timing]
        PollSeconds = 2
        """)

        document.addPreset("new-preset")
        document.setPreset("new-preset", key: "ImageCopies", value: "7")
        document.removePreset("existing")

        XCTAssertEqual(document.presets().order, ["new-preset"])
        XCTAssertEqual(document.presets().values["new-preset"]?["ImageCopies"], "7")
        XCTAssertEqual(document.globals()["PollSeconds"], "2")
    }
}
