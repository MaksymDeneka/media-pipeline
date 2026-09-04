using MediaPipeline.Core.Configuration;

namespace MediaPipeline.Core.Tools;

public sealed record VideoEncoder(string Name, string Description);

public static class VideoEncoderSelector
{
    public static IReadOnlyList<string> Candidates(VideoOptions options)
    {
        // Heatup approach: capability decides – VideoToolbox is probed with a real
        // test encode where allowed, otherwise CPU libx264. NVENC/AMF paths were
        // removed: V2 uses a precomputed bitrate ladder with two-pass libx264
        // (slow) or single-pass VideoToolbox, no CRF/CQ/QP tuning.
        // PreferVideoToolbox=false is an explicit opt-out (forces libx264).
        var candidates = new List<string>();
        if (options.PreferVideoToolbox)
        {
            candidates.Add("h264_videotoolbox");
        }

        candidates.Add("libx264");
        return candidates;
    }

    public static async Task<VideoEncoder> SelectAsync(
        Toolchain tools,
        VideoOptions options,
        CancellationToken cancellationToken = default)
    {
        foreach (var candidate in Candidates(options))
        {
            if (await ProbeAsync(tools.FFmpeg, candidate, cancellationToken))
            {
                return new VideoEncoder(candidate, Description(candidate));
            }
        }

        throw new InvalidOperationException("FFmpeg has no usable H.264 encoder.");
    }

    private static async Task<bool> ProbeAsync(
        string ffmpeg,
        string encoder,
        CancellationToken cancellationToken)
    {
        var result = await ProcessRunner.RunAsync(
            ffmpeg,
            [
                "-hide_banner",
                "-loglevel",
                "error",
                "-f",
                "lavfi",
                "-i",
                "color=c=black:s=256x256:r=1:d=1",
                "-frames:v",
                "1",
                "-c:v",
                encoder,
                "-pix_fmt",
                "yuv420p",
                "-f",
                "null",
                "-",
            ],
            cancellationToken);

        return result.Succeeded;
    }

    private static string Description(string encoder) => encoder switch
    {
        "h264_videotoolbox" => "Apple VideoToolbox",
        "libx264" => "CPU libx264",
        _ => encoder,
    };
}
