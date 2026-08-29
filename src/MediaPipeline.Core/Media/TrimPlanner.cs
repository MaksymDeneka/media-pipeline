using System.Security.Cryptography;

namespace MediaPipeline.Core.Media;

public sealed record TrimRange(bool CanTrim, int MinMs, int MaxMs, string Reason);

public static class TrimPlanner
{
    public static TrimRange GetRange(double durationSeconds, int configuredMinMs, int configuredMaxMs)
    {
        var durationMs = (int)Math.Floor(durationSeconds * 1000);
        if (durationMs < 500)
        {
            return new TrimRange(false, 0, 0, "video is shorter than 500 ms");
        }

        if (durationMs < 2000)
        {
            var safeMax = Math.Min(100, (int)Math.Floor(durationMs * 0.10));
            return safeMax < 10
                ? new TrimRange(false, 0, 0, "video is too short for safe trimming")
                : new TrimRange(true, 10, safeMax, "short video safety range");
        }

        var safeConfiguredMax = Math.Min(configuredMaxMs, durationMs - 1000);
        return safeConfiguredMax < configuredMinMs
            ? new TrimRange(false, 0, 0, "configured trim range would make output too short")
            : new TrimRange(true, configuredMinMs, safeConfiguredMax, "configured trim range");
    }

    public static int PickMilliseconds(TrimRange range, ISet<int> usedValues, int copyCount)
    {
        if (!range.CanTrim)
        {
            return 0;
        }

        var rangeSize = range.MaxMs - range.MinMs + 1;
        var mustBeUnique = rangeSize >= copyCount;
        int value;
        var attempts = 0;

        do
        {
            value = RandomNumberGenerator.GetInt32(range.MinMs, range.MaxMs + 1);
            attempts++;
        }
        while (mustBeUnique && usedValues.Contains(value) && attempts < 50);

        usedValues.Add(value);
        return value;
    }
}
