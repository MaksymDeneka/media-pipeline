namespace MediaPipeline.Core.Media;

public sealed record SegmentPlan(int Index, double StartSeconds, double DurationSeconds);

public static class SegmentPlanner
{
    public static IReadOnlyList<SegmentPlan> Plan(
        double durationSeconds,
        int targetSeconds,
        int minimumSeconds)
    {
        var durationMs = (int)Math.Floor(durationSeconds * 1000);
        var targetMs = checked(targetSeconds * 1000);
        var minimumMs = checked(minimumSeconds * 1000);

        if (durationMs <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(durationSeconds), "Duration must be positive.");
        }

        if (targetMs <= 0 || minimumMs <= 0 || minimumMs > targetMs)
        {
            throw new ArgumentOutOfRangeException(
                nameof(minimumSeconds),
                "Segment durations must be positive and the minimum cannot exceed the target.");
        }

        var durations = new List<int>();
        if (durationMs <= targetMs)
        {
            durations.Add(durationMs);
        }
        else
        {
            var fullCount = durationMs / targetMs;
            var remainderMs = durationMs - fullCount * targetMs;
            durations.AddRange(Enumerable.Repeat(targetMs, fullCount));

            if (remainderMs > 0)
            {
                AddRemainder(durations, remainderMs, minimumMs);
            }
        }

        var segments = new List<SegmentPlan>(durations.Count);
        var startMs = 0;
        for (var index = 0; index < durations.Count; index++)
        {
            segments.Add(new SegmentPlan(
                index + 1,
                startMs / 1000.0,
                durations[index] / 1000.0));
            startMs += durations[index];
        }

        return segments;
    }

    private static void AddRemainder(List<int> durations, int remainderMs, int minimumMs)
    {
        if (remainderMs >= minimumMs)
        {
            durations.Add(remainderMs);
            return;
        }

        var neededMs = minimumMs - remainderMs;
        var borrowedMs = 0;
        for (var index = durations.Count - 1; index >= 0 && borrowedMs < neededMs; index--)
        {
            var availableMs = durations[index] - minimumMs;
            if (availableMs <= 0)
            {
                continue;
            }

            var takeMs = Math.Min(availableMs, neededMs - borrowedMs);
            durations[index] -= takeMs;
            borrowedMs += takeMs;
        }

        if (borrowedMs == neededMs)
        {
            durations.Add(remainderMs + borrowedMs);
            return;
        }

        var lastIndex = durations.Count - 1;
        durations[lastIndex] += remainderMs + borrowedMs;
    }
}
