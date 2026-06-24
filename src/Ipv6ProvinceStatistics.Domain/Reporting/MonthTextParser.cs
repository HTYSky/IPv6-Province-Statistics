using System.Globalization;
using System.Text.RegularExpressions;

namespace Ipv6ProvinceStatistics.Domain.Reporting;

public static class MonthTextParser
{
    private static readonly Regex FullChineseMonthShapePattern = new(
        @"(?<!\p{Nd})\p{Nd}{4}年\p{Nd}{1,2}月",
        RegexOptions.CultureInvariant);

    private static readonly Regex FullChineseMonthPattern = new(
        @"(?<!\p{Nd})(?<year>20[0-9]{2})年(?<month>0?[1-9]|1[0-2])月",
        RegexOptions.CultureInvariant);

    private static readonly Regex CompactMonthPattern = new(
        @"(?<!\p{Nd})(?<year>20[0-9]{2})(?<month>0[1-9]|1[0-2])(?!\p{Nd})",
        RegexOptions.CultureInvariant);

    private static readonly Regex PartialMonthPattern = new(
        @"(?<!\p{Nd})(?<month>0?[1-9]|1[0-2])月",
        RegexOptions.CultureInvariant);

    public static IReadOnlyList<MonthMarker> Extract(string source, string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return Array.AsReadOnly(Array.Empty<MonthMarker>());
        }

        Match[] fullChineseShapes = FullChineseMonthShapePattern.Matches(text).Cast<Match>().ToArray();
        var fullCandidates = new List<MarkerCandidate>();
        AddFullCandidates(fullCandidates, FullChineseMonthPattern, text);
        AddFullCandidates(fullCandidates, CompactMonthPattern, text);

        var candidates = new List<MarkerCandidate>(fullCandidates);
        foreach (Match match in PartialMonthPattern.Matches(text))
        {
            if (fullChineseShapes.Any(shape =>
                    Overlaps(shape.Index, shape.Length, match.Index, match.Length)))
            {
                continue;
            }

            candidates.Add(new MarkerCandidate(
                match.Index,
                match.Length,
                null,
                Parse(match.Groups["month"])));
        }

        var seen = new HashSet<(int? Year, int Month)>();
        var markers = new List<MonthMarker>();

        foreach (MarkerCandidate candidate in candidates.OrderBy(candidate => candidate.Index))
        {
            if (seen.Add((candidate.Year, candidate.Month)))
            {
                markers.Add(new MonthMarker(candidate.Year, candidate.Month, source));
            }
        }

        return Array.AsReadOnly(markers.ToArray());
    }

    private static void AddFullCandidates(
        ICollection<MarkerCandidate> candidates,
        Regex pattern,
        string text)
    {
        foreach (Match match in pattern.Matches(text))
        {
            candidates.Add(new MarkerCandidate(
                match.Index,
                match.Length,
                Parse(match.Groups["year"]),
                Parse(match.Groups["month"])));
        }
    }

    private static int Parse(Group group) =>
        int.Parse(group.Value, NumberStyles.None, CultureInfo.InvariantCulture);

    private static bool Overlaps(int leftIndex, int leftLength, int rightIndex, int rightLength) =>
        leftIndex < rightIndex + rightLength && rightIndex < leftIndex + leftLength;

    private readonly record struct MarkerCandidate(int Index, int Length, int? Year, int Month);
}
