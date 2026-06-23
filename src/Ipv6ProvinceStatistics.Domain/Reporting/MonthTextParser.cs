using System.Globalization;
using System.Text.RegularExpressions;

namespace Ipv6ProvinceStatistics.Domain.Reporting;

public static class MonthTextParser
{
    private static readonly Regex FullChineseMonthPattern = new(
        @"(?<!\d)(?<year>20\d{2})年(?<month>0?[1-9]|1[0-2])月",
        RegexOptions.CultureInvariant);

    private static readonly Regex CompactMonthPattern = new(
        @"(?<!\d)(?<year>20\d{2})(?<month>0[1-9]|1[0-2])(?!\d)",
        RegexOptions.CultureInvariant);

    private static readonly Regex PartialMonthPattern = new(
        @"(?<!\d)(?<month>0?[1-9]|1[0-2])月",
        RegexOptions.CultureInvariant);

    public static IReadOnlyList<MonthMarker> Extract(string source, string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return [];
        }

        var fullCandidates = new List<MarkerCandidate>();
        AddFullCandidates(fullCandidates, FullChineseMonthPattern, text);
        AddFullCandidates(fullCandidates, CompactMonthPattern, text);

        var candidates = new List<MarkerCandidate>(fullCandidates);
        foreach (Match match in PartialMonthPattern.Matches(text))
        {
            if (fullCandidates.Any(candidate => candidate.Overlaps(match.Index, match.Length)))
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

        return markers;
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

    private readonly record struct MarkerCandidate(int Index, int Length, int? Year, int Month)
    {
        public bool Overlaps(int index, int length) => Index < index + length && index < Index + Length;
    }
}
