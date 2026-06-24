using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Domain.Reporting;

public sealed record MonthResolution(ReportMonth? Month, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid =>
        Month.HasValue &&
        Month.Value.Year is >= 2000 and <= 9999 &&
        Month.Value.Month is >= 1 and <= 12 &&
        Issues.Count == 0;
}

public static class MonthResolver
{
    public static MonthResolution Resolve(
        IEnumerable<MonthMarker> markers,
        ReportMonth? selected)
    {
        MonthMarker[] markerList = markers.ToArray();
        MonthMarker[] invalidMarkers = markerList
            .Where(marker =>
                marker.Month is < 1 or > 12 || marker.Year is < 2000 or > 2099)
            .ToArray();

        if (invalidMarkers.Length > 0)
        {
            return Invalid(
                "MONTH_INVALID",
                WithSources("A source contains an invalid report month marker.", invalidMarkers));
        }

        if (selected.HasValue && !IsValid(selected.Value))
        {
            return Invalid(
                "MONTH_INVALID",
                WithSources("The selected report month is invalid.", markerList));
        }

        (int Year, int Month)[] fullMonths = markerList
            .Where(marker => marker.Year.HasValue)
            .Select(marker => (marker.Year!.Value, marker.Month))
            .Distinct()
            .ToArray();

        if (fullMonths.Length > 1)
        {
            return Conflict(
                "Sources contain more than one explicit report month.",
                markerList.Where(marker => marker.Year.HasValue));
        }

        int[] partialMonths = markerList
            .Where(marker => !marker.Year.HasValue)
            .Select(marker => marker.Month)
            .Distinct()
            .ToArray();

        if (fullMonths.Length == 1)
        {
            (int year, int month) = fullMonths[0];
            var resolved = new ReportMonth(year, month);

            if (partialMonths.Any(partialMonth => partialMonth != month))
            {
                return Conflict(
                    "A month-only marker conflicts with the explicit report month.",
                    markerList);
            }

            if (selected.HasValue && selected.Value != resolved)
            {
                return Conflict(
                    "The selected report month conflicts with the explicit report month.",
                    markerList.Where(marker => marker.Year.HasValue));
            }

            return Valid(resolved);
        }

        if (partialMonths.Length > 1)
        {
            return Conflict(
                "Sources contain conflicting month-only markers.",
                markerList.Where(marker => !marker.Year.HasValue));
        }

        if (partialMonths.Length == 1)
        {
            int partialMonth = partialMonths[0];
            if (!selected.HasValue)
            {
                return Invalid(
                    "MONTH_YEAR_MISSING",
                    WithSources(
                        "A report month was found, but its year is missing.",
                        markerList.Where(marker => !marker.Year.HasValue)));
            }

            if (selected.Value.Month != partialMonth)
            {
                return Conflict(
                    "The selected report month conflicts with the source month.",
                    markerList.Where(marker => !marker.Year.HasValue));
            }

            return Valid(selected.Value);
        }

        return selected.HasValue
            ? Valid(selected.Value)
            : Invalid("MONTH_MISSING", "No report month was found or selected.");
    }

    private static MonthResolution Valid(ReportMonth month) => new(month, []);

    private static MonthResolution Conflict(
        string message,
        IEnumerable<MonthMarker> markers) =>
        Invalid("MONTH_CONFLICT", WithSources(message, markers));

    private static bool IsValid(ReportMonth month) =>
        month.Year is >= 2000 and <= 9999 && month.Month is >= 1 and <= 12;

    private static string WithSources(string message, IEnumerable<MonthMarker> markers)
    {
        string sources = string.Join(
            ", ",
            markers
                .Select(marker => marker.Source)
                .Where(source => !string.IsNullOrWhiteSpace(source))
                .Distinct(StringComparer.Ordinal));

        return sources.Length == 0 ? message : $"{message} Sources: {sources}.";
    }

    private static MonthResolution Invalid(string code, string message) =>
        new(null, [new ValidationIssue(code, message)]);
}
