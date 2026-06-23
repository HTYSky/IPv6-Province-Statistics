using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Domain.Reporting;

public sealed record MonthResolution(ReportMonth? Month, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid => Month is not null && Issues.Count == 0;
}

public static class MonthResolver
{
    public static MonthResolution Resolve(
        IEnumerable<MonthMarker> markers,
        ReportMonth? selected)
    {
        MonthMarker[] markerList = markers.ToArray();
        (int Year, int Month)[] fullMonths = markerList
            .Where(marker => marker.Year.HasValue)
            .Select(marker => (marker.Year!.Value, marker.Month))
            .Distinct()
            .ToArray();

        if (fullMonths.Length > 1)
        {
            return Conflict("Sources contain more than one explicit report month.");
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
                return Conflict("A month-only marker conflicts with the explicit report month.");
            }

            if (selected.HasValue && selected.Value != resolved)
            {
                return Conflict("The selected report month conflicts with the explicit report month.");
            }

            return Valid(resolved);
        }

        if (partialMonths.Length > 1)
        {
            return Conflict("Sources contain conflicting month-only markers.");
        }

        if (partialMonths.Length == 1)
        {
            int partialMonth = partialMonths[0];
            if (!selected.HasValue)
            {
                return Invalid("MONTH_YEAR_MISSING", "A report month was found, but its year is missing.");
            }

            if (selected.Value.Month != partialMonth)
            {
                return Conflict("The selected report month conflicts with the source month.");
            }

            return Valid(selected.Value);
        }

        return selected.HasValue
            ? Valid(selected.Value)
            : Invalid("MONTH_MISSING", "No report month was found or selected.");
    }

    private static MonthResolution Valid(ReportMonth month) => new(month, []);

    private static MonthResolution Conflict(string message) => Invalid("MONTH_CONFLICT", message);

    private static MonthResolution Invalid(string code, string message) =>
        new(null, [new ValidationIssue(code, message)]);
}
