using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Application.Models;

public sealed record WorkbookInspection(
    string Path,
    IReadOnlyList<SourceWorkbookKind> MatchingKinds,
    IReadOnlyList<MonthMarker> MonthMarkers,
    IReadOnlyList<ValidationIssue> Issues);
