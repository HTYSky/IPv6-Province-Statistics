using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Application.Models;

public sealed record SourceReadResult(
    SourceWorkbookKind Kind,
    IReadOnlyDictionary<Province, IReadOnlyDictionary<MetricKey, decimal>> Values,
    IReadOnlyList<ValidationIssue> Issues);

public sealed record ProvinceAssemblyResult(
    IReadOnlyDictionary<Province, ProvinceReportInput> Reports,
    IReadOnlyList<ValidationIssue> Issues);
