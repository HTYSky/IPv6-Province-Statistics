namespace Ipv6ProvinceStatistics.Domain.Validation;

public sealed record ValidationIssue(
    string Code,
    string Message,
    string? FileName = null,
    string? Sheet = null,
    string? Province = null,
    string? Cell = null);
