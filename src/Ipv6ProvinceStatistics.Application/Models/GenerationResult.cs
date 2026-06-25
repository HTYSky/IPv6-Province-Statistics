using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record GenerationResult(bool Succeeded, string? OutputDirectory, int OutputCount,
    TimeSpan Duration, IReadOnlyList<ValidationIssue> Issues);
