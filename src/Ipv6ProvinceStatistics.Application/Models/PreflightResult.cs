using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record PreflightResult(PreparedBatch? Batch, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid => Batch is not null && Issues.Count == 0;
}
