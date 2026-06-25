using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public interface IReportProcessingService
{
    Task<PreflightResult> PreflightAsync(IReadOnlyList<string> paths, ReportMonth? selectedMonth,
        IProgress<ProcessingProgress>? progress, CancellationToken token);
    Task<GenerationResult> GenerateAsync(PreparedBatch batch, string outputParent,
        IProgress<ProcessingProgress>? progress, CancellationToken token);
    Task DiscardAsync(PreparedBatch batch);
}
