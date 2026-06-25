using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public interface IOutputTransaction
{
    Task<string> CreateStagingAsync(string outputParent, Guid taskId);
    Task<string> PublishAsync(string stagingPath, string outputParent, ReportMonth month, CancellationToken token);
    Task CleanupAsync(string stagingPath);
}
