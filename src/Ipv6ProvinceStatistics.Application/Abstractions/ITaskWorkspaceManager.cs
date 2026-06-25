using Ipv6ProvinceStatistics.Application.Models;
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public interface ITaskWorkspaceManager
{
    Task<TaskWorkspace> CreateAsync(IReadOnlyList<string> sourcePaths, CancellationToken cancellationToken);
    Task CleanupAsync(TaskWorkspace workspace);
}
