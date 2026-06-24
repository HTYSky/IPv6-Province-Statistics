using Ipv6ProvinceStatistics.Application.Models;

namespace Ipv6ProvinceStatistics.Application.Abstractions;

public interface IWorkbookInspector
{
    Task<WorkbookInspection> InspectAsync(string path, CancellationToken cancellationToken);
}
