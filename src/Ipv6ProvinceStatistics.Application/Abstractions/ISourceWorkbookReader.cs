using Ipv6ProvinceStatistics.Application.Models;

namespace Ipv6ProvinceStatistics.Application.Abstractions;

public interface ISourceWorkbookReader
{
    Task<SourceReadResult> ReadAsync(
        string path,
        SourceWorkbookKind kind,
        CancellationToken cancellationToken);
}
