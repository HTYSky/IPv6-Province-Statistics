namespace Ipv6ProvinceStatistics.Application.Abstractions;
public sealed record LogSource(string FileName, long Length, DateTime LastWriteTimeUtc, string Sha256,
    string? IdentifiedKind);
public sealed record OperationLogEntry(Guid TaskId, DateTimeOffset Timestamp, string Version, string Status,
    int? Year, int? Month, IReadOnlyList<LogSource> Sources, string? OutputDirectory,
    int OutputCount, long DurationMilliseconds, IReadOnlyList<string> Errors);
public interface IOperationLogger
{
    Task WriteAsync(OperationLogEntry entry, CancellationToken cancellationToken);
}
