namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record SourceFileSnapshot(string OriginalPath, string SnapshotPath, string FileName,
    long Length, DateTime LastWriteTimeUtc, string Sha256);
