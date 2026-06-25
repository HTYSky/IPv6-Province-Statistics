namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record TaskWorkspace(Guid TaskId, string Root, IReadOnlyList<SourceFileSnapshot> Sources);
public sealed class SourceChangedException(string path) : IOException($"Source changed while copying: {path}");
