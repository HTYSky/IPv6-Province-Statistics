using System.Text.Json;
using Ipv6ProvinceStatistics.Application.Abstractions;
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public sealed class JsonOperationLogger(string? root = null) : IOperationLogger
{
    private readonly string _root = root ?? AppPaths.Logs;
    public async Task WriteAsync(OperationLogEntry entry, CancellationToken token)
    {
        Directory.CreateDirectory(_root);
        var path = Path.Combine(_root, $"{entry.Timestamp:yyyyMMdd-HHmmss}-{entry.TaskId:N}.json");
        await File.WriteAllTextAsync(path, JsonSerializer.Serialize(entry,
            new JsonSerializerOptions { WriteIndented = true }), token);
        foreach (var stale in Directory.GetFiles(_root, "*.json").OrderBy(File.GetCreationTimeUtc).Skip(30))
            File.Delete(stale);
    }
}
