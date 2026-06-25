using System.Security.Cryptography;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public sealed class TaskWorkspaceManager(string? root = null) : ITaskWorkspaceManager
{
    private readonly string _root = root ?? AppPaths.Workspaces;
    public async Task<TaskWorkspace> CreateAsync(IReadOnlyList<string> paths, CancellationToken token)
    {
        var id = Guid.NewGuid();
        var workspaceRoot = Path.Combine(_root, id.ToString("N"));
        Directory.CreateDirectory(workspaceRoot);
        var snapshots = new List<SourceFileSnapshot>();
        try
        {
            foreach (var source in paths)
            {
                token.ThrowIfCancellationRequested();
                var before = new FileInfo(source);
                var beforeLength = before.Length; var beforeWrite = before.LastWriteTimeUtc;
                var destination = Path.Combine(workspaceRoot, $"{snapshots.Count + 1}-{Path.GetFileName(source)}");
                await using var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                await using var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
                await input.CopyToAsync(output, token);
                var after = new FileInfo(source);
                if (beforeLength != after.Length || beforeWrite != after.LastWriteTimeUtc)
                    throw new SourceChangedException(source);
                await using var copied = File.OpenRead(destination);
                snapshots.Add(new(source, destination, Path.GetFileName(source), beforeLength, beforeWrite,
                    Convert.ToHexString(await SHA256.HashDataAsync(copied, token))));
            }
            return new(id, workspaceRoot, snapshots);
        }
        catch { if (Directory.Exists(workspaceRoot)) Directory.Delete(workspaceRoot, true); throw; }
    }
    public Task CleanupAsync(TaskWorkspace workspace)
    {
        if (Directory.Exists(workspace.Root)) Directory.Delete(workspace.Root, true);
        return Task.CompletedTask;
    }
}
