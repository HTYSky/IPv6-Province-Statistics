using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.FileSystem;

public sealed class TaskWorkspaceManagerTests
{
    [Fact]
    public async Task Workspace_copies_sources_and_records_sha256()
    {
        var root = TempDirectories.Next();
        var source = Path.Combine(root, "1.xlsx");
        await File.WriteAllTextAsync(source, "source-bytes");
        var manager = new TaskWorkspaceManager(Path.Combine(root, "local-app-data"));
        var workspace = await manager.CreateAsync([source], CancellationToken.None);
        Assert.Single(workspace.Sources);
        Assert.True(File.Exists(workspace.Sources[0].SnapshotPath));
        Assert.Equal(64, workspace.Sources[0].Sha256.Length);
        await manager.CleanupAsync(workspace);
        Assert.False(Directory.Exists(workspace.Root));
    }
}
