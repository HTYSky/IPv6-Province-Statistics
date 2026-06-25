namespace Ipv6ProvinceStatistics.UnitTests.Services;

public sealed class ReportProcessingServiceTests
{
    [Fact]
    public async Task Service_preflights_and_generates_exactly_31_files()
    {
        using var h = ProcessingHarness.Create(["1.xlsx","4.xlsx","5.xlsx","8.xlsx"]);
        var pre = await h.Service.PreflightAsync(h.Paths, null, null, CancellationToken.None);
        Assert.True(pre.IsValid);
        var result = await h.Service.GenerateAsync(pre.Batch!, h.OutputParent, null, CancellationToken.None);
        Assert.True(result.Succeeded);
        Assert.Equal(31, result.OutputCount);
        Assert.Equal(31, Directory.GetFiles(result.OutputDirectory!, "*.xlsx").Length);
        Assert.Contains("北京-2026年05月.xlsx", Directory.GetFiles(result.OutputDirectory!).Select(Path.GetFileName));
    }

    [Fact]
    public async Task Service_rejects_duplicate_kind_and_cleans_workspace()
    {
        using var h = ProcessingHarness.Create(["1-a.xlsx","1-b.xlsx","4.xlsx","5.xlsx"]);
        var result = await h.Service.PreflightAsync(h.Paths, null, null, CancellationToken.None);
        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, x => x.Code == "SOURCE_KIND_COUNT");
        Assert.True(h.Workspace.CleanupCalled);
        Assert.Empty(Directory.GetDirectories(h.OutputParent));
    }
}
