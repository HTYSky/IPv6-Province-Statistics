using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.FileSystem;

public sealed class OutputTransactionTests
{
    [Fact]
    public async Task Output_publishes_31_files_to_numbered_same_volume_directory()
    {
        var parent = TempDirectories.Next();
        Directory.CreateDirectory(Path.Combine(parent, "2026年05月统计结果"));
        var transaction = new OutputTransaction();
        var staging = await transaction.CreateStagingAsync(parent, Guid.Parse("11111111-1111-1111-1111-111111111111"));
        foreach (var province in ProvinceCatalog.All)
            await File.WriteAllTextAsync(Path.Combine(staging, $"{province.Name}-2026年05月.xlsx"), province.Name);
        var published = await transaction.PublishAsync(staging, parent, new ReportMonth(2026, 5), CancellationToken.None);
        Assert.EndsWith("2026年05月统计结果 (2)", published);
        Assert.Equal(31, Directory.GetFiles(published, "*.xlsx").Length);
    }
}
