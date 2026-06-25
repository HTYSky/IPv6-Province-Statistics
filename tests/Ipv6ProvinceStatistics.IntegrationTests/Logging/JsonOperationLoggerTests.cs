using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Logging;

public sealed class JsonOperationLoggerTests
{
    [Fact]
    public async Task Logger_never_serializes_business_values_and_keeps_30_files()
    {
        var root = TempDirectories.Next();
        var logger = new JsonOperationLogger(root);
        for (var index = 0; index < 31; index++)
            await logger.WriteAsync(new OperationLogEntry(Guid.NewGuid(), DateTimeOffset.UtcNow, "1.0.0",
                "Succeeded", 2026, 5, [new("1.xlsx", 100, DateTime.UnixEpoch, "ABC", "Table1")],
                "C:\\output", 31, 1000, []), CancellationToken.None);
        Assert.Equal(30, Directory.GetFiles(root, "*.json").Length);
        Assert.DoesNotContain("CellValue", await File.ReadAllTextAsync(Directory.GetFiles(root).First()));
    }
}
