using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public sealed class OutputTransaction : IOutputTransaction
{
    public Task<string> CreateStagingAsync(string parent, Guid id)
    {
        Directory.CreateDirectory(parent);
        var path = Path.Combine(parent, $".ipv6stats-{id:N}");
        Directory.CreateDirectory(path);
        if (OperatingSystem.IsWindows()) new DirectoryInfo(path).Attributes |= FileAttributes.Hidden;
        return Task.FromResult(path);
    }
    public Task<string> PublishAsync(string staging, string parent, ReportMonth month, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var expected = ProvinceCatalog.All.Select(province =>
            $"{province.Name}-{month.FileSuffix}.xlsx").ToHashSet(StringComparer.OrdinalIgnoreCase);
        var actual = Directory.GetFiles(staging, "*.xlsx").Select(p => Path.GetFileName(p)!)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!actual.SetEquals(expected))
            throw new InvalidDataException("Staging directory must contain the exact 31 province workbook names.");
        var target = OutputDirectoryNaming.NextAvailable(parent, month);
        Directory.Move(staging, target);
        if (OperatingSystem.IsWindows()) new DirectoryInfo(target).Attributes &= ~FileAttributes.Hidden;
        return Task.FromResult(target);
    }
    public Task CleanupAsync(string staging)
    {
        if (Directory.Exists(staging)) Directory.Delete(staging, true);
        return Task.CompletedTask;
    }
}
