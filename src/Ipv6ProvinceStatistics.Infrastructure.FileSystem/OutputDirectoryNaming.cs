using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public static class OutputDirectoryNaming
{
    public static string NextAvailable(string parent, ReportMonth month)
    {
        var basePath = Path.Combine(parent, month.FolderName);
        if (!Directory.Exists(basePath)) return basePath;
        for (var suffix = 2; suffix < int.MaxValue; suffix++)
        {
            var candidate = $"{basePath} ({suffix})";
            if (!Directory.Exists(candidate)) return candidate;
        }
        throw new IOException("无法分配输出目录名称。");
    }
}
