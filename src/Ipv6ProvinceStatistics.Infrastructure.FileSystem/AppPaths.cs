namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public static class AppPaths
{
    public static string Root => Path.Combine(Environment.GetFolderPath(
        Environment.SpecialFolder.LocalApplicationData), "IPv6ProvinceStatistics");
    public static string Workspaces => Path.Combine(Root, "Workspaces");
    public static string Logs => Path.Combine(Root, "Logs");
}
