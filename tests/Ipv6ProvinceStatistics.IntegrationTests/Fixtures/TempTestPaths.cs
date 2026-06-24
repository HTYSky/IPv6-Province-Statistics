namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

internal static class TempTestPaths
{
    private static readonly string SessionRoot;

    static TempTestPaths()
    {
        SessionRoot = Path.Combine(
            Path.GetTempPath(),
            "Ipv6ProvinceStatistics.Tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(SessionRoot);
        AppDomain.CurrentDomain.ProcessExit += (_, _) => TryDeleteSessionRoot();
    }

    internal static string CreateDirectory()
    {
        string path = Path.Combine(SessionRoot, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    private static void TryDeleteSessionRoot()
    {
        try
        {
            if (Directory.Exists(SessionRoot))
            {
                Directory.Delete(SessionRoot, recursive: true);
            }
        }
        catch
        {
            // Process-exit cleanup is best effort only.
        }
    }
}

internal static class TempDirectories
{
    public static string Next() => TempTestPaths.CreateDirectory();
}

internal static class TempFiles
{
    public static string Next(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName) || Path.GetFileName(fileName) != fileName)
        {
            throw new ArgumentException("A plain file name is required.", nameof(fileName));
        }

        return Path.Combine(TempDirectories.Next(), fileName);
    }
}
