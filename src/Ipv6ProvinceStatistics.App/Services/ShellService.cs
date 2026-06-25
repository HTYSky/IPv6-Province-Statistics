using System.Diagnostics;

namespace Ipv6ProvinceStatistics.App.Services;

public interface IShellService
{
    void OpenFolder(string path);
}

public sealed class ShellService : IShellService
{
    public void OpenFolder(string path) => Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
}
