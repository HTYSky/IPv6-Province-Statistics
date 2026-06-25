using System.IO;
using Ipv6ProvinceStatistics.Application.Models;

namespace Ipv6ProvinceStatistics.App.ViewModels;

public sealed record FileCardViewModel(SourceWorkbookKind Kind, string FullPath, bool IsValid)
{
    public string FileName => Path.GetFileName(FullPath);
}
