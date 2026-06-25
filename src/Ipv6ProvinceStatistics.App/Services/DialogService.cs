using Microsoft.Win32;

namespace Ipv6ProvinceStatistics.App.Services;

public interface IUiDialogService
{
    IReadOnlyList<string> SelectExcelFiles();
    string? SelectOutputFolder();
}

public sealed class DialogService : IUiDialogService
{
    public IReadOnlyList<string> SelectExcelFiles()
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", Multiselect = true };
        return dialog.ShowDialog() == true ? dialog.FileNames : [];
    }
    public string? SelectOutputFolder()
    {
        var dialog = new OpenFolderDialog { Title = "选择统计结果保存位置", Multiselect = false };
        return dialog.ShowDialog() == true ? dialog.FolderName : null;
    }
}
