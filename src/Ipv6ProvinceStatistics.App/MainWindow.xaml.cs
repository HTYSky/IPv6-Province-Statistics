using System.IO;
using System.Windows;
using Ipv6ProvinceStatistics.App.ViewModels;

namespace Ipv6ProvinceStatistics.App;

public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel;
    public MainWindow(MainWindowViewModel viewModel)
    {
        InitializeComponent(); DataContext = _viewModel = viewModel;
    }
    private async void OnFilesDropped(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        var paths = ((string[])e.Data.GetData(DataFormats.FileDrop))
            .Where(p => Path.GetExtension(p).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)).ToArray();
        await _viewModel.LoadFilesAsync(paths);
    }
    private async void OnClosed(object? sender, EventArgs e) => await _viewModel.CloseAsync();
}
