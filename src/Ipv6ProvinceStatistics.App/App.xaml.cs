using System.Windows;
using Ipv6ProvinceStatistics.App.Services;
using Ipv6ProvinceStatistics.App.ViewModels;
using Ipv6ProvinceStatistics.Application.Services;
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

namespace Ipv6ProvinceStatistics.App;

public partial class App : System.Windows.Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, args) =>
        {
            MessageBox.Show("程序遇到未预期错误。未生成任何正式结果，请查看本地日志。", "IPv6 省级统计助手",
                MessageBoxButton.OK, MessageBoxImage.Error); args.Handled = true;
        };
        var processing = new ReportProcessingService(
            new TaskWorkspaceManager(), new OpenXmlWorkbookInspector(),
            new OpenXmlSourceWorkbookReader(),
            new TemplateReportExporter(new TemplateResourceProvider()),
            new OutputTransaction(), new JsonOperationLogger(),
            typeof(App).Assembly.GetName().Version?.ToString() ?? "1.0.0");
        var viewModel = new MainWindowViewModel(processing, new DialogService(), new ShellService());
        new MainWindow(viewModel).Show();
    }
}
