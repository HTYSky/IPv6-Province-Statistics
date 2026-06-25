using System.Collections.ObjectModel;
using Ipv6ProvinceStatistics.App.Services;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.App.ViewModels;

public sealed class MainWindowViewModel : ObservableObject
{
    private readonly IReportProcessingService _processing;
    private readonly IUiDialogService _dialogs;
    private readonly IShellService _shell;
    private PreparedBatch? _batch;
    private IReadOnlyList<string> _selectedPaths = [];
    private CancellationTokenSource? _cancellation;
    private string? _lastOutput;

    private string _statusText = "请选择本月四张源表";
    private string _resultDetail = string.Empty;
    private string _outputParent = string.Empty;
    private int _selectedYear = DateTime.Now.Year;
    private int _selectedMonth = DateTime.Now.Month;
    private bool _isBusy;

    public MainWindowViewModel(IReportProcessingService processing, IUiDialogService dialogs, IShellService shell)
    {
        _processing = processing; _dialogs = dialogs; _shell = shell;
        SelectFilesCommand = new(async () => await LoadFilesAsync(_dialogs.SelectExcelFiles()), () => !IsBusy);
        RevalidateMonthCommand = new(async () => await RevalidateAsync(), () => _selectedPaths.Count == 4 && !IsBusy);
        SelectOutputCommand = new(() => { var f = _dialogs.SelectOutputFolder(); if (f is not null) OutputParent = f; return Task.CompletedTask; }, () => !IsBusy);
        GenerateCommand = new(async () => await GenerateAsync(), () => _batch is not null && !IsBusy && !string.IsNullOrWhiteSpace(OutputParent));
        CancelCommand = new(() => { _cancellation?.Cancel(); return Task.CompletedTask; }, () => IsBusy);
        OpenOutputCommand = new(() => { if (_lastOutput is not null) _shell.OpenFolder(_lastOutput); return Task.CompletedTask; }, () => _lastOutput is not null);
    }

    public ObservableCollection<FileCardViewModel> Files { get; } = [];
    public ObservableCollection<string> Issues { get; } = [];

    public string StatusText { get => _statusText; private set => Set(ref _statusText, value); }
    public string ResultDetail { get => _resultDetail; private set => Set(ref _resultDetail, value); }
    public string OutputParent { get => _outputParent; set => Set(ref _outputParent, value); }
    public int SelectedYear { get => _selectedYear; set => Set(ref _selectedYear, value); }
    public int SelectedMonth { get => _selectedMonth; set => Set(ref _selectedMonth, value); }
    public bool IsBusy { get => _isBusy; private set { if (Set(ref _isBusy, value)) RefreshCommands(); } }

    public AsyncRelayCommand SelectFilesCommand { get; }
    public AsyncRelayCommand RevalidateMonthCommand { get; }
    public AsyncRelayCommand SelectOutputCommand { get; }
    public AsyncRelayCommand GenerateCommand { get; }
    public AsyncRelayCommand CancelCommand { get; }
    public AsyncRelayCommand OpenOutputCommand { get; }

    public async Task LoadFilesAsync(IReadOnlyList<string> paths)
    {
        _selectedPaths = paths.ToArray();
        await PreflightAsync(_selectedPaths, null);
    }

    private async Task RevalidateAsync()
        => await PreflightAsync(_selectedPaths, new ReportMonth(SelectedYear, SelectedMonth));

    private async Task PreflightAsync(IReadOnlyList<string> paths, ReportMonth? selected)
    {
        Files.Clear(); Issues.Clear(); _lastOutput = null; ResultDetail = string.Empty; IsBusy = true;
        StatusText = "正在识别和校验...";
        _cancellation = new CancellationTokenSource();
        try
        {
            var progress = new Progress<ProcessingProgress>(p => StatusText = p.Message);
            var result = await Task.Run(() => _processing.PreflightAsync(paths, selected, progress, _cancellation.Token), _cancellation.Token);
            if (!result.IsValid)
            {
                if (result.Issues.Count > 0) Issues.Add("预检失败，已丢弃临时数据。以下是发现的问题：");
                foreach (var issue in result.Issues) Issues.Add(FormatIssue(issue));
                StatusText = "请修正以上问题后重试";
                _batch = null;
            }
            else
            {
                _batch = result.Batch;
                foreach (var source in _batch!.IdentifiedSources)
                    Files.Add(new(source.Key, source.Value.OriginalPath, true));
                StatusText = $"{_batch.Reports.Count}/31 个省份数据完整";
                SelectedYear = _batch.Month.Year; SelectedMonth = _batch.Month.Month;
            }
        }
        catch (OperationCanceledException) { StatusText = "已取消"; }
        catch (Exception ex) { Issues.Add($"内部错误：{ex.Message}"); StatusText = "请修正以上问题后重试"; _batch = null; }
        finally { IsBusy = false; _cancellation?.Dispose(); _cancellation = null; }
    }

    private async Task GenerateAsync()
    {
        if (_batch is null) return;
        IsBusy = true; StatusText = "正在生成...";
        _cancellation = new CancellationTokenSource();
        try
        {
            var progress = new Progress<ProcessingProgress>(p => StatusText = p.Message);
            var result = await Task.Run(() => _processing.GenerateAsync(_batch, OutputParent, progress, _cancellation.Token), _cancellation.Token);
            if (result.Succeeded)
            {
                _lastOutput = result.OutputDirectory;
                ResultDetail = $"输出：{result.OutputDirectory} · 耗时 {result.Duration.TotalSeconds:F1} 秒";
                StatusText = "已生成 31 个统计表"; _batch = null;
            }
            else
            {
                Issues.Clear();
                foreach (var issue in result.Issues) Issues.Add(FormatIssue(issue));
                StatusText = "生成失败，未留下正式结果"; _batch = null;
            }
        }
        catch (OperationCanceledException) { StatusText = "已取消"; }
        catch (Exception ex) { Issues.Clear(); Issues.Add($"内部错误：{ex.Message}"); StatusText = "生成失败，未留下正式结果"; _batch = null; }
        finally { IsBusy = false; _cancellation?.Dispose(); _cancellation = null; }
    }

    public async Task CloseAsync()
    {
        _cancellation?.Cancel();
        if (_batch is not null) { await _processing.DiscardAsync(_batch); _batch = null; }
    }

    private void RefreshCommands()
    {
        SelectFilesCommand.RaiseCanExecuteChanged();
        RevalidateMonthCommand.RaiseCanExecuteChanged();
        SelectOutputCommand.RaiseCanExecuteChanged();
        GenerateCommand.RaiseCanExecuteChanged();
        CancelCommand.RaiseCanExecuteChanged();
        OpenOutputCommand.RaiseCanExecuteChanged();
    }

    private static string FormatIssue(ValidationIssue issue)
    {
        var location = new[] { issue.FileName, issue.Sheet, issue.Province, issue.Cell }
            .Where(value => !string.IsNullOrWhiteSpace(value));
        var prefix = string.Join(" → ", location);
        return string.IsNullOrEmpty(prefix) ? issue.Message : $"{prefix}：{issue.Message}";
    }
}
