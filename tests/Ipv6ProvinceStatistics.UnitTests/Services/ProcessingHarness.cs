using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Application.Services;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.UnitTests.Services;

public sealed class ProcessingHarness : IDisposable
{
    private readonly string _root;
    public IReadOnlyList<string> Paths { get; }
    public string OutputParent { get; }
    public StubWorkspaceManager Workspace { get; }
    public IReportProcessingService Service { get; }

    private ProcessingHarness(IReadOnlyList<string> fileNames)
    {
        _root = Path.Combine(Path.GetTempPath(), "Ipv6ProvinceStatistics.UnitTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_root);
        OutputParent = Path.Combine(_root, "output");
        Directory.CreateDirectory(OutputParent);
        Paths = fileNames.Select(name => Path.Combine(_root, name)).ToArray();
        foreach (var p in Paths) File.WriteAllText(p, "fixture");
        Workspace = new StubWorkspaceManager(Path.Combine(_root, "workspace"));
        Service = new ReportProcessingService(Workspace, new StubInspector(), new StubReader(),
            new StubExporter(), new StubOutput(), new StubLogger(), "test");
    }

    public static ProcessingHarness Create(IReadOnlyList<string> fileNames) => new(fileNames);
    public void Dispose() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }

    public sealed class StubWorkspaceManager(string wr) : ITaskWorkspaceManager
    {
        public bool CleanupCalled { get; private set; }
        public Task<TaskWorkspace> CreateAsync(IReadOnlyList<string> paths, CancellationToken t)
        {
            Directory.CreateDirectory(wr);
            var sources = paths.Select(p =>
            {
                var snap = Path.Combine(wr, Path.GetFileName(p));
                File.Copy(p, snap, true);
                return new SourceFileSnapshot(p, snap, Path.GetFileName(p), new FileInfo(p).Length,
                    File.GetLastWriteTimeUtc(p), "TEST-HASH");
            }).ToArray();
            return Task.FromResult(new TaskWorkspace(Guid.NewGuid(), wr, sources));
        }
        public Task CleanupAsync(TaskWorkspace w)
        { CleanupCalled = true; if (Directory.Exists(w.Root)) Directory.Delete(w.Root, true); return Task.CompletedTask; }
    }

    private sealed class StubInspector : IWorkbookInspector
    {
        public Task<WorkbookInspection> InspectAsync(string path, CancellationToken t)
        {
            var kind = Path.GetFileName(path)[0] switch { '1' => SourceWorkbookKind.Table1, '4' => SourceWorkbookKind.Table4,
                '5' => SourceWorkbookKind.Table5, '8' => SourceWorkbookKind.Table8, _ => throw new Exception() };
            return Task.FromResult(new WorkbookInspection(path, [kind], [new MonthMarker(2026, 5, path)], []));
        }
    }

    private sealed class StubReader : ISourceWorkbookReader
    {
        public Task<SourceReadResult> ReadAsync(string path, SourceWorkbookKind kind, CancellationToken t)
        {
            var allMetrics = Enum.GetValues<MetricKey>().ToDictionary(m => m, _ => 100m);
            var values = new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>();
            foreach (var p in ProvinceCatalog.All) values[p] = allMetrics;
            return Task.FromResult(new SourceReadResult(kind, values, []));
        }
    }

    private sealed class StubExporter : ITemplateReportExporter
    {
        public Task<IReadOnlyList<ValidationIssue>> ValidateTemplateAsync(CancellationToken t)
            => Task.FromResult<IReadOnlyList<ValidationIssue>>([]);
        public async Task<IReadOnlyList<ValidationIssue>> ExportAsync(Province p, ReportMonth m,
            ProvinceReportInput i, string o, CancellationToken t)
        { await File.WriteAllTextAsync(o, p.Name, t); return []; }
    }

    private sealed class StubOutput : IOutputTransaction
    {
        public Task<string> CreateStagingAsync(string parent, Guid id)
        { var p = Path.Combine(parent, $".stage-{id:N}"); Directory.CreateDirectory(p); return Task.FromResult(p); }
        public Task<string> PublishAsync(string staging, string parent, ReportMonth month, CancellationToken t)
        {
            if (Directory.GetFiles(staging, "*.xlsx").Length != 31) throw new Exception("not 31");
            var d = Path.Combine(parent, $"{month.FileSuffix}统计结果");
            Directory.Move(staging, d);
            return Task.FromResult(d);
        }
        public Task CleanupAsync(string path) { if (Directory.Exists(path)) Directory.Delete(path, true); return Task.CompletedTask; }
    }

    private sealed class StubLogger : IOperationLogger
    {
        public Task WriteAsync(OperationLogEntry e, CancellationToken t) => Task.CompletedTask;
    }
}
