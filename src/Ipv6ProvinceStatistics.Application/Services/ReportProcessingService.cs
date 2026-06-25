using System.Diagnostics;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Application.Services;
public sealed class ReportProcessingService(
    ITaskWorkspaceManager workspaces, IWorkbookInspector inspector, ISourceWorkbookReader reader,
    ITemplateReportExporter exporter, IOutputTransaction output, IOperationLogger logger,
    string applicationVersion) : IReportProcessingService
{
    public async Task<PreflightResult> PreflightAsync(IReadOnlyList<string> paths, ReportMonth? selected,
        IProgress<ProcessingProgress>? progress, CancellationToken token)
    {
        if (paths.Count != 4 || paths.Distinct(StringComparer.OrdinalIgnoreCase).Count() != 4)
            return new(null, [new("SOURCE_COUNT", "必须选择四个不同的 XLSX 文件。")]);
        TaskWorkspace? workspace = null;
        var inspections = new List<WorkbookInspection>();
        MonthResolution? resolvedMonth = null;
        try
        {
            progress?.Report(new(ProcessingStage.Snapshotting, 0, 4, "正在创建源文件快照"));
            workspace = await workspaces.CreateAsync(paths, token);
            foreach (var source in workspace.Sources)
            {
                progress?.Report(new(ProcessingStage.Identifying, inspections.Count, 4, source.FileName));
                var inspection = await inspector.InspectAsync(source.SnapshotPath, token);
                inspections.Add(inspection with { Issues = inspection.Issues
                    .Select(i => i with { FileName = source.FileName }).ToArray() });
            }
            var issues = inspections.SelectMany(x => x.Issues).ToList();
            foreach (var kind in Enum.GetValues<SourceWorkbookKind>())
                if (inspections.Count(x => x.MatchingKinds.Count == 1 && x.MatchingKinds[0] == kind) != 1)
                    issues.Add(new("SOURCE_KIND_COUNT", $"{kind} 必须且只能识别出一个文件。"));
            resolvedMonth = MonthResolver.Resolve(inspections.SelectMany(x => x.MonthMarkers).ToArray(), selected);
            issues.AddRange(resolvedMonth.Issues);
            if (issues.Count == 0)
            {
                var reads = new List<SourceReadResult>();
                foreach (var inspection in inspections)
                {
                    var read = await reader.ReadAsync(inspection.Path, inspection.MatchingKinds.Single(), token);
                    var original = workspace.Sources.Single(s => s.SnapshotPath == inspection.Path);
                    reads.Add(read with { Issues = read.Issues
                        .Select(i => i with { FileName = original.FileName }).ToArray() });
                }
                var assembly = new ProvinceDataAssembler().Assemble(reads);
                issues.AddRange(assembly.Issues);
                issues.AddRange(await exporter.ValidateTemplateAsync(token));
                foreach (var report in assembly.Reports.Values)
                    issues.AddRange(ReportCalculator.Calculate(report).Issues);
                if (assembly.Reports.Count > 0)
                {
                    var identified = inspections.Where(i => i.MatchingKinds.Count == 1)
                        .Select(i => new { Kind = i.MatchingKinds[0], Source = workspace.Sources.Single(s => s.SnapshotPath == i.Path) })
                        .GroupBy(x => x.Kind).Where(g => g.Count() == 1)
                        .ToDictionary(g => g.Key, g => g.Single().Source);
                    return new(new(workspace, resolvedMonth!.Month!.Value, assembly.Reports, identified), []);
                }
            }
            var recognized = BuildIdentified(inspections, workspace);
            await workspaces.CleanupAsync(workspace);
            await LogAsync(workspace, "PreflightFailed", resolvedMonth?.Month, recognized, null, 0, TimeSpan.Zero, issues, token);
            return new(null, issues);
        }
        catch (OperationCanceledException)
        {
            if (workspace is not null) await workspaces.CleanupAsync(workspace);
            throw;
        }
        catch (Exception exception)
        {
            var issue = new ValidationIssue("INTERNAL_ERROR", exception.Message);
            if (workspace is not null)
            {
                await LogAsync(workspace, "PreflightFailed", resolvedMonth?.Month,
                    BuildIdentified(inspections, workspace), null, 0, TimeSpan.Zero, [issue], CancellationToken.None);
                await workspaces.CleanupAsync(workspace);
            }
            return new(null, [issue]);
        }
    }

    public async Task<GenerationResult> GenerateAsync(PreparedBatch batch, string parent,
        IProgress<ProcessingProgress>? progress, CancellationToken token)
    {
        var watch = Stopwatch.StartNew();
        string? staging = null;
        var issues = new List<ValidationIssue>();
        try
        {
            staging = await output.CreateStagingAsync(parent, batch.Workspace.TaskId);
            var provinces = batch.Reports.Keys.ToArray();
            for (var index = 0; index < provinces.Length; index++)
            {
                token.ThrowIfCancellationRequested();
                var province = provinces[index];
                progress?.Report(new(ProcessingStage.Generating, index, provinces.Length, $"正在生成 {province.Name}"));
                var path = Path.Combine(staging, $"{province.Name}-{batch.Month.FileSuffix}.xlsx");
                issues.AddRange(await exporter.ExportAsync(province, batch.Month, batch.Reports[province], path, token));
                if (issues.Count > 0) break;
            }
            if (issues.Count > 0)
            {
                await output.CleanupAsync(staging);
                await workspaces.CleanupAsync(batch.Workspace);
                await LogAsync(batch.Workspace, "Failed", batch.Month, batch.IdentifiedSources, null, 0, watch.Elapsed, issues, token);
                return new(false, null, 0, watch.Elapsed, issues);
            }
            progress?.Report(new(ProcessingStage.Publishing, provinces.Length, provinces.Length, "正在发布完整结果"));
            var published = await output.PublishAsync(staging, parent, batch.Month, token);
            staging = null;
            await workspaces.CleanupAsync(batch.Workspace);
            await LogAsync(batch.Workspace, "Succeeded", batch.Month, batch.IdentifiedSources, published, 31, watch.Elapsed, [], token);
            return new(true, published, 31, watch.Elapsed, []);
        }
        catch (OperationCanceledException)
        {
            if (staging is not null) await output.CleanupAsync(staging);
            await workspaces.CleanupAsync(batch.Workspace);
            await LogAsync(batch.Workspace, "Canceled", batch.Month, batch.IdentifiedSources, null, 0, watch.Elapsed, [], CancellationToken.None);
            throw;
        }
        catch (Exception exception)
        {
            if (staging is not null) await output.CleanupAsync(staging);
            await workspaces.CleanupAsync(batch.Workspace);
            issues.Add(new("INTERNAL_ERROR", exception.Message));
            await LogAsync(batch.Workspace, "Failed", batch.Month, batch.IdentifiedSources, null, 0, watch.Elapsed, issues, CancellationToken.None);
            return new(false, null, 0, watch.Elapsed, issues);
        }
    }

    public Task DiscardAsync(PreparedBatch batch) => workspaces.CleanupAsync(batch.Workspace);

    private static IReadOnlyDictionary<SourceWorkbookKind, SourceFileSnapshot> BuildIdentified(
        IReadOnlyList<WorkbookInspection> inspections, TaskWorkspace workspace) => inspections
        .Where(i => i.MatchingKinds.Count == 1)
        .Select(i => new { Kind = i.MatchingKinds[0], Source = workspace.Sources.Single(s => s.SnapshotPath == i.Path) })
        .GroupBy(x => x.Kind).Where(g => g.Count() == 1).ToDictionary(g => g.Key, g => g.Single().Source);

    private Task LogAsync(TaskWorkspace workspace, string status, ReportMonth? month,
        IReadOnlyDictionary<SourceWorkbookKind, SourceFileSnapshot> identified,
        string? directory, int count, TimeSpan duration, IReadOnlyList<ValidationIssue> issues,
        CancellationToken token)
    {
        var kindsByPath = identified.ToDictionary(p => p.Value.SnapshotPath, p => p.Key.ToString(), StringComparer.OrdinalIgnoreCase);
        var sources = workspace.Sources.Select(source => new LogSource(source.FileName, source.Length,
            source.LastWriteTimeUtc, source.Sha256, kindsByPath.TryGetValue(source.SnapshotPath, out var kind) ? kind : null)).ToArray();
        return logger.WriteAsync(new(workspace.TaskId, DateTimeOffset.UtcNow, applicationVersion, status,
            month?.Year, month?.Month, sources, directory, count, (long)duration.TotalMilliseconds,
            issues.Select(i => $"{i.Code}: {i.Message}").ToArray()), token);
    }
}
