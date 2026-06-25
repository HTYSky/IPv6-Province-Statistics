using Ipv6ProvinceStatistics.Application.Services;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.EndToEnd;

public sealed class MonthlyGenerationTests
{
    [Fact]
    public async Task Real_pipeline_generates_and_verifies_31_named_workbooks()
    {
        var root = TempDirectories.Next();
        var fixture = MonthlyFixtureBuilder.Create(Path.Combine(root, "inputs"), false);
        var output = Path.Combine(root, "outputs");
        var service = CreateService(root);
        var preflight = await service.PreflightAsync(fixture.Paths, null, null, CancellationToken.None);
        Assert.True(preflight.IsValid, string.Join(Environment.NewLine, preflight.Issues.Select(x => x.Message)));
        var result = await service.GenerateAsync(preflight.Batch!, output, null, CancellationToken.None);
        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Issues.Select(x => x.Message)));
        Assert.Equal(31, Directory.GetFiles(result.OutputDirectory!, "*.xlsx").Length);
        foreach (var province in ProvinceCatalog.All)
            Assert.True(File.Exists(Path.Combine(result.OutputDirectory!, $"{province.Name}-2026年05月.xlsx")));
    }

    [Fact]
    public async Task Invalid_required_value_produces_no_formal_output()
    {
        var root = TempDirectories.Next();
        var fixture = MonthlyFixtureBuilder.Create(Path.Combine(root, "inputs"), true);
        var output = Path.Combine(root, "outputs");
        var service = CreateService(root);
        var preflight = await service.PreflightAsync(fixture.Paths, null, null, CancellationToken.None);
        Assert.False(preflight.IsValid);
        Assert.Contains(preflight.Issues, issue => issue.Code == "VALUE_INVALID" && issue.Cell == "N3");
    }

    private static ReportProcessingService CreateService(string root) => new(
        new TaskWorkspaceManager(Path.Combine(root, "local", "workspaces")), new OpenXmlWorkbookInspector(),
        new OpenXmlSourceWorkbookReader(), new TemplateReportExporter(new TemplateResourceProvider()),
        new OutputTransaction(), new JsonOperationLogger(Path.Combine(root, "local", "logs")), "1.0.0-test");
}
