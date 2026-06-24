using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Application.Services;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.UnitTests.Services;

public sealed class ProvinceDataAssemblerTests
{
    private static readonly IReadOnlyDictionary<SourceWorkbookKind, MetricKey[]> ExpectedMetrics =
        new Dictionary<SourceWorkbookKind, MetricKey[]>
        {
            [SourceWorkbookKind.Table1] =
            [
                MetricKey.MetroTotal,
                MetricKey.MetroIpv6,
                MetricKey.MobileCoreTotal,
                MetricKey.MobileCoreIpv6,
            ],
            [SourceWorkbookKind.Table4] =
            [
                MetricKey.HumanTotal,
                MetricKey.HumanIpv6,
                MetricKey.IotTotal,
                MetricKey.IotIpv6,
            ],
            [SourceWorkbookKind.Table5] =
            [
                MetricKey.InternetTotal,
                MetricKey.InternetIpv6,
                MetricKey.InternetOneGTotal,
                MetricKey.InternetOneGIpv6,
                MetricKey.IdcTotal,
                MetricKey.IdcIpv6,
                MetricKey.IdcTenGTotal,
                MetricKey.IdcTenGIpv6,
            ],
            [SourceWorkbookKind.Table8] =
            [
                MetricKey.BroadbandTotal,
                MetricKey.BroadbandIpv6,
            ],
        };

    [Fact]
    public void MetricKeyDefinesExactlyThe18ReportMetrics()
    {
        string[] expectedNames =
        [
            "MetroTotal", "MetroIpv6", "MobileCoreTotal", "MobileCoreIpv6",
            "InternetTotal", "InternetIpv6", "InternetOneGTotal", "InternetOneGIpv6",
            "IdcTotal", "IdcIpv6", "IdcTenGTotal", "IdcTenGIpv6",
            "HumanTotal", "HumanIpv6", "IotTotal", "IotIpv6",
            "BroadbandTotal", "BroadbandIpv6",
        ];

        Assert.Equal(expectedNames, Enum.GetNames<MetricKey>());
    }

    [Fact]
    public void SourceWorkbookKindDefinesExactlyTheFourSupportedTables()
    {
        Assert.Equal(
            ["Table1", "Table4", "Table5", "Table8"],
            Enum.GetNames<SourceWorkbookKind>());
    }

    [Fact]
    public void ExpectedMetricsByKindMatchesTheSourceTableContract()
    {
        Assert.Equal(ExpectedMetrics.Keys, ProvinceDataAssembler.ExpectedMetricsByKind.Keys);

        foreach ((SourceWorkbookKind kind, MetricKey[] metrics) in ExpectedMetrics)
        {
            Assert.Equal(metrics, ProvinceDataAssembler.ExpectedMetricsByKind[kind]);
        }
    }

    [Fact]
    public void AssembleCombinesCompleteSourcesForAllCanonicalProvinces()
    {
        var assembler = new ProvinceDataAssembler();

        ProvinceAssemblyResult result = assembler.Assemble(CreateCompleteSources());

        Assert.Empty(result.Issues);
        Assert.Equal(31, result.Reports.Count);
        Province beijing = Assert.Single(
            result.Reports.Keys,
            province => province.Name == "北京");
        Assert.Equal(1m, result.Reports[beijing][MetricKey.MetroTotal]);
        Assert.Equal(18, result.Reports[beijing].Values.Count);
    }

    [Fact]
    public void AssembleReportsMissingProvinceAndReturnsNoPartialReports()
    {
        List<SourceReadResult> sources = CreateCompleteSources();
        int table8Index = sources.FindIndex(source => source.Kind == SourceWorkbookKind.Table8);
        SourceReadResult table8 = sources[table8Index];
        Province beijing = ProvinceCatalog.All.Single(province => province.Name == "北京");
        var withoutBeijing = table8.Values.ToDictionary(entry => entry.Key, entry => entry.Value);
        withoutBeijing.Remove(beijing);
        sources[table8Index] = table8 with { Values = withoutBeijing };

        ProvinceAssemblyResult result = new ProvinceDataAssembler().Assemble(sources);

        ValidationIssue issue = Assert.Single(
            result.Issues,
            candidate => candidate.Code == "PROVINCE_MISSING");
        Assert.Contains("北京", issue.Message, StringComparison.Ordinal);
        Assert.Empty(result.Reports);
    }

    [Fact]
    public void AssembleReportsMissingMetricWithKindProvinceAndMetricAndReturnsNoPartialReports()
    {
        List<SourceReadResult> sources = CreateCompleteSources();
        int table5Index = sources.FindIndex(source => source.Kind == SourceWorkbookKind.Table5);
        SourceReadResult table5 = sources[table5Index];
        Province beijing = ProvinceCatalog.All.Single(province => province.Name == "北京");
        var table5Values = table5.Values.ToDictionary(entry => entry.Key, entry => entry.Value);
        var beijingValues = table5Values[beijing].ToDictionary(entry => entry.Key, entry => entry.Value);
        beijingValues.Remove(MetricKey.IdcTenGIpv6);
        table5Values[beijing] = beijingValues;
        sources[table5Index] = table5 with { Values = table5Values };

        ProvinceAssemblyResult result = new ProvinceDataAssembler().Assemble(sources);

        ValidationIssue issue = Assert.Single(
            result.Issues,
            candidate => candidate.Code == "METRIC_MISSING");
        Assert.Contains("Table5", issue.Message, StringComparison.Ordinal);
        Assert.Contains("北京", issue.Message, StringComparison.Ordinal);
        Assert.Contains("IdcTenGIpv6", issue.Message, StringComparison.Ordinal);
        Assert.Empty(result.Reports);
    }

    [Fact]
    public void AssembleReportsEveryMissingAndDuplicateSourceKindAndReturnsNoPartialReports()
    {
        List<SourceReadResult> sources = CreateCompleteSources();
        sources.RemoveAll(source => source.Kind == SourceWorkbookKind.Table8);
        sources.Add(sources.Single(source => source.Kind == SourceWorkbookKind.Table1));

        ProvinceAssemblyResult result = new ProvinceDataAssembler().Assemble(sources);

        ValidationIssue[] countIssues = result.Issues
            .Where(issue => issue.Code == "SOURCE_KIND_COUNT")
            .ToArray();
        Assert.Equal(2, countIssues.Length);
        Assert.Contains(countIssues, issue => issue.Message.Contains("Table1", StringComparison.Ordinal));
        Assert.Contains(countIssues, issue => issue.Message.Contains("Table8", StringComparison.Ordinal));
        Assert.Empty(result.Reports);
    }

    [Fact]
    public void AssembleAggregatesSourceIssuesAndReturnsNoPartialReports()
    {
        List<SourceReadResult> sources = CreateCompleteSources();
        ValidationIssue sourceIssue = new("SOURCE_CELL_INVALID", "源单元格无效");
        sources[0] = sources[0] with { Issues = [sourceIssue] };

        ProvinceAssemblyResult result = new ProvinceDataAssembler().Assemble(sources);

        Assert.Contains(sourceIssue, result.Issues);
        Assert.Empty(result.Reports);
    }

    [Fact]
    public void ProvinceReportInputDefensivelyCopiesValues()
    {
        Dictionary<MetricKey, decimal> values = CreateAllMetricValues();
        var input = new ProvinceReportInput(values);

        values[MetricKey.MetroTotal] = 99m;

        Assert.Equal(1m, input[MetricKey.MetroTotal]);
        IDictionary<MetricKey, decimal> dictionaryView =
            Assert.IsAssignableFrom<IDictionary<MetricKey, decimal>>(input.Values);
        Assert.Throws<NotSupportedException>(() => dictionaryView[MetricKey.MetroTotal] = 2m);
    }

    [Fact]
    public void ProvinceReportInputRejectsMissingMetricsAndListsEveryMissingMetric()
    {
        Dictionary<MetricKey, decimal> values = CreateAllMetricValues();
        values.Remove(MetricKey.MetroTotal);
        values.Remove(MetricKey.BroadbandIpv6);

        ArgumentException exception = Assert.Throws<ArgumentException>(
            () => new ProvinceReportInput(values));

        Assert.Contains("MetroTotal", exception.Message, StringComparison.Ordinal);
        Assert.Contains("BroadbandIpv6", exception.Message, StringComparison.Ordinal);
    }

    private static List<SourceReadResult> CreateCompleteSources()
    {
        return ExpectedMetrics
            .Select(entry => new SourceReadResult(
                entry.Key,
                ProvinceCatalog.All.ToDictionary(
                    province => province,
                    _ => (IReadOnlyDictionary<MetricKey, decimal>)entry.Value.ToDictionary(
                        metric => metric,
                        _ => 1m)),
                []))
            .ToList();
    }

    private static Dictionary<MetricKey, decimal> CreateAllMetricValues()
    {
        return Enum.GetValues<MetricKey>().ToDictionary(metric => metric, _ => 1m);
    }
}
