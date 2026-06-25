using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Extraction;

public sealed class Table1ExtractorTests
{
    [Fact]
    public async Task CarriesProvinceGroupToExactChinaUnicomRowAndReadsFourMetrics()
    {
        string path = TempFiles.Next("1.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B4", "北   京"),
                    TestCell.SharedText("C4", "中国电信"),
                    TestCell.SharedText("C5", "中国移动"),
                    TestCell.SharedText("C6", "中国联通"),
                    TestCell.Number("D6", "101.25"),
                    TestCell.SharedText("E6", "61.5"),
                    TestCell.Number("G6", "44.75"),
                    TestCell.Number("H6", "22.125"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        SourceReadResult result = await reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            CancellationToken.None);

        Assert.Equal(SourceWorkbookKind.Table1, result.Kind);
        Assert.Empty(result.Issues);
        KeyValuePair<Province, IReadOnlyDictionary<MetricKey, decimal>> province =
            Assert.Single(result.Values);
        Assert.Equal("北京", province.Key.Name);
        Assert.Equal(101.25m, province.Value[MetricKey.MetroTotal]);
        Assert.Equal(61.5m, province.Value[MetricKey.MetroIpv6]);
        Assert.Equal(44.75m, province.Value[MetricKey.MobileCoreTotal]);
        Assert.Equal(22.125m, province.Value[MetricKey.MobileCoreIpv6]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("NULL")]
    [InlineData("not-a-number")]
    public async Task InvalidSharedTextValueReportsItsCellContextAndReturnsNoValues(
        string invalidValue)
    {
        string path = TempFiles.Next("1.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B4", "天津"),
                    TestCell.SharedText("C4", "中国联通"),
                    TestCell.Number("D4", "1"),
                    TestCell.Number("E4", "2"),
                    TestCell.Number("G4", "3"),
                    TestCell.Number("H4", "4"),
                    TestCell.SharedText("B5", "北京市"),
                    TestCell.SharedText("C5", "中国电信"),
                    TestCell.SharedText("C6", "中国联通"),
                    TestCell.Number("D6", "101.25"),
                    TestCell.SharedText("E6", invalidValue),
                    TestCell.Number("G6", "44.75"),
                    TestCell.Number("H6", "22.125"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        SourceReadResult result = await reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            CancellationToken.None);

        Assert.Empty(result.Values);
        ValidationIssue issue = Assert.Single(result.Issues);
        Assert.Equal("VALUE_INVALID", issue.Code);
        Assert.Equal("1.xlsx", issue.FileName);
        Assert.Equal("分省统计表", issue.Sheet);
        Assert.Equal("北京", issue.Province);
        Assert.Equal("E6", issue.Cell);
        Assert.Contains("有效数字", issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ChinaUnicomRowWithoutProvinceGroupReportsItsCarrierCell()
    {
        string path = TempFiles.Next("无分组.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("C4", "中国联通"),
                    TestCell.Number("D4", "1"),
                    TestCell.Number("E4", "2"),
                    TestCell.Number("G4", "3"),
                    TestCell.Number("H4", "4"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        SourceReadResult result = await reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            CancellationToken.None);

        Assert.Empty(result.Values);
        ValidationIssue issue = Assert.Single(result.Issues);
        Assert.Equal("PROVINCE_GROUP_MISSING", issue.Code);
        Assert.Equal("无分组.xlsx", issue.FileName);
        Assert.Equal("分省统计表", issue.Sheet);
        Assert.Null(issue.Province);
        Assert.Equal("C4", issue.Cell);
    }

    [Fact]
    public async Task UnknownNonEmptyProvinceBreaksGroupInheritanceWithoutLeakingItsText()
    {
        const string unknownProvince = "未知区域";
        string path = TempFiles.Next("省份边界.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B4", "北京"),
                    TestCell.SharedText("C4", "中国联通"),
                    TestCell.Number("D4", "1"),
                    TestCell.Number("E4", "2"),
                    TestCell.Number("G4", "3"),
                    TestCell.Number("H4", "4"),
                    TestCell.SharedText("B5", unknownProvince),
                    TestCell.SharedText("C5", "中国电信"),
                    TestCell.SharedText("C6", "中国联通"),
                    TestCell.Number("D6", "5"),
                    TestCell.Number("E6", "6"),
                    TestCell.Number("G6", "7"),
                    TestCell.Number("H6", "8"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        SourceReadResult result = await reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            CancellationToken.None);

        Assert.Empty(result.Values);
        Assert.Collection(
            result.Issues,
            issue =>
            {
                Assert.Equal("PROVINCE_INVALID", issue.Code);
                Assert.Equal("省份边界.xlsx", issue.FileName);
                Assert.Equal("分省统计表", issue.Sheet);
                Assert.Null(issue.Province);
                Assert.Equal("B5", issue.Cell);
                Assert.Contains("B5", issue.Message, StringComparison.Ordinal);
                Assert.DoesNotContain(unknownProvince, issue.Message, StringComparison.Ordinal);
            },
            issue =>
            {
                Assert.Equal("PROVINCE_GROUP_MISSING", issue.Code);
                Assert.Equal("C6", issue.Cell);
            });
    }

    [Fact]
    public async Task DuplicateChinaUnicomRowsForAProvinceReportDuplicateAndReturnNoValues()
    {
        string path = TempFiles.Next("重复.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B4", "北京"),
                    TestCell.SharedText("C4", "中国联通"),
                    TestCell.Number("D4", "1"),
                    TestCell.Number("E4", "2"),
                    TestCell.Number("G4", "3"),
                    TestCell.Number("H4", "4"),
                    TestCell.SharedText("C5", "中国联通"),
                    TestCell.Number("D5", "5"),
                    TestCell.Number("E5", "6"),
                    TestCell.Number("G5", "7"),
                    TestCell.Number("H5", "8"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        SourceReadResult result = await reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            CancellationToken.None);

        Assert.Empty(result.Values);
        ValidationIssue issue = Assert.Single(result.Issues);
        Assert.Equal("PROVINCE_DUPLICATE", issue.Code);
        Assert.Equal("重复.xlsx", issue.FileName);
        Assert.Equal("分省统计表", issue.Sheet);
        Assert.Equal("北京", issue.Province);
    }

    [Fact]
    public async Task ReadsScientificNotationAndThousandsSeparatedSharedText()
    {
        string path = TempFiles.Next("数字格式.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B4", "北京"),
                    TestCell.SharedText("C4", "中国联通"),
                    TestCell.SharedText("D4", "1.0125E2"),
                    TestCell.SharedText("E4", "1,061.5"),
                    TestCell.Number("G4", "4.475E1"),
                    TestCell.SharedText("H4", "2,212.5"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        SourceReadResult result = await reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            CancellationToken.None);

        Assert.Empty(result.Issues);
        IReadOnlyDictionary<MetricKey, decimal> metrics = Assert.Single(result.Values).Value;
        Assert.Equal(101.25m, metrics[MetricKey.MetroTotal]);
        Assert.Equal(1061.5m, metrics[MetricKey.MetroIpv6]);
        Assert.Equal(44.75m, metrics[MetricKey.MobileCoreTotal]);
        Assert.Equal(2212.5m, metrics[MetricKey.MobileCoreIpv6]);
    }

    [Theory]
    [InlineData("12,34.5")]
    [InlineData("1,2,3")]
    [InlineData("1E-29")]
    [InlineData("1E-30")]
    [InlineData("0.12345678901234567890123456789")]
    public async Task RejectsMalformedOrLossyDecimalText(string invalidValue)
    {
        string path = TempFiles.Next("无损数字.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B4", "北京"),
                    TestCell.SharedText("C4", "中国联通"),
                    TestCell.SharedText("D4", invalidValue),
                    TestCell.Number("E4", "2"),
                    TestCell.Number("G4", "3"),
                    TestCell.Number("H4", "4"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        SourceReadResult result = await reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            CancellationToken.None);

        Assert.Empty(result.Values);
        ValidationIssue issue = Assert.Single(result.Issues);
        Assert.Equal("VALUE_INVALID", issue.Code);
        Assert.Equal("D4", issue.Cell);
    }

    [Fact]
    public async Task IgnoresCarrierTextThatOnlyContainsChinaUnicom()
    {
        string path = TempFiles.Next("联通合计.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B4", "北京"),
                    TestCell.SharedText("C4", "中国联通合计"),
                    TestCell.Number("D4", "999"),
                    TestCell.Number("E4", "999"),
                    TestCell.Number("G4", "999"),
                    TestCell.Number("H4", "999"),
                    TestCell.SharedText("C5", " 中国联通 "),
                    TestCell.Number("D5", "1"),
                    TestCell.Number("E5", "2"),
                    TestCell.Number("G5", "3"),
                    TestCell.Number("H5", "4"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        SourceReadResult result = await reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            CancellationToken.None);

        Assert.Empty(result.Issues);
        IReadOnlyDictionary<MetricKey, decimal> metrics = Assert.Single(result.Values).Value;
        Assert.Equal(1m, metrics[MetricKey.MetroTotal]);
        Assert.Equal(2m, metrics[MetricKey.MetroIpv6]);
        Assert.Equal(3m, metrics[MetricKey.MobileCoreTotal]);
        Assert.Equal(4m, metrics[MetricKey.MobileCoreIpv6]);
    }

    [Fact]
    public async Task PropagatesCancellation()
    {
        string path = TempFiles.Next("取消.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B4", "北京"),
                    TestCell.SharedText("C4", "中国联通"),
                    TestCell.Number("D4", "1"),
                    TestCell.Number("E4", "2"),
                    TestCell.Number("G4", "3"),
                    TestCell.Number("H4", "4"),
                ]));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Task<SourceReadResult> readTask = reader.ReadAsync(
            path,
            SourceWorkbookKind.Table1,
            cancellation.Token);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => readTask);
        Assert.True(readTask.IsCanceled);
    }

    [Fact]
    public async Task RejectsKindsWithoutAnImplementedExtractor()
    {
        var kind = (SourceWorkbookKind)99;
        string path = TempFiles.Next("尚未支持.xlsx");
        TestWorkbookBuilder.Create(path, new TestSheet("数据", []));
        ISourceWorkbookReader reader = new OpenXmlSourceWorkbookReader();

        NotSupportedException exception = await Assert.ThrowsAsync<NotSupportedException>(
            () => reader.ReadAsync(path, kind, CancellationToken.None));

        Assert.Equal($"Extractor not implemented: {kind}", exception.Message);
    }
}
