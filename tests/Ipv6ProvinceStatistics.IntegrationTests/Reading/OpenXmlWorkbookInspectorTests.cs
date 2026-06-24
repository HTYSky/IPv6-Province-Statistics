using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Reading;

public sealed class OpenXmlWorkbookInspectorTests
{
    private readonly OpenXmlWorkbookInspector _inspector = new();

    [Theory]
    [InlineData("IDC汇总 (客户)")]
    [InlineData("IDC汇总（客户）")]
    public async Task IdentifiesTable5AliasesAndReadsMonthFromFileName(string idcName)
    {
        string path = WorkbookFixtures.CreateTable5(idcName, "5-202605.xlsx");

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Equal([SourceWorkbookKind.Table5], inspection.MatchingKinds);
        Assert.Contains(
            inspection.MonthMarkers,
            marker => marker.Year == 2026 && marker.Month == 5 && marker.Source == "5-202605.xlsx");
        Assert.Empty(inspection.Issues);
    }

    [Theory]
    [InlineData("省统计")]
    [InlineData("1-省统计")]
    public async Task IdentifiesTable8SheetAliasesAndReadsTheFirstColumnMonth(string sheetName)
    {
        string path = WorkbookFixtures.CreateTable8(sheetName, "8.xlsx");

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Equal([SourceWorkbookKind.Table8], inspection.MatchingKinds);
        Assert.Equal([new(2026, 5, "8表A列")], inspection.MonthMarkers);
        Assert.Empty(inspection.Issues);
    }

    [Fact]
    public async Task IdentifiesTable1AndReadsMonthsFromFileNameAndTitle()
    {
        string path = WorkbookFixtures.CreateTable1("1-202604.xlsx");

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Equal([SourceWorkbookKind.Table1], inspection.MatchingKinds);
        Assert.Equal(
            [new(2026, 4, "1-202604.xlsx"), new(2026, 5, "1表标题")],
            inspection.MonthMarkers);
        Assert.Empty(inspection.Issues);
    }

    [Fact]
    public async Task IdentifiesTable4WithNormalizedHeaders()
    {
        string path = WorkbookFixtures.CreateTable4("4.xlsx");

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Equal([SourceWorkbookKind.Table4], inspection.MatchingKinds);
        Assert.Empty(inspection.Issues);
    }

    [Fact]
    public async Task ReportsAmbiguousWhenMoreThanOneStructureMatches()
    {
        string path = WorkbookFixtures.CreateAmbiguous();

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Equal([SourceWorkbookKind.Table1, SourceWorkbookKind.Table4], inspection.MatchingKinds);
        var issue = Assert.Single(inspection.Issues);
        Assert.Equal("WORKBOOK_KIND_AMBIGUOUS", issue.Code);
        Assert.Equal("ambiguous.xlsx", issue.FileName);
        Assert.Contains("2", issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ReportsUnreadableWithoutLeakingTheUnderlyingError()
    {
        string path = TempFiles.Next("corrupt.xlsx");
        await File.WriteAllBytesAsync(path, [0x50, 0x4B, 0x03]);

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Empty(inspection.MatchingKinds);
        var issue = Assert.Single(inspection.Issues);
        Assert.Equal("WORKBOOK_UNREADABLE", issue.Code);
        Assert.Equal("corrupt.xlsx", issue.FileName);
        Assert.DoesNotContain("End of Central Directory", issue.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ReportsUnknownWhenTable8IsMissingARequiredHeader()
    {
        string path = TempFiles.Next("missing-j2.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "省统计",
                [
                    TestCell.SharedText("A2", "月"),
                    TestCell.SharedText("B2", "省份"),
                    TestCell.SharedText("D2", "总流量"),
                    TestCell.Number("A4", "202605"),
                ]));

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Empty(inspection.MatchingKinds);
        var issue = Assert.Single(inspection.Issues);
        Assert.Equal("WORKBOOK_STRUCTURE_UNKNOWN", issue.Code);
        Assert.Equal("missing-j2.xlsx", issue.FileName);
        Assert.Contains("必要", issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ReturnsReadOnlySnapshots()
    {
        string path = WorkbookFixtures.CreateTable8("省统计", "8.xlsx");

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Throws<NotSupportedException>(
            () => Assert.IsAssignableFrom<IList<SourceWorkbookKind>>(inspection.MatchingKinds)
                .Add(SourceWorkbookKind.Table1));
        Assert.Throws<NotSupportedException>(
            () => Assert.IsAssignableFrom<IList<MonthMarker>>(inspection.MonthMarkers)
                .Clear());
        Assert.Throws<NotSupportedException>(
            () => Assert.IsAssignableFrom<IList<ValidationIssue>>(inspection.Issues)
                .Clear());
    }

    [Fact]
    public async Task PropagatesCancellation()
    {
        string path = WorkbookFixtures.CreateTable8("省统计", "8.xlsx");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => _inspector.InspectAsync(path, cancellation.Token));
    }

    [Fact]
    public void HeaderTextRemovesUnicodeWhitespaceAndNormalizesParenthesesAndCase()
    {
        Assert.Equal("IDC汇总(客户)", HeaderText.Normalize(" I\u00a0D\u2003C汇总\u3000（客户） "));
        Assert.True(HeaderText.Contains("IP\u202fV6 日均流量（PB）", "ipv6日均流量(pb)"));
        Assert.Equal(string.Empty, HeaderText.Normalize(null));
    }
}
