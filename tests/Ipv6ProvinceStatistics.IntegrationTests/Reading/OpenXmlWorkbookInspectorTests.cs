using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
    [InlineData("省统计", "8表 省统计!A4")]
    [InlineData("1-省统计", "8表 1-省统计!A4")]
    public async Task IdentifiesTable8SheetAliasesAndReadsTheFirstColumnMonth(
        string sheetName,
        string expectedSource)
    {
        string path = WorkbookFixtures.CreateTable8(sheetName, "8.xlsx");

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Equal([SourceWorkbookKind.Table8], inspection.MatchingKinds);
        Assert.Equal([new(2026, 5, expectedSource)], inspection.MonthMarkers);
        Assert.Empty(inspection.Issues);
    }

    [Fact]
    public async Task SkipsNonMonthValuesBeforeTheFirstTable8Month()
    {
        string path = TempFiles.Next("8-scan.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "1-省统计",
                [
                    TestCell.SharedText("A2", "月"),
                    TestCell.SharedText("B2", "省份"),
                    TestCell.SharedText("D2", "总流量"),
                    TestCell.SharedText("J2", "IPv6总流量"),
                    TestCell.SharedText("A4", "说明"),
                    TestCell.Number("A5", "202605"),
                ]));

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Equal([SourceWorkbookKind.Table8], inspection.MatchingKinds);
        Assert.Equal([new(2026, 5, "8表 1-省统计!A5")], inspection.MonthMarkers);
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

    [Theory]
    [InlineData("missing-shared-strings")]
    [InlineData("missing-worksheet-relationship")]
    public async Task ReportsUnreadableForKnownMalformedWorkbookStructures(string damage)
    {
        string fileName = $"{damage}.xlsx";
        string path = WorkbookFixtures.CreateTable1(fileName);
        DamageWorkbook(path, damage);

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Empty(inspection.MatchingKinds);
        ValidationIssue issue = Assert.Single(inspection.Issues);
        Assert.Equal("WORKBOOK_UNREADABLE", issue.Code);
        Assert.Equal(fileName, issue.FileName);
    }

    [Fact]
    public async Task DoesNotConvertArgumentErrorsToUnreadableIssues()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => _inspector.InspectAsync(null!, default));
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

    [Theory]
    [InlineData("IPv6总流量")]
    [InlineData("IPv4总流量")]
    [InlineData("总流量（仅IPv6）")]
    public async Task DoesNotUseASingleVersionHeaderForTheTable8OrdinaryTotal(
        string ordinaryTotalHeader)
    {
        string path = TempFiles.Next("table8-ipv6-only-total.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "省统计",
                [
                    TestCell.SharedText("A2", "月"),
                    TestCell.SharedText("B2", "省份"),
                    TestCell.SharedText("D2", ordinaryTotalHeader),
                    TestCell.SharedText("J2", "IPv6总流量"),
                ]));

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Empty(inspection.MatchingKinds);
        Assert.Equal("WORKBOOK_STRUCTURE_UNKNOWN", Assert.Single(inspection.Issues).Code);
    }

    [Fact]
    public async Task AcceptsTable8OrdinaryTotalWithAV4AndV6Description()
    {
        string path = TempFiles.Next("table8-combined-total.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "1-省统计",
                [
                    TestCell.SharedText("A2", "月"),
                    TestCell.SharedText("B2", "省份"),
                    TestCell.SharedText("D2", "总流量\n(v4+v6，GB)"),
                    TestCell.SharedText("J2", "IPv6总流量(GB)"),
                ]));

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Equal([SourceWorkbookKind.Table8], inspection.MatchingKinds);
        Assert.Empty(inspection.Issues);
    }

    [Theory]
    [InlineData("互联网专线汇总", "E2", "IPv6总流量")]
    [InlineData("互联网专线汇总", "K2", "IPv6总流量")]
    [InlineData("IDC汇总(客户)", "F2", "IPv6总流量")]
    [InlineData("IDC汇总(客户)", "M2", "IPv6总流量")]
    [InlineData("互联网专线汇总", "E2", "IPv4总流量")]
    [InlineData("IDC汇总(客户)", "F2", "总流量(仅IPv6)")]
    public async Task DoesNotUseASingleVersionHeaderForATable5OrdinaryTotal(
        string sheetName,
        string address,
        string replacementHeader)
    {
        string path = TempFiles.Next("table5-ipv6-only-total.xlsx");
        TestCell[] internetHeaders =
        [
            TestCell.SharedText("E2", "总流量"),
            TestCell.SharedText("F2", "IPv6流量"),
            TestCell.SharedText("K2", "总流量"),
            TestCell.SharedText("L2", "IPv6流量"),
            TestCell.SharedText("M2", "省"),
        ];
        TestCell[] idcHeaders =
        [
            TestCell.SharedText("F2", "总流量"),
            TestCell.SharedText("G2", "V6日流量"),
            TestCell.SharedText("M2", "总流量"),
            TestCell.SharedText("N2", "V6日流量"),
            TestCell.SharedText("O2", "省"),
        ];
        TestCell[] targetHeaders = sheetName == "互联网专线汇总"
            ? internetHeaders
            : idcHeaders;
        int targetIndex = Array.FindIndex(
            targetHeaders,
            cell => string.Equals(cell.Address, address, StringComparison.Ordinal));
        targetHeaders[targetIndex] = TestCell.SharedText(address, replacementHeader);
        TestWorkbookBuilder.Create(
            path,
            new TestSheet("互联网专线汇总", internetHeaders),
            new TestSheet("IDC汇总(客户)", idcHeaders));

        WorkbookInspection inspection = await _inspector.InspectAsync(path, default);

        Assert.Empty(inspection.MatchingKinds);
        Assert.Equal("WORKBOOK_STRUCTURE_UNKNOWN", Assert.Single(inspection.Issues).Code);
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
    public async Task ReturnsBeforeALargeWorkbookInspectionCompletes()
    {
        const int cellCount = 100_000;
        string path = TempFiles.Next("large-inspection.xlsx");
        TestCell[] cells = Enumerable.Range(1, cellCount)
            .Select(row => TestCell.Number($"A{row}", "1"))
            .ToArray();
        TestWorkbookBuilder.Create(path, new TestSheet("数据", cells));

        Stopwatch stopwatch = Stopwatch.StartNew();
        Task<WorkbookInspection> inspectionTask = _inspector.InspectAsync(path, default);
        stopwatch.Stop();

        Assert.False(inspectionTask.IsCompleted);
        Assert.True(
            stopwatch.Elapsed < TimeSpan.FromSeconds(2),
            $"InspectAsync blocked its caller for {stopwatch.Elapsed}.");
        WorkbookInspection inspection = await inspectionTask;
        Assert.Equal("WORKBOOK_STRUCTURE_UNKNOWN", Assert.Single(inspection.Issues).Code);
    }

    [Fact]
    public void HeaderTextRemovesUnicodeWhitespaceAndNormalizesParenthesesAndCase()
    {
        Assert.Equal("IDC汇总(客户)", HeaderText.Normalize(" I\u00a0D\u2003C汇总\u3000（客户） "));
        Assert.Equal("IDC汇总(客户)", HeaderText.Normalize("ＩＤＣ汇总（客户）"));
        Assert.True(HeaderText.Contains("IP\u202fV6 日均流量（PB）", "ipv6日均流量(pb)"));
        Assert.True(HeaderText.Contains("ＩＰＶ６总流量", "IPv6总流量"));
        Assert.Equal("A\u200bB", HeaderText.Normalize("A\u200bB"));
        Assert.Equal(string.Empty, HeaderText.Normalize(null));
    }

    private static void DamageWorkbook(string path, string damage)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, isEditable: true);
        WorkbookPart workbookPart = Assert.IsType<WorkbookPart>(document.WorkbookPart);
        if (damage == "missing-shared-strings")
        {
            SharedStringTablePart sharedStringsPart =
                Assert.IsType<SharedStringTablePart>(workbookPart.SharedStringTablePart);
            workbookPart.DeletePart(sharedStringsPart);
            return;
        }

        Workbook workbook = Assert.IsType<Workbook>(workbookPart.Workbook);
        Sheets sheets = Assert.IsType<Sheets>(workbook.GetFirstChild<Sheets>());
        Sheet sheet = Assert.Single(sheets.Elements<Sheet>());
        WorksheetPart worksheetPart = Assert.IsType<WorksheetPart>(
            workbookPart.GetPartById(Assert.IsType<string>(sheet.Id?.Value)));
        workbookPart.DeletePart(worksheetPart);
    }
}
