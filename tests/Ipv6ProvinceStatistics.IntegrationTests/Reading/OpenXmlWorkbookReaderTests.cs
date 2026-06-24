using System.Diagnostics;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;
using Xunit.Abstractions;

namespace Ipv6ProvinceStatistics.IntegrationTests.Reading;

public sealed class OpenXmlWorkbookReaderTests
{
    private readonly ITestOutputHelper _output;

    public OpenXmlWorkbookReaderTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void ReadsSharedInlineNumericAndFormulaCachedCells()
    {
        string path = TempFiles.Next("示例 workbook.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "数据",
                [
                    TestCell.SharedText("A1", "北京市"),
                    TestCell.InlineText("B1", "中国联通"),
                    TestCell.Number("C1", "6.5E3"),
                    TestCell.Formula("D1", "C1/2", "3250"),
                ]));

        using OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false);

        Assert.Equal(["数据"], reader.SheetNames);
        Assert.True(reader.HasSheet("数据"));
        Assert.False(reader.HasSheet("数据副本"));
        Assert.NotNull(reader.FindCell("数据", "a1"));
        Assert.Equal("北京市", reader.GetText("数据", "A1"));
        Assert.Equal("中国联通", reader.GetText("数据", "B1"));
        Assert.Equal("6.5E3", reader.GetText("数据", "C1"));
        Assert.True(reader.TryGetDecimal("数据", "C1", out decimal numericValue));
        Assert.Equal(6500m, numericValue);
        Assert.Equal("C1/2", reader.GetFormula("数据", "D1"));
        Assert.Equal("3250", reader.GetText("数据", "D1"));
        Assert.Equal([1U], reader.GetPopulatedRows("数据"));
    }

    [Fact]
    public void ReadsBooleanValuesAndReturnsNullForAnEmptyAddress()
    {
        string path = TempFiles.Next("boolean.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "数据",
                [TestCell.Number("A1", "1"), TestCell.Number("A2", "0")]));
        MarkCellsAsBoolean(path);

        using OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false);

        Assert.Equal("TRUE", reader.GetText("数据", "A1"));
        Assert.Equal("FALSE", reader.GetText("数据", "A2"));
        Assert.Null(reader.GetText("数据", "B1"));
        Assert.False(reader.TryGetDecimal("数据", "B1", out _));
        Assert.Null(reader.GetFormula("数据", "B1"));
    }

    [Fact]
    public void MissingSheetThrowsWithTheRequestedName()
    {
        string path = TempFiles.Next("missing-sheet.xlsx");
        TestWorkbookBuilder.Create(path, new TestSheet("数据", []));
        using OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false);

        KeyNotFoundException exception = Assert.Throws<KeyNotFoundException>(
            () => reader.FindCell("不存在的工作表", "A1"));

        Assert.Contains("不存在的工作表", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SheetNamesIsAReadOnlySnapshot()
    {
        string path = TempFiles.Next("sheet-names.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet("数据", []),
            new TestSheet("省份 汇总", []));
        using OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false);
        IList<string> names = Assert.IsAssignableFrom<IList<string>>(reader.SheetNames);

        Assert.Throws<NotSupportedException>(() => names.Add("篡改"));
        Assert.Equal(["数据", "省份 汇总"], reader.SheetNames);
    }

    [Fact]
    public void SharedStringCellRejectsAMissingSharedStringTable()
    {
        string path = TempFiles.Next("missing-shared-strings.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet("数据", [TestCell.SharedText("A1", "北京市")]));
        RemoveSharedStringTable(path);
        using OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => reader.GetText("数据", "A1"));

        Assert.Contains("A1", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("not-an-index")]
    [InlineData("-1")]
    [InlineData("99")]
    public void SharedStringCellRejectsAnInvalidIndex(string rawIndex)
    {
        string path = TempFiles.Next("invalid-shared-string-index.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet("数据", [TestCell.SharedText("A1", "北京市")]));
        SetCellValue(path, "A1", rawIndex);
        using OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => reader.GetText("数据", "A1"));

        Assert.Contains("A1", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OpenRejectsAWorksheetWithAMissingRelationship()
    {
        string path = TempFiles.Next("missing-worksheet-relationship.xlsx");
        TestWorkbookBuilder.Create(path, new TestSheet("数据", []));
        string relationshipId = RemoveWorksheetRelationship(path);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
        {
            using OpenXmlWorkbookReader _ = OpenXmlWorkbookReader.Open(path, editable: false);
        });

        Assert.Contains("数据", exception.Message, StringComparison.Ordinal);
        Assert.Contains(relationshipId, exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("A1")]
    public void OpenRejectsMissingOrDuplicateCellReferences(string? replacementReference)
    {
        string path = TempFiles.Next("invalid-cell-reference.xlsx");
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "数据",
                [TestCell.Number("A1", "1"), TestCell.Number("A2", "2")]));
        SetCellReference(path, "A2", replacementReference);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
        {
            using OpenXmlWorkbookReader _ = OpenXmlWorkbookReader.Open(path, editable: false);
        });

        Assert.Contains("数据", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RepeatedReadsUseAnAddressIndex()
    {
        const int cellCount = 20_000;
        string path = TempFiles.Next("indexed-reads.xlsx");
        TestCell[] cells = Enumerable.Range(1, cellCount)
            .Select(row => TestCell.Number(
                $"A{row.ToString(CultureInfo.InvariantCulture)}",
                row.ToString(CultureInfo.InvariantCulture)))
            .ToArray();
        TestWorkbookBuilder.Create(path, new TestSheet("数据", cells));

        Stopwatch stopwatch = Stopwatch.StartNew();
        using (OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false))
        {
            for (int row = 1; row <= cellCount; row++)
            {
                string expected = row.ToString(CultureInfo.InvariantCulture);
                Assert.Equal(expected, reader.GetText("数据", $"A{expected}"));
            }
        }

        stopwatch.Stop();
        _output.WriteLine(
            $"Opened and read {cellCount.ToString(CultureInfo.InvariantCulture)} cells in " +
            $"{stopwatch.Elapsed.TotalMilliseconds.ToString("F0", CultureInfo.InvariantCulture)} ms.");
        Assert.True(
            stopwatch.Elapsed < TimeSpan.FromSeconds(5),
            $"Opening and reading {cellCount} cells took {stopwatch.Elapsed}.");
    }

    private static void MarkCellsAsBoolean(string path)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, isEditable: true);
        WorkbookPart workbookPart = Assert.IsType<WorkbookPart>(document.WorkbookPart);
        WorksheetPart worksheetPart = Assert.Single(workbookPart.WorksheetParts);
        Worksheet worksheet = Assert.IsType<Worksheet>(worksheetPart.Worksheet);

        foreach (Cell cell in worksheet.Descendants<Cell>())
        {
            cell.DataType = CellValues.Boolean;
        }

        worksheet.Save();
    }

    private static void RemoveSharedStringTable(string path)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, isEditable: true);
        WorkbookPart workbookPart = Assert.IsType<WorkbookPart>(document.WorkbookPart);
        SharedStringTablePart sharedStringsPart =
            Assert.IsType<SharedStringTablePart>(workbookPart.SharedStringTablePart);

        workbookPart.DeletePart(sharedStringsPart);
    }

    private static void SetCellValue(string path, string address, string value)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, isEditable: true);
        WorkbookPart workbookPart = Assert.IsType<WorkbookPart>(document.WorkbookPart);
        WorksheetPart worksheetPart = Assert.Single(workbookPart.WorksheetParts);
        Worksheet worksheet = Assert.IsType<Worksheet>(worksheetPart.Worksheet);
        Cell cell = Assert.Single(
            worksheet.Descendants<Cell>(),
            cell => string.Equals(cell.CellReference?.Value, address, StringComparison.Ordinal));
        cell.CellValue = new CellValue(value);
        worksheet.Save();
    }

    private static string RemoveWorksheetRelationship(string path)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, isEditable: true);
        WorkbookPart workbookPart = Assert.IsType<WorkbookPart>(document.WorkbookPart);
        Workbook workbook = Assert.IsType<Workbook>(workbookPart.Workbook);
        Sheets workbookSheets = Assert.IsType<Sheets>(workbook.GetFirstChild<Sheets>());
        Sheet sheet = Assert.Single(workbookSheets.Elements<Sheet>());
        string relationshipId = sheet.Id?.Value
            ?? throw new InvalidDataException("Test sheet relationship id is missing.");
        WorksheetPart worksheetPart = Assert.IsType<WorksheetPart>(
            workbookPart.GetPartById(relationshipId));

        workbookPart.DeletePart(worksheetPart);
        return relationshipId;
    }

    private static void SetCellReference(
        string path,
        string currentReference,
        string? replacementReference)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, isEditable: true);
        WorkbookPart workbookPart = Assert.IsType<WorkbookPart>(document.WorkbookPart);
        WorksheetPart worksheetPart = Assert.Single(workbookPart.WorksheetParts);
        Worksheet worksheet = Assert.IsType<Worksheet>(worksheetPart.Worksheet);
        Cell cell = Assert.Single(
            worksheet.Descendants<Cell>(),
            cell => string.Equals(
                cell.CellReference?.Value,
                currentReference,
                StringComparison.Ordinal));
        cell.CellReference = replacementReference;
        worksheet.Save();
    }
}
