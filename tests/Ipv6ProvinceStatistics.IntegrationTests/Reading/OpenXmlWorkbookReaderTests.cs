using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Reading;

public sealed class OpenXmlWorkbookReaderTests
{
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
}
