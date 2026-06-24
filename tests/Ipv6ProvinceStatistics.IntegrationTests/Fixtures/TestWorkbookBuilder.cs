using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

internal enum TestCellKind
{
    SharedText,
    InlineText,
    Number,
    Formula,
}

internal sealed record TestCell(
    string Address,
    TestCellKind Kind,
    string Value,
    string? CachedValue = null)
{
    public static TestCell SharedText(string address, string value) =>
        new(address, TestCellKind.SharedText, value);

    public static TestCell InlineText(string address, string value) =>
        new(address, TestCellKind.InlineText, value);

    public static TestCell Number(string address, string value) =>
        new(address, TestCellKind.Number, value);

    public static TestCell Formula(string address, string formula, string cachedValue) =>
        new(address, TestCellKind.Formula, formula, cachedValue);
}

internal sealed record TestSheet(string Name, IReadOnlyList<TestCell> Cells);

internal static class TestWorkbookBuilder
{
    public static void Create(string path, params TestSheet[] sheets)
    {
        string? directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using SpreadsheetDocument document = SpreadsheetDocument.Create(
            path,
            SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        SharedStringTablePart sharedStringsPart =
            workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringsPart.SharedStringTable = new SharedStringTable();
        var workbookSheets = new Sheets();
        workbookPart.Workbook.Append(workbookSheets);

        uint sheetId = 1;
        foreach (TestSheet sheet in sheets)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            HashSet<string> addresses = new(StringComparer.OrdinalIgnoreCase);
            var orderedCells = sheet.Cells
                .Select(cell => (Cell: cell, Address: ParseAddress(cell.Address)))
                .OrderBy(item => item.Address.Row)
                .ThenBy(item => item.Address.Column)
                .ToArray();

            foreach (IGrouping<uint, (TestCell Cell, ParsedAddress Address)> rowGroup in
                     orderedCells.GroupBy(item => item.Address.Row))
            {
                var row = new Row { RowIndex = rowGroup.Key };
                foreach ((TestCell cell, ParsedAddress address) in rowGroup)
                {
                    if (!addresses.Add(address.Reference))
                    {
                        throw new ArgumentException(
                            $"Duplicate cell address '{cell.Address}'.",
                            nameof(sheets));
                    }

                    row.Append(CreateCell(cell, address.Reference, sharedStringsPart));
                }

                sheetData.Append(row);
            }

            worksheetPart.Worksheet.Save();
            workbookSheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = sheet.Name,
            });
        }

        sharedStringsPart.SharedStringTable.Save();
        workbookPart.Workbook.Save();
    }

    private static Cell CreateCell(
        TestCell cell,
        string reference,
        SharedStringTablePart sharedStringsPart) =>
        cell.Kind switch
        {
            TestCellKind.SharedText => CreateSharedTextCell(
                reference,
                cell.Value,
                sharedStringsPart),
            TestCellKind.InlineText => new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(CreateText(cell.Value)),
            },
            TestCellKind.Number => new Cell
            {
                CellReference = reference,
                DataType = CellValues.Number,
                CellValue = new CellValue(cell.Value),
            },
            TestCellKind.Formula => new Cell
            {
                CellReference = reference,
                CellFormula = new CellFormula(cell.Value),
                CellValue = cell.CachedValue is null ? null : new CellValue(cell.CachedValue),
            },
            _ => throw new ArgumentOutOfRangeException(nameof(cell)),
        };

    private static Cell CreateSharedTextCell(
        string reference,
        string value,
        SharedStringTablePart sharedStringsPart)
    {
        SharedStringTable sharedStrings = sharedStringsPart.SharedStringTable
            ?? throw new InvalidOperationException("Shared string table was not initialized.");
        int index = sharedStrings.ChildElements.Count;
        sharedStrings.AppendChild(new SharedStringItem(CreateText(value)));

        return new Cell
        {
            CellReference = reference,
            DataType = CellValues.SharedString,
            CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture)),
        };
    }

    private static Text CreateText(string value) =>
        new(value) { Space = SpaceProcessingModeValues.Preserve };

    private static ParsedAddress ParseAddress(string address)
    {
        if (string.IsNullOrEmpty(address))
        {
            throw new ArgumentException("Cell address is required.", nameof(address));
        }

        string normalized = address.ToUpperInvariant();
        int digitStart = 0;
        while (digitStart < normalized.Length &&
               normalized[digitStart] is >= 'A' and <= 'Z')
        {
            digitStart++;
        }

        if (digitStart == 0 || digitStart == normalized.Length ||
            !uint.TryParse(
                normalized.AsSpan(digitStart),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out uint row) ||
            row == 0)
        {
            throw new ArgumentException($"Invalid cell address '{address}'.", nameof(address));
        }

        uint column = 0;
        foreach (char character in normalized.AsSpan(0, digitStart))
        {
            column = checked((column * 26) + (uint)(character - 'A' + 1));
        }

        string reference = string.Concat(
            normalized.AsSpan(0, digitStart),
            row.ToString(CultureInfo.InvariantCulture));
        return new ParsedAddress(row, column, reference);
    }

    private readonly record struct ParsedAddress(uint Row, uint Column, string Reference);
}
