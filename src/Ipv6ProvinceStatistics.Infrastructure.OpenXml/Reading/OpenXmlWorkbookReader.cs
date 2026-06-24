using System.Collections.ObjectModel;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

internal sealed class OpenXmlWorkbookReader : IDisposable
{
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbookPart;
    private readonly Dictionary<string, WorksheetPart> _worksheets;

    private OpenXmlWorkbookReader(SpreadsheetDocument document)
    {
        _document = document;
        _workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("XLSX package does not contain a workbook part.");
        Workbook workbook = _workbookPart.Workbook
            ?? throw new InvalidDataException("Workbook part does not contain a workbook.");

        Sheet[] sheets = workbook.GetFirstChild<Sheets>()?
            .Elements<Sheet>()
            .ToArray()
            ?? [];

        _worksheets = new Dictionary<string, WorksheetPart>(StringComparer.Ordinal);
        foreach (Sheet sheet in sheets)
        {
            string name = sheet.Name?.Value
                ?? throw new InvalidDataException("Workbook contains a sheet without a name.");
            string relationshipId = sheet.Id?.Value
                ?? throw new InvalidDataException($"Worksheet '{name}' has no relationship id.");
            if (_workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                throw new InvalidDataException($"Sheet '{name}' does not reference a worksheet part.");
            }

            _worksheets.Add(name, worksheetPart);
        }

        SheetNames = new ReadOnlyCollection<string>(sheets.Select(GetSheetName).ToArray());
    }

    public IReadOnlyList<string> SheetNames { get; }

    public static OpenXmlWorkbookReader Open(string path, bool editable)
    {
        SpreadsheetDocument document = SpreadsheetDocument.Open(path, editable);
        try
        {
            return new OpenXmlWorkbookReader(document);
        }
        catch
        {
            document.Dispose();
            throw;
        }
    }

    public bool HasSheet(string sheetName) => _worksheets.ContainsKey(sheetName);

    public Cell? FindCell(string sheetName, string address)
    {
        WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
        return GetWorksheet(worksheetPart, sheetName)
            .Descendants<Cell>()
            .FirstOrDefault(cell => string.Equals(
                cell.CellReference?.Value,
                address,
                StringComparison.OrdinalIgnoreCase));
    }

    public string? GetText(string sheetName, string address)
    {
        Cell? cell = FindCell(sheetName, address);
        if (cell is null)
        {
            return null;
        }

        CellValues? dataType = cell.DataType?.Value;
        string? text;
        if (dataType == CellValues.SharedString)
        {
            text = GetSharedString(cell);
        }
        else if (dataType == CellValues.InlineString)
        {
            text = cell.InlineString?.InnerText;
        }
        else if (dataType == CellValues.Boolean)
        {
            text = cell.CellValue?.InnerText switch
            {
                "1" => "TRUE",
                "0" => "FALSE",
                string value => value,
                null => null,
            };
        }
        else
        {
            text = cell.CellValue?.InnerText;
        }

        return string.IsNullOrEmpty(text) ? null : text;
    }

    public bool TryGetDecimal(string sheetName, string address, out decimal value) =>
        decimal.TryParse(
            GetText(sheetName, address),
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out value);

    public string? GetFormula(string sheetName, string address) =>
        FindCell(sheetName, address)?.CellFormula?.InnerText;

    public IReadOnlyList<uint> GetPopulatedRows(string sheetName)
    {
        WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
        uint[] rowNumbers = GetWorksheet(worksheetPart, sheetName)
            .Descendants<Row>()
            .Where(row => row.Elements<Cell>().Any(IsPopulated))
            .Select(row => row.RowIndex?.Value)
            .Where(rowIndex => rowIndex.HasValue)
            .Select(rowIndex => rowIndex.GetValueOrDefault())
            .ToArray();

        return Array.AsReadOnly(rowNumbers);
    }

    public void Dispose() => _document.Dispose();

    private static string GetSheetName(Sheet sheet) =>
        sheet.Name?.Value
        ?? throw new InvalidDataException("Workbook contains a sheet without a name.");

    private static bool IsPopulated(Cell cell) =>
        cell.CellValue is not null ||
        cell.InlineString is not null ||
        cell.CellFormula is not null;

    private static Worksheet GetWorksheet(WorksheetPart worksheetPart, string sheetName) =>
        worksheetPart.Worksheet
        ?? throw new InvalidDataException($"Worksheet part for sheet '{sheetName}' is empty.");

    private WorksheetPart GetWorksheetPart(string sheetName)
    {
        if (!_worksheets.TryGetValue(sheetName, out WorksheetPart? worksheetPart))
        {
            throw new KeyNotFoundException($"Workbook does not contain sheet '{sheetName}'.");
        }

        return worksheetPart;
    }

    private string? GetSharedString(Cell cell)
    {
        if (!int.TryParse(
                cell.CellValue?.InnerText,
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out int index) ||
            index < 0)
        {
            return null;
        }

        SharedStringTable? sharedStrings =
            _workbookPart.SharedStringTablePart?.SharedStringTable;

        return sharedStrings?
            .Elements<SharedStringItem>()
            .ElementAtOrDefault(index)?
            .InnerText;
    }
}
