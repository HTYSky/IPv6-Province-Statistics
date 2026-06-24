using System.Collections.ObjectModel;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

internal sealed class OpenXmlWorkbookReader : IDisposable
{
    private readonly SpreadsheetDocument _document;
    private readonly Dictionary<string, WorksheetCache> _worksheets;
    private readonly string[]? _sharedStrings;

    private OpenXmlWorkbookReader(SpreadsheetDocument document)
    {
        _document = document;
        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new InvalidDataException("XLSX package does not contain a workbook part.");
        Workbook workbook = workbookPart.Workbook
            ?? throw new InvalidDataException("Workbook part does not contain a workbook.");

        Sheet[] sheets = workbook.GetFirstChild<Sheets>()?
            .Elements<Sheet>()
            .ToArray()
            ?? [];

        _worksheets = new Dictionary<string, WorksheetCache>(StringComparer.Ordinal);
        foreach (Sheet sheet in sheets)
        {
            string name = sheet.Name?.Value
                ?? throw new InvalidDataException("Workbook contains a sheet without a name.");
            string relationshipId = sheet.Id?.Value ?? "<missing>";
            if (sheet.Id?.Value is not string existingRelationshipId ||
                !workbookPart.TryGetPartById(existingRelationshipId, out OpenXmlPart? part) ||
                part is not WorksheetPart worksheetPart)
            {
                throw new InvalidDataException(
                    $"Sheet '{name}' relationship '{relationshipId}' does not reference a worksheet part.");
            }

            Worksheet worksheet = GetWorksheet(worksheetPart, name);
            var cells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (Cell cell in worksheet.Descendants<Cell>())
            {
                string? reference = cell.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(reference))
                {
                    throw new InvalidDataException(
                        $"Sheet '{name}' contains a cell without a cell reference.");
                }

                if (!cells.TryAdd(reference, cell))
                {
                    throw new InvalidDataException(
                        $"Sheet '{name}' contains duplicate cell reference '{reference}'.");
                }
            }

            _worksheets.Add(name, new WorksheetCache(worksheetPart, cells));
        }

        _sharedStrings = workbookPart.SharedStringTablePart?
            .SharedStringTable?
            .Elements<SharedStringItem>()
            .Select(item => item.InnerText)
            .ToArray();
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
        WorksheetCache worksheet = GetWorksheetCache(sheetName);
        return worksheet.Cells.TryGetValue(address, out Cell? cell) ? cell : null;
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
            text = GetSharedString(cell, sheetName, address);
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
        WorksheetCache worksheet = GetWorksheetCache(sheetName);
        uint[] rowNumbers = GetWorksheet(worksheet.Part, sheetName)
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

    private WorksheetCache GetWorksheetCache(string sheetName)
    {
        if (!_worksheets.TryGetValue(sheetName, out WorksheetCache? worksheet))
        {
            throw new KeyNotFoundException($"Workbook does not contain sheet '{sheetName}'.");
        }

        return worksheet;
    }

    private string GetSharedString(Cell cell, string sheetName, string address)
    {
        if (!int.TryParse(
                cell.CellValue?.InnerText,
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out int index) ||
            index < 0)
        {
            throw InvalidSharedString(
                sheetName,
                address,
                "the shared-string index is not a non-negative integer");
        }

        if (_sharedStrings is null)
        {
            throw InvalidSharedString(sheetName, address, "the shared-string table is missing");
        }

        if (index >= _sharedStrings.Length)
        {
            throw InvalidSharedString(
                sheetName,
                address,
                $"shared-string index {index} is out of range");
        }

        return _sharedStrings[index];
    }

    private static InvalidDataException InvalidSharedString(
        string sheetName,
        string address,
        string reason) =>
        new($"Invalid shared-string cell '{sheetName}'!{address}: {reason}.");

    private sealed record WorksheetCache(
        WorksheetPart Part,
        IReadOnlyDictionary<string, Cell> Cells);
}
