using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

public static class TemplateSanitizer
{
    public static void Sanitize(string input, string output)
    {
        File.Copy(input, output, true);
        using var document = SpreadsheetDocument.Open(output, true);
        var workbook = document.WorkbookPart!;
        var sheet = workbook.Workbook!.Sheets!.Elements<Sheet>().Single(x => x.Name == "Sheet1");
        var part = (WorksheetPart)workbook.GetPartById(sheet.Id!);
        foreach (var address in TemplateCellMap.Inputs.Values.Concat(FormulaManifest.Formulas.Keys))
        {
            var cell = part.Worksheet!.Descendants<Cell>().Single(x =>
                string.Equals(x.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
            cell.CellValue = null;
            if (TemplateCellMap.Inputs.Values.Contains(address, StringComparer.OrdinalIgnoreCase))
                cell.DataType = null;
        }
        part.Worksheet!.Save();
        var issues = TemplateValidator.Validate(output);
        if (issues.Count > 0)
            throw new InvalidDataException(string.Join(Environment.NewLine, issues.Select(x => x.Message)));
    }
}
