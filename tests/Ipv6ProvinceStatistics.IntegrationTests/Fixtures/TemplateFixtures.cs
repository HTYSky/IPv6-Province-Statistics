using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

public static class TemplateFixtures
{
    public static string CreatePopulatedTemplate()
    {
        var path = TempFiles.Next("populated-template.xlsx");
        var inputs = TemplateCellMap.Inputs.Values
            .Select(address => TestCell.Number(address, "999"));
        var formulas = FormulaManifest.Formulas
            .Select(pair => TestCell.Formula(pair.Key, pair.Value, "1"));
        TestWorkbookBuilder.Create(path,
            new TestSheet("Sheet1", inputs.Concat(formulas).ToArray()));

        using var document = SpreadsheetDocument.Open(path, true);
        var styles = document.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1U },
            new Fills(new Fill()) { Count = 1U },
            new Borders(new Border()) { Count = 1U },
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            new CellFormats(new CellFormat()) { Count = 1U });
        styles.Stylesheet.Save();
        return path;
    }
}
