using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

public sealed class TemplateReportExporter(TemplateResourceProvider provider) : ITemplateReportExporter
{
    public async Task<IReadOnlyList<ValidationIssue>> ExportAsync(
        Province province, ReportMonth month, ProvinceReportInput input,
        string outputPath, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var template = provider.ReadTemplate();
        await File.WriteAllBytesAsync(outputPath, template, cancellationToken);

        {
            using var document = SpreadsheetDocument.Open(outputPath, true);
            var workbook = document.WorkbookPart!;
            var sheet = workbook.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name == "Sheet1");
            var part = (WorksheetPart)workbook.GetPartById(sheet.Id!);

            foreach (var entry in TemplateCellMap.Inputs)
            {
                var cell = part.Worksheet!.Descendants<Cell>().Single(c =>
                    string.Equals(c.CellReference?.Value, entry.Value, StringComparison.OrdinalIgnoreCase));
                cell.CellValue = new CellValue(FormattableString.Invariant($"{input[entry.Key]}"));
                cell.DataType = CellValues.Number;
            }

            var calcResult = ReportCalculator.Calculate(input);
            if (calcResult.Issues.Count > 0)
                return calcResult.Issues;

            foreach (var (address, value) in calcResult.Values)
            {
                var cell = part.Worksheet!.Descendants<Cell>().Single(c =>
                    string.Equals(c.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
                cell.CellValue = new CellValue(FormattableString.Invariant($"{value}"));
                cell.DataType = CellValues.Number;
            }

            part.Worksheet!.Save();
        }
        return ReportWorkbookVerifier.Verify(outputPath, input, ReportCalculator.Calculate(input).Values);
    }

    public async Task<IReadOnlyList<ValidationIssue>> ValidateTemplateAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var tempPath = Path.GetTempFileName();
        try
        {
            var template = provider.ReadTemplate();
            await File.WriteAllBytesAsync(tempPath, template, cancellationToken);
            return TemplateValidator.Validate(tempPath, provider.ReadTemplateSha256());
        }
        finally
        {
            File.Delete(tempPath);
        }
    }
}
