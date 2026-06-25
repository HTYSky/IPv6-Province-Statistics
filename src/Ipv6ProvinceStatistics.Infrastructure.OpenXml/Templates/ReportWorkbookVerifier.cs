using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

public static class ReportWorkbookVerifier
{
    public static IReadOnlyList<ValidationIssue> Verify(string path, ProvinceReportInput input,
        IReadOnlyDictionary<string, decimal> calculated)
    {
        var issues = new List<ValidationIssue>();
        using var book = OpenXmlWorkbookReader.Open(path, false);

        foreach (var mapping in TemplateCellMap.Inputs)
        {
            if (!book.TryGetDecimal("Sheet1", mapping.Value, out var actual))
            {
                issues.Add(new("OUTPUT_INPUT_MISSING", $"输出表缺少输入值 {mapping.Value}。", Cell: mapping.Value));
            }
            else if (actual != input[mapping.Key])
            {
                issues.Add(new("OUTPUT_INPUT_MISMATCH",
                    $"输出表输入值 {mapping.Value} 不匹配。", Cell: mapping.Value));
            }
        }

        foreach (var cached in calculated)
        {
            if (!book.TryGetDecimal("Sheet1", cached.Key, out var actual))
            {
                issues.Add(new("OUTPUT_CACHE_MISSING", $"输出表缺少缓存值 {cached.Key}。", Cell: cached.Key));
            }
            else if (actual != cached.Value)
            {
                issues.Add(new("OUTPUT_CACHE_MISMATCH",
                    $"输出表缓存值 {cached.Key} 不匹配。", Cell: cached.Key));
            }
            else if (!string.Equals(book.GetFormula("Sheet1", cached.Key),
                FormulaManifest.Formulas[cached.Key], StringComparison.Ordinal))
            {
                issues.Add(new("OUTPUT_FORMULA_MISMATCH",
                    $"输出表公式 {cached.Key} 不匹配。", Cell: cached.Key));
            }
        }

        return issues;
    }
}
