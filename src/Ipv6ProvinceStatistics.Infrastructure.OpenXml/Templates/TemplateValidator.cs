using System.Security.Cryptography;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

public static class TemplateValidator
{
    public static IReadOnlyList<ValidationIssue> Validate(string path, string? expectedSha256 = null)
    {
        var issues = new List<ValidationIssue>();
        if (expectedSha256 is not null)
        {
            using var stream = File.OpenRead(path);
            var actual = Convert.ToHexString(SHA256.HashData(stream));
            if (!actual.Equals(expectedSha256.Trim(), StringComparison.OrdinalIgnoreCase))
                issues.Add(new("TEMPLATE_HASH", "内置模板哈希不匹配。"));
        }
        using var book = OpenXmlWorkbookReader.Open(path, false);
        if (!book.HasSheet("Sheet1"))
            return new ValidationIssue[] { new("TEMPLATE_SHEET", "内置模板缺少 Sheet1。") };
        foreach (var address in TemplateCellMap.Inputs.Values)
        {
            if (book.FindCell("Sheet1", address) is null)
                issues.Add(new("TEMPLATE_INPUT_MISSING", $"模板缺少输入单元格 {address}。", Cell: address));
            else if (book.GetFormula("Sheet1", address) is not null)
                issues.Add(new("TEMPLATE_INPUT_FORMULA", $"输入单元格 {address} 不应包含公式。", Cell: address));
        }
        foreach (var expected in FormulaManifest.Formulas)
            if (!string.Equals(book.GetFormula("Sheet1", expected.Key), expected.Value, StringComparison.Ordinal))
                issues.Add(new("TEMPLATE_FORMULA", $"公式 {expected.Key} 与版本清单不一致。", Cell: expected.Key));
        return issues;
    }
}
