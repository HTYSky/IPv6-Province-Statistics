using System.Collections.ObjectModel;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table8Extractor
{
    public static SourceReadResult Extract(
        OpenXmlWorkbookReader reader,
        string path,
        CancellationToken cancellationToken)
    {
        var issues = new List<ValidationIssue>();
        var values = new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>();
        var seen = new HashSet<Province>();
        string fileName = Path.GetFileName(path);
        string? sheet = reader.SheetNames.FirstOrDefault(
            name => name == "省统计" || name.EndsWith("-省统计", StringComparison.Ordinal));

        if (sheet is null)
        {
            return new SourceReadResult(SourceWorkbookKind.Table8,
                ExtractorSupport.AsReadOnly(new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>()),
                [new ValidationIssue("SHEET_MISSING", "未找到扩频统计 sheet。", fileName)]);
        }

        foreach (uint row in reader.GetPopulatedRows(sheet).Where(row => row >= 4).Order())
        {
            cancellationToken.ThrowIfCancellationRequested();
            string provinceAddress = $"B{row}";
            string? provinceText = reader.GetText(sheet, provinceAddress);
            if (!ProvinceCatalog.TryResolve(provinceText, out Province province)) continue;

            if (!seen.Add(province))
            {
                issues.Add(new ValidationIssue("PROVINCE_DUPLICATE",
                    $"{sheet} 的省份 {province.Name} 重复。", fileName, sheet, province.Name));
                continue;
            }

            int before = issues.Count;
            bool ok1 = ExtractorSupport.TryReadDecimal(reader, path, sheet, $"D{row}", province, issues, out decimal total);
            bool ok2 = ExtractorSupport.TryReadDecimal(reader, path, sheet, $"J{row}", province, issues, out decimal ipv6);
            if (!ok1 || !ok2) continue;

            var metrics = new ReadOnlyDictionary<MetricKey, decimal>(
                new Dictionary<MetricKey, decimal>
                {
                    [MetricKey.BroadbandTotal] = total,
                    [MetricKey.BroadbandIpv6] = ipv6
                });
            ExtractorSupport.AddProvince(values, path, sheet, province, metrics, issues);
        }

        return new SourceReadResult(SourceWorkbookKind.Table8,
            issues.Count == 0 ? ExtractorSupport.AsReadOnly(values)
                : ExtractorSupport.AsReadOnly(new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>()),
            issues.AsReadOnly());
    }
}
