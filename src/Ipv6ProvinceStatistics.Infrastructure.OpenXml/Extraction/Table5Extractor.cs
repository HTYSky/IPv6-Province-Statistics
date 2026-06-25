using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table5Extractor
{
    private const string InternetSheet = "互联网专线汇总";
    private const string IdcSheetNormalized = "IDC汇总(客户)";

    public static SourceReadResult Extract(
        OpenXmlWorkbookReader reader,
        string path,
        CancellationToken cancellationToken)
    {
        string idcSheet = FindIdcSheet(reader);
        var issues = new List<ValidationIssue>();
        var merged = new Dictionary<Province, Dictionary<MetricKey, decimal>>();

        ReadSheet(InternetSheet, "M", 3, new()
        {
            ["E"] = MetricKey.InternetTotal,
            ["F"] = MetricKey.InternetIpv6,
            ["K"] = MetricKey.InternetOneGTotal,
            ["L"] = MetricKey.InternetOneGIpv6,
        });
        cancellationToken.ThrowIfCancellationRequested();
        ReadSheet(idcSheet, "O", 3, new()
        {
            ["F"] = MetricKey.IdcTotal,
            ["G"] = MetricKey.IdcIpv6,
            ["M"] = MetricKey.IdcTenGTotal,
            ["N"] = MetricKey.IdcTenGIpv6,
        });

        IReadOnlyDictionary<Province, IReadOnlyDictionary<MetricKey, decimal>> values;
        if (issues.Count == 0)
        {
            values = merged.ToDictionary(
                kvp => kvp.Key,
                kvp => (IReadOnlyDictionary<MetricKey, decimal>)new ReadOnlyDictionary<MetricKey, decimal>(
                    kvp.Value));
        }
        else
        {
            values = new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>();
        }

        return new SourceReadResult(SourceWorkbookKind.Table5, values, issues.AsReadOnly());

        void ReadSheet(string sheetName, string provinceColumn, uint firstRow,
            IReadOnlyDictionary<string, MetricKey> columnMap)
        {
            var seen = new HashSet<Province>();
            foreach (uint row in reader.GetPopulatedRows(sheetName).Where(row => row >= firstRow).Order())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!ProvinceCatalog.TryResolve(
                        reader.GetText(sheetName, $"{provinceColumn}{row}"),
                        out Province province)) continue;
                if (!seen.Add(province))
                {
                    issues.Add(new ValidationIssue("PROVINCE_DUPLICATE",
                        $"工作表 {sheetName} 中省份 {province.Name} 出现重复。",
                        Path.GetFileName(path), sheetName, province.Name));
                    continue;
                }
                if (!merged.TryGetValue(province, out Dictionary<MetricKey, decimal>? metrics))
                {
                    metrics = [];
                    merged[province] = metrics;
                }
                bool allValid = true;
                foreach (var mapping in columnMap)
                {
                    if (!ExtractorSupport.TryReadDecimal(reader, path, sheetName,
                            $"{mapping.Key}{row}", province, issues, out decimal value))
                        allValid = false;
                    else metrics[mapping.Value] = value;
                }
                if (!allValid) merged.Remove(province);
            }
        }
    }

    private static string FindIdcSheet(OpenXmlWorkbookReader reader)
    {
        foreach (string name in reader.SheetNames)
            if (string.Equals(HeaderText.Normalize(name), IdcSheetNormalized, StringComparison.OrdinalIgnoreCase))
                return name;
        throw new InvalidOperationException("5 表中未找到 IDC 汇总工作表。");
    }
}
