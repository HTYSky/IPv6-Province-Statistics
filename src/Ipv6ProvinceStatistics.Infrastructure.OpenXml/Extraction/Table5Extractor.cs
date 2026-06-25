using System.Collections.ObjectModel;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table5Extractor
{
    private const string InternetSheet = "互联网专线汇总";

    public static SourceReadResult Extract(
        OpenXmlWorkbookReader reader,
        string path,
        CancellationToken cancellationToken)
    {
        var issues = new List<ValidationIssue>();
        var merged = new Dictionary<Province, Dictionary<MetricKey, decimal>>();
        string fileName = Path.GetFileName(path);

        ReadInternetSheet();
        string? idcSheet = FindIdcSheet(reader);
        if (idcSheet is not null)
        {
            ReadIdcSheet(idcSheet);
        }

        var values = merged.ToDictionary(
            x => x.Key,
            x => (IReadOnlyDictionary<MetricKey, decimal>)new ReadOnlyDictionary<MetricKey, decimal>(x.Value));
        return new SourceReadResult(
            SourceWorkbookKind.Table5,
            issues.Count == 0 ? ExtractorSupport.AsReadOnly(values)
                : ExtractorSupport.AsReadOnly(new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>()),
            issues.AsReadOnly());

        void ReadInternetSheet()
        {
            var seen = new HashSet<Province>();
            foreach (uint row in reader.GetPopulatedRows(InternetSheet).Where(row => row >= 3).Order())
            {
                cancellationToken.ThrowIfCancellationRequested();
                string provinceAddress = $"M{row}";
                string? provinceText = reader.GetText(InternetSheet, provinceAddress);
                if (!ProvinceCatalog.TryResolve(provinceText, out Province province))
                {
                    continue;
                }
                if (!seen.Add(province))
                {
                    issues.Add(new ValidationIssue(
                        "PROVINCE_DUPLICATE",
                        $"{InternetSheet} 的省份 {province.Name} 重复。",
                        fileName, InternetSheet, province.Name));
                    continue;
                }
                int before = issues.Count;
                bool ok1 = ExtractorSupport.TryReadDecimal(reader, path, InternetSheet, $"E{row}", province, issues, out decimal t1);
                bool ok2 = ExtractorSupport.TryReadDecimal(reader, path, InternetSheet, $"F{row}", province, issues, out decimal t2);
                bool ok3 = ExtractorSupport.TryReadDecimal(reader, path, InternetSheet, $"K{row}", province, issues, out decimal t3);
                bool ok4 = ExtractorSupport.TryReadDecimal(reader, path, InternetSheet, $"L{row}", province, issues, out decimal t4);
                if (!ok1 || !ok2 || !ok3 || !ok4) continue;
                if (!merged.TryGetValue(province, out var metrics)) merged[province] = metrics = [];
                metrics[MetricKey.InternetTotal] = t1;
                metrics[MetricKey.InternetIpv6] = t2;
                metrics[MetricKey.InternetOneGTotal] = t3;
                metrics[MetricKey.InternetOneGIpv6] = t4;
            }
        }

        void ReadIdcSheet(string sheet)
        {
            var seen = new HashSet<Province>();
            foreach (uint row in reader.GetPopulatedRows(sheet).Where(row => row >= 3).Order())
            {
                cancellationToken.ThrowIfCancellationRequested();
                string provinceAddress = $"O{row}";
                string? provinceText = reader.GetText(sheet, provinceAddress);
                if (!ProvinceCatalog.TryResolve(provinceText, out Province province)) continue;
                if (!seen.Add(province))
                {
                    issues.Add(new ValidationIssue(
                        "PROVINCE_DUPLICATE", $"{sheet} 的省份 {province.Name} 重复。",
                        fileName, sheet, province.Name));
                    continue;
                }
                int before = issues.Count;
                bool ok1 = ExtractorSupport.TryReadDecimal(reader, path, sheet, $"F{row}", province, issues, out decimal t1);
                bool ok2 = ExtractorSupport.TryReadDecimal(reader, path, sheet, $"G{row}", province, issues, out decimal t2);
                bool ok3 = ExtractorSupport.TryReadDecimal(reader, path, sheet, $"M{row}", province, issues, out decimal t3);
                bool ok4 = ExtractorSupport.TryReadDecimal(reader, path, sheet, $"N{row}", province, issues, out decimal t4);
                if (!ok1 || !ok2 || !ok3 || !ok4) continue;
                if (!merged.TryGetValue(province, out var metrics)) merged[province] = metrics = [];
                metrics[MetricKey.IdcTotal] = t1;
                metrics[MetricKey.IdcIpv6] = t2;
                metrics[MetricKey.IdcTenGTotal] = t3;
                metrics[MetricKey.IdcTenGIpv6] = t4;
            }
        }
    }

    private static string? FindIdcSheet(OpenXmlWorkbookReader reader)
    {
        foreach (var name in reader.SheetNames)
        {
            if (HeaderText.Normalize(name) == "IDC汇总(客户)") return name;
        }
        return null;
    }
}
