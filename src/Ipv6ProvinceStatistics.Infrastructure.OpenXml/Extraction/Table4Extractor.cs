using System.Collections.ObjectModel;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table4Extractor
{
    public static SourceReadResult Extract(
        OpenXmlWorkbookReader reader,
        string path,
        CancellationToken cancellationToken)
    {
        var issues = new List<ValidationIssue>();
        var merged = new Dictionary<Province, Dictionary<MetricKey, decimal>>();
        string fileName = Path.GetFileName(path);

        ReadSheet("人网统计", MetricKey.HumanTotal, MetricKey.HumanIpv6);
        ReadSheet("物网统计", MetricKey.IotTotal, MetricKey.IotIpv6);

        var values = merged.ToDictionary(
            x => x.Key,
            x => (IReadOnlyDictionary<MetricKey, decimal>)new ReadOnlyDictionary<MetricKey, decimal>(x.Value));
        return new SourceReadResult(
            SourceWorkbookKind.Table4,
            issues.Count == 0 ? ExtractorSupport.AsReadOnly(values)
                : ExtractorSupport.AsReadOnly(new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>()),
            issues.AsReadOnly());

        void ReadSheet(string sheet, MetricKey totalKey, MetricKey ipv6Key)
        {
            var seen = new HashSet<Province>();
            foreach (uint row in reader.GetPopulatedRows(sheet).Where(row => row >= 5).Order())
            {
                cancellationToken.ThrowIfCancellationRequested();
                string provinceAddress = $"G{row}";
                string? provinceText = reader.GetText(sheet, provinceAddress);
                if (!ProvinceCatalog.TryResolve(provinceText, out Province province))
                {
                    continue;
                }

                if (!seen.Add(province))
                {
                    issues.Add(new ValidationIssue(
                        "PROVINCE_DUPLICATE",
                        $"{sheet} 的省份 {province.Name} 重复。",
                        fileName, sheet, province.Name));
                    continue;
                }

                int before = issues.Count;
                bool totalValid = ExtractorSupport.TryReadDecimal(
                    reader, path, sheet, $"R{row}", province, issues, out decimal total);
                bool ipv6Valid = ExtractorSupport.TryReadDecimal(
                    reader, path, sheet, $"S{row}", province, issues, out decimal ipv6);

                if (!totalValid || !ipv6Valid)
                {
                    continue;
                }

                if (!merged.TryGetValue(province, out var metrics))
                {
                    merged[province] = metrics = [];
                }

                metrics[totalKey] = total;
                metrics[ipv6Key] = ipv6;
            }
        }
    }
}
