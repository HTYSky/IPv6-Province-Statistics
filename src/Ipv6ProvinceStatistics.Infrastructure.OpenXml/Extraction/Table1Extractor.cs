using System.Collections.ObjectModel;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table1Extractor
{
    private const string Sheet = "分省统计表";

    public static SourceReadResult Extract(
        OpenXmlWorkbookReader reader,
        string path,
        CancellationToken cancellationToken)
    {
        var issues = new List<ValidationIssue>();
        var values = new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>();
        Province currentProvince = default;
        bool hasCurrentProvince = false;

        foreach (uint row in reader.GetPopulatedRows(Sheet).Where(row => row >= 4).Order())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (ProvinceCatalog.TryResolve(reader.GetText(Sheet, $"B{row}"), out Province province))
            {
                currentProvince = province;
                hasCurrentProvince = true;
            }

            string carrierAddress = $"C{row}";
            if (!string.Equals(
                    HeaderText.Normalize(reader.GetText(Sheet, carrierAddress)),
                    "中国联通",
                    StringComparison.Ordinal))
            {
                continue;
            }

            if (!hasCurrentProvince)
            {
                issues.Add(new ValidationIssue(
                    "PROVINCE_GROUP_MISSING",
                    $"联通数据行 {Sheet}!{carrierAddress} 之前没有可识别的省份分组。",
                    Path.GetFileName(path),
                    Sheet,
                    Cell: carrierAddress));
                continue;
            }

            bool metroTotalValid = ExtractorSupport.TryReadDecimal(
                reader, path, Sheet, $"D{row}", currentProvince, issues, out decimal metroTotal);
            bool metroIpv6Valid = ExtractorSupport.TryReadDecimal(
                reader, path, Sheet, $"E{row}", currentProvince, issues, out decimal metroIpv6);
            bool mobileCoreTotalValid = ExtractorSupport.TryReadDecimal(
                reader, path, Sheet, $"G{row}", currentProvince, issues, out decimal mobileCoreTotal);
            bool mobileCoreIpv6Valid = ExtractorSupport.TryReadDecimal(
                reader, path, Sheet, $"H{row}", currentProvince, issues, out decimal mobileCoreIpv6);

            if (!(metroTotalValid & metroIpv6Valid & mobileCoreTotalValid & mobileCoreIpv6Valid))
            {
                continue;
            }

            var metrics = new ReadOnlyDictionary<MetricKey, decimal>(
                new Dictionary<MetricKey, decimal>
                {
                    [MetricKey.MetroTotal] = metroTotal,
                    [MetricKey.MetroIpv6] = metroIpv6,
                    [MetricKey.MobileCoreTotal] = mobileCoreTotal,
                    [MetricKey.MobileCoreIpv6] = mobileCoreIpv6,
                });
            ExtractorSupport.AddProvince(values, path, Sheet, currentProvince, metrics, issues);
        }

        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyDictionary<Province, IReadOnlyDictionary<MetricKey, decimal>> resultValues =
            issues.Count == 0
                ? ExtractorSupport.AsReadOnly(values)
                : ExtractorSupport.AsReadOnly(
                    new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>());
        return new SourceReadResult(SourceWorkbookKind.Table1, resultValues, issues.AsReadOnly());
    }
}
