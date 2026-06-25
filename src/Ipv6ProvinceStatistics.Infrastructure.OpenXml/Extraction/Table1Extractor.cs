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
        var seenProvinces = new HashSet<Province>();
        Province currentProvince = default;
        bool hasCurrentProvince = false;

        foreach (uint row in reader.GetPopulatedRows(Sheet).Where(row => row >= 4).Order())
        {
            cancellationToken.ThrowIfCancellationRequested();
            string provinceAddress = $"B{row}";
            string? provinceText = reader.GetText(Sheet, provinceAddress);
            if (provinceText is not null &&
                ProvinceCatalog.TryResolve(provinceText, out Province province))
            {
                currentProvince = province;
                hasCurrentProvince = true;
            }
            else if (provinceText is not null)
            {
                currentProvince = default;
                hasCurrentProvince = false;
                issues.Add(new ValidationIssue(
                    "PROVINCE_INVALID",
                    $"省份单元格 {Sheet}!{provinceAddress} 包含无法识别的省份。",
                    Path.GetFileName(path),
                    Sheet,
                    Cell: provinceAddress));
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

            bool isDuplicate = !seenProvinces.Add(currentProvince);
            if (isDuplicate)
            {
                string dupCell = $"C{row}";
                issues.Add(new ValidationIssue(
                    "PROVINCE_DUPLICATE",
                    $"省份 {currentProvince.Name} 在 {Sheet}!{dupCell} 存在重复的联通数据行。",
                    Path.GetFileName(path),
                    Sheet,
                    currentProvince.Name,
                    Cell: dupCell));
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

            if (isDuplicate)
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
