using System.Collections.ObjectModel;
using System.Globalization;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class ExtractorSupport
{
    public static bool TryReadDecimal(
        OpenXmlWorkbookReader reader,
        string file,
        string sheet,
        string address,
        Province province,
        ICollection<ValidationIssue> issues,
        out decimal value)
    {
        string? text = reader.GetText(sheet, address);
        if (decimal.TryParse(
                text,
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture,
                out value))
        {
            return true;
        }

        issues.Add(new ValidationIssue(
            "VALUE_INVALID",
            $"单元格 {sheet}!{address} 的值不是有效数字。",
            Path.GetFileName(file),
            sheet,
            province.Name,
            address));
        return false;
    }

    public static bool AddProvince(
        Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>> values,
        string file,
        string sheet,
        Province province,
        IReadOnlyDictionary<MetricKey, decimal> metrics,
        ICollection<ValidationIssue> issues)
    {
        if (values.TryAdd(province, metrics))
        {
            return true;
        }

        issues.Add(new ValidationIssue(
            "PROVINCE_DUPLICATE",
            $"工作表 '{sheet}' 中存在重复的 {province.Name} 数据。",
            Path.GetFileName(file),
            sheet,
            province.Name));
        return false;
    }

    public static IReadOnlyDictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>
        AsReadOnly(
            Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>> values) =>
        new ReadOnlyDictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>(values);
}
