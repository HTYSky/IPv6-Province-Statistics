using System.Collections.ObjectModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class ExtractorSupport
{
    private static readonly Regex DecimalTextPattern = new(
        @"^[+-]?(?:(?:[0-9]{1,3}(?:,[0-9]{3})+)|[0-9]+)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?$",
        RegexOptions.CultureInvariant | RegexOptions.NonBacktracking);

    public static bool TryReadDecimal(
        OpenXmlWorkbookReader reader,
        string file,
        string sheet,
        string address,
        Province province,
        ICollection<ValidationIssue> issues,
        out decimal value)
    {
        if (TryParseLosslessDecimal(reader.GetText(sheet, address), out value))
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

    private static bool TryParseLosslessDecimal(string? text, out decimal value)
    {
        value = default;
        if (text is null)
        {
            return false;
        }

        string normalized = text.Trim();
        if (!DecimalTextPattern.IsMatch(normalized) ||
            !decimal.TryParse(
                normalized,
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture,
                out value) ||
            !TryCanonicalize(normalized, out CanonicalDecimal source) ||
            !TryCanonicalize(
                value.ToString("G29", CultureInfo.InvariantCulture),
                out CanonicalDecimal parsed))
        {
            return false;
        }

        return source == parsed;
    }

    private static bool TryCanonicalize(string text, out CanonicalDecimal canonical)
    {
        canonical = default;
        int mantissaStart = text[0] is '+' or '-' ? 1 : 0;
        bool isNegative = text[0] == '-';
        int exponentMarker = text.IndexOfAny('e', 'E');
        int mantissaEnd = exponentMarker < 0 ? text.Length : exponentMarker;
        long explicitExponent = 0;
        if (exponentMarker >= 0 &&
            !long.TryParse(
                text.AsSpan(exponentMarker + 1),
                NumberStyles.AllowLeadingSign,
                CultureInfo.InvariantCulture,
                out explicitExponent))
        {
            return false;
        }

        string mantissa = text[mantissaStart..mantissaEnd]
            .Replace(",", string.Empty, StringComparison.Ordinal);
        int decimalPoint = mantissa.IndexOf('.');
        int fractionalDigits = decimalPoint < 0 ? 0 : mantissa.Length - decimalPoint - 1;
        string digits = mantissa.Replace(".", string.Empty, StringComparison.Ordinal);
        int firstNonZero = 0;
        while (firstNonZero < digits.Length && digits[firstNonZero] == '0')
        {
            firstNonZero++;
        }

        if (firstNonZero == digits.Length)
        {
            canonical = new CanonicalDecimal(false, "0", 0);
            return true;
        }

        int lastNonZero = digits.Length - 1;
        while (digits[lastNonZero] == '0')
        {
            lastNonZero--;
        }

        int trailingZeros = digits.Length - lastNonZero - 1;
        long exponent;
        try
        {
            exponent = checked(explicitExponent - fractionalDigits + trailingZeros);
        }
        catch (OverflowException)
        {
            return false;
        }

        canonical = new CanonicalDecimal(
            isNegative,
            digits[firstNonZero..(lastNonZero + 1)],
            exponent);
        return true;
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

    private readonly record struct CanonicalDecimal(
        bool IsNegative,
        string SignificantDigits,
        long Exponent);
}
