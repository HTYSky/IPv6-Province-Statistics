using System.Collections.ObjectModel;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Domain.Reporting;

public sealed record FormulaCalculationResult(
    IReadOnlyDictionary<string, decimal> Values,
    IReadOnlyList<ValidationIssue> Issues);

public static class ReportCalculator
{
    private const decimal GbpsToPbPerDay = 3600m * 24m / 8m / 1024m / 1024m;

    private static readonly IReadOnlyDictionary<string, decimal> EmptyValues =
        new ReadOnlyDictionary<string, decimal>(new Dictionary<string, decimal>());

    private static readonly IReadOnlyList<ValidationIssue> EmptyIssues =
        Array.AsReadOnly(Array.Empty<ValidationIssue>());

    public static FormulaCalculationResult Calculate(ProvinceReportInput input)
    {
        ArgumentNullException.ThrowIfNull(input);

        decimal c2 = input[MetricKey.MetroTotal] * GbpsToPbPerDay;
        decimal e2 = input[MetricKey.MetroIpv6] * GbpsToPbPerDay;
        decimal c9 = input[MetricKey.MobileCoreTotal] * GbpsToPbPerDay;
        decimal e9 = input[MetricKey.MobileCoreIpv6] * GbpsToPbPerDay;
        decimal d4 = input[MetricKey.BroadbandTotal];
        decimal e4 = input[MetricKey.BroadbandIpv6];
        decimal d5 = input[MetricKey.InternetTotal] * GbpsToPbPerDay;
        decimal e5 = input[MetricKey.InternetIpv6] * GbpsToPbPerDay;
        decimal d6 = input[MetricKey.InternetOneGTotal] * GbpsToPbPerDay;
        decimal e6 = input[MetricKey.InternetOneGIpv6] * GbpsToPbPerDay;
        decimal d7 = input[MetricKey.IdcTotal] * GbpsToPbPerDay;
        decimal e7 = input[MetricKey.IdcIpv6] * GbpsToPbPerDay;
        decimal d8 = input[MetricKey.IdcTenGTotal] * GbpsToPbPerDay;
        decimal e8 = input[MetricKey.IdcTenGIpv6] * GbpsToPbPerDay;
        decimal d10 = input[MetricKey.HumanTotal];
        decimal e10 = input[MetricKey.HumanIpv6];
        decimal d11 = input[MetricKey.IotTotal];
        decimal e11 = input[MetricKey.IotIpv6];
        decimal c3 = c2 - c9 * 0.9m;
        decimal e3 = e2 - e9 * 0.9m;

        (string Name, decimal Value)[] denominators =
        [
            ("C2", c2),
            ("C3", c3),
            ("D4", d4),
            ("D5", d5),
            ("D6", d6),
            ("D7", d7),
            ("D8", d8),
            ("C9", c9),
            ("D10", d10),
            ("D11", d11),
            ("C3+C9", c3 + c9),
            ("D4+D5+D7", d4 + d5 + d7),
            ("D10+D11", d10 + d11),
        ];

        string[] zeroDenominators = denominators
            .Where(denominator => denominator.Value == 0m)
            .Select(denominator => denominator.Name)
            .ToArray();
        if (zeroDenominators.Length > 0)
        {
            IReadOnlyList<ValidationIssue> issues = Array.AsReadOnly(
            [
                new ValidationIssue(
                    "FORMULA_DIVIDE_BY_ZERO",
                    $"模板公式分母为零：{string.Join(", ", zeroDenominators)}。"),
            ]);
            return new FormulaCalculationResult(EmptyValues, issues);
        }

        decimal c3AndC9 = c3 + c9;
        decimal d4AndD5AndD7 = d4 + d5 + d7;
        decimal d10AndD11 = d10 + d11;
        decimal g9 = c9 / c3AndC9;

        var values = new Dictionary<string, decimal>
        {
            ["C2"] = c2,
            ["E2"] = e2,
            ["C9"] = c9,
            ["E9"] = e9,
            ["D4"] = d4,
            ["E4"] = e4,
            ["D5"] = d5,
            ["E5"] = e5,
            ["D6"] = d6,
            ["E6"] = e6,
            ["D7"] = d7,
            ["E7"] = e7,
            ["D8"] = d8,
            ["E8"] = e8,
            ["D10"] = d10,
            ["E10"] = e10,
            ["D11"] = d11,
            ["E11"] = e11,
            ["C3"] = c3,
            ["E3"] = e3,
            ["F2"] = e2 / c2,
            ["F3"] = e3 / c3,
            ["F4"] = e4 / d4,
            ["F5"] = e5 / d5,
            ["F6"] = e6 / d6,
            ["F7"] = e7 / d7,
            ["F8"] = e8 / d8,
            ["F9"] = e9 / c9,
            ["F10"] = e10 / d10,
            ["F11"] = e11 / d11,
            ["G3"] = c3 / c3AndC9,
            ["G4"] = d4 / d4AndD5AndD7,
            ["G5"] = d5 / d4AndD5AndD7,
            ["G7"] = d7 / d4AndD5AndD7,
            ["G9"] = g9,
            ["G10"] = g9 * (d10 / d10AndD11),
            ["G11"] = g9 * (d11 / d10AndD11),
        };

        return new FormulaCalculationResult(
            new ReadOnlyDictionary<string, decimal>(values),
            EmptyIssues);
    }
}
