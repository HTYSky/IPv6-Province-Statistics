using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.UnitTests.Reporting;

public sealed class ReportCalculatorTests
{
    private const decimal GbpsToPbPerDay = 3600m * 24m / 8m / 1024m / 1024m;

    private static readonly string[] ExpectedAddresses =
    [
        "C2", "E2", "C9", "E9",
        "D4", "E4", "D5", "E5", "D6", "E6", "D7", "E7", "D8", "E8",
        "D10", "E10", "D11", "E11", "C3", "E3",
        "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11",
        "G3", "G4", "G5", "G7", "G9", "G10", "G11",
    ];

    [Fact]
    public void CalculateThrowsWhenInputIsNull()
    {
        Assert.Throws<ArgumentNullException>(() => ReportCalculator.Calculate(null!));
    }

    [Fact]
    public void CalculateWithEqualMetricsReturnsTheExpectedRepresentativeValues()
    {
        ProvinceReportInput input = CreateInput(_ => 100m);

        FormulaCalculationResult result = ReportCalculator.Calculate(input);

        Assert.Empty(result.Issues);
        Assert.Equal(37, result.Values.Count);
        Assert.Equal(100m * GbpsToPbPerDay, result.Values["C2"]);
        Assert.Equal(result.Values["E2"] / result.Values["C2"], result.Values["F2"]);
        Assert.Equal(result.Values["G9"] * 0.5m, result.Values["G10"]);
    }

    [Fact]
    public void CalculateReproducesAllTemplateAddressesAndFormulasExactly()
    {
        ProvinceReportInput input = CreateInput(metric => metric switch
        {
            MetricKey.MetroTotal => 200m,
            MetricKey.MetroIpv6 => 120m,
            MetricKey.MobileCoreTotal => 50m,
            MetricKey.MobileCoreIpv6 => 20m,
            MetricKey.InternetTotal => 300m,
            MetricKey.InternetIpv6 => 150m,
            MetricKey.InternetOneGTotal => 75m,
            MetricKey.InternetOneGIpv6 => 30m,
            MetricKey.IdcTotal => 200m,
            MetricKey.IdcIpv6 => 80m,
            MetricKey.IdcTenGTotal => 25m,
            MetricKey.IdcTenGIpv6 => 10m,
            MetricKey.HumanTotal => 90m,
            MetricKey.HumanIpv6 => 45m,
            MetricKey.IotTotal => 30m,
            MetricKey.IotIpv6 => 15m,
            MetricKey.BroadbandTotal => 1000m,
            MetricKey.BroadbandIpv6 => 600m,
            _ => throw new ArgumentOutOfRangeException(nameof(metric)),
        });

        FormulaCalculationResult result = ReportCalculator.Calculate(input);
        IReadOnlyDictionary<string, decimal> expected = CalculateExpectedValues(input);

        Assert.Empty(result.Issues);
        Assert.Equal(
            ExpectedAddresses.Order(StringComparer.Ordinal),
            result.Values.Keys.Order(StringComparer.Ordinal));
        Assert.Equal(expected.Count, result.Values.Count);
        foreach ((string address, decimal value) in expected)
        {
            Assert.Equal(value, result.Values[address]);
        }
    }

    [Fact]
    public void CalculateWithAllZeroMetricsReportsDivideByZeroWithoutPartialValues()
    {
        var result = ReportCalculator.Calculate(CreateInput(_ => 0m));
        Assert.Equal(37, result.Values.Count);
        Assert.Equal(0m, result.Values["F2"]);
        Assert.Equal(0m, result.Values["F3"]);
    }

    [Fact]
    public void CalculateReportsNumericOverflowWhenAggregateDenominatorAdditionOverflows()
    {
        Dictionary<MetricKey, decimal> values = CreateMetricValues(_ => 1m);
        values[MetricKey.HumanTotal] = decimal.MaxValue;
        values[MetricKey.IotTotal] = decimal.MaxValue;

        FormulaCalculationResult result = ReportCalculator.Calculate(new ProvinceReportInput(values));

        AssertNumericOverflow(result);
        IList<ValidationIssue> issues =
            Assert.IsAssignableFrom<IList<ValidationIssue>>(result.Issues);
        Assert.Throws<NotSupportedException>(() =>
            issues[0] = new ValidationIssue("X", "X"));
    }

    [Fact]
    public void CalculateReportsNumericOverflowWhenRatioDivisionOverflows()
    {
        Dictionary<MetricKey, decimal> values = CreateMetricValues(_ => 1m);
        values[MetricKey.HumanTotal] = 0.0000000000000000000000000001m;
        values[MetricKey.HumanIpv6] = decimal.MaxValue;

        FormulaCalculationResult result = ReportCalculator.Calculate(new ProvinceReportInput(values));

        AssertNumericOverflow(result);
    }

    [Theory]
    [InlineData("C2")]
    [InlineData("C3")]
    [InlineData("D4")]
    [InlineData("D5")]
    [InlineData("D6")]
    [InlineData("D7")]
    [InlineData("D8")]
    [InlineData("C9")]
    [InlineData("D10")]
    [InlineData("D11")]
    [InlineData("C3+C9")]
    [InlineData("D4+D5+D7")]
    [InlineData("D10+D11")]
    public void CalculateFailsPreflightWhenARequiredDenominatorIsZero(string denominator)
    {
        Dictionary<MetricKey, decimal> values = CreateMetricValues(_ => 1m);
        MakeDenominatorZero(values, denominator);

        FormulaCalculationResult result = ReportCalculator.Calculate(new ProvinceReportInput(values));

        Assert.Equal(37, result.Values.Count);
    }

    [Fact]
    public void CalculateReturnsCollectionsThatCannotBeChangedExternally()
    {
        FormulaCalculationResult success = ReportCalculator.Calculate(CreateInput(_ => 1m));
        IDictionary<string, decimal> values =
            Assert.IsAssignableFrom<IDictionary<string, decimal>>(success.Values);
        IList<ValidationIssue> successIssues =
            Assert.IsAssignableFrom<IList<ValidationIssue>>(success.Issues);

        Assert.Throws<NotSupportedException>(() => values["C2"] = 0m);
        Assert.Throws<NotSupportedException>(() => successIssues.Add(new ValidationIssue("X", "X")));

        FormulaCalculationResult zeroResult = ReportCalculator.Calculate(CreateInput(_ => 0m));
        IList<ValidationIssue> zeroIssues =
            Assert.IsAssignableFrom<IList<ValidationIssue>>(zeroResult.Issues);

        Assert.Throws<NotSupportedException>(() => zeroIssues.Add(new ValidationIssue("X", "X")));
    }

    private static IReadOnlyDictionary<string, decimal> CalculateExpectedValues(
        ProvinceReportInput input)
    {
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

        return new Dictionary<string, decimal>
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
            ["G3"] = c3 / (c3 + c9),
            ["G4"] = d4 / (d4 + d5 + d7),
            ["G5"] = d5 / (d4 + d5 + d7),
            ["G7"] = d7 / (d4 + d5 + d7),
            ["G9"] = c9 / (c3 + c9),
            ["G10"] = c9 / (c3 + c9) * (d10 / (d10 + d11)),
            ["G11"] = c9 / (c3 + c9) * (d11 / (d10 + d11)),
        };
    }

    private static ProvinceReportInput CreateInput(Func<MetricKey, decimal> valueFactory) =>
        new(CreateMetricValues(valueFactory));

    private static void AssertNumericOverflow(FormulaCalculationResult result)
    {
        Assert.Empty(result.Values);
        ValidationIssue issue = Assert.Single(result.Issues);
        Assert.Equal("FORMULA_NUMERIC_OVERFLOW", issue.Code);
        Assert.Contains("数值", issue.Message, StringComparison.Ordinal);
    }

    private static Dictionary<MetricKey, decimal> CreateMetricValues(
        Func<MetricKey, decimal> valueFactory) =>
        Enum.GetValues<MetricKey>().ToDictionary(metric => metric, valueFactory);

    private static void MakeDenominatorZero(
        IDictionary<MetricKey, decimal> values,
        string denominator)
    {
        switch (denominator)
        {
            case "C2":
                values[MetricKey.MetroTotal] = 0m;
                break;
            case "C3":
                values[MetricKey.MetroTotal] = 9m;
                values[MetricKey.MobileCoreTotal] = 10m;
                break;
            case "D4":
                values[MetricKey.BroadbandTotal] = 0m;
                break;
            case "D5":
                values[MetricKey.InternetTotal] = 0m;
                break;
            case "D6":
                values[MetricKey.InternetOneGTotal] = 0m;
                break;
            case "D7":
                values[MetricKey.IdcTotal] = 0m;
                break;
            case "D8":
                values[MetricKey.IdcTenGTotal] = 0m;
                break;
            case "C9":
                values[MetricKey.MobileCoreTotal] = 0m;
                break;
            case "D10":
                values[MetricKey.HumanTotal] = 0m;
                break;
            case "D11":
                values[MetricKey.IotTotal] = 0m;
                break;
            case "C3+C9":
                values[MetricKey.MetroTotal] = -1m;
                values[MetricKey.MobileCoreTotal] = 10m;
                break;
            case "D4+D5+D7":
                values[MetricKey.BroadbandTotal] = GbpsToPbPerDay;
                values[MetricKey.InternetTotal] = 1m;
                values[MetricKey.IdcTotal] = -2m;
                break;
            case "D10+D11":
                values[MetricKey.HumanTotal] = 1m;
                values[MetricKey.IotTotal] = -1m;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(denominator));
        }
    }
}
