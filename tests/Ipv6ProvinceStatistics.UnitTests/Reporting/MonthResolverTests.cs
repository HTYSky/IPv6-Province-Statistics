using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.UnitTests.Reporting;

public sealed class MonthResolverTests
{
    [Fact]
    public void ExtractRecognizesFullChineseYearAndMonth()
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract(
            "file-name",
            "1-2026年5月IPv6关键指标进展情况.xlsx");

        MonthMarker marker = Assert.Single(markers);
        Assert.Equal(new MonthMarker(2026, 5, "file-name"), marker);
    }

    [Fact]
    public void ExtractRecognizesCompactYearAndMonth()
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract(
            "file-name",
            "5-IP网数据采集分析-202605月.xlsx");

        MonthMarker marker = Assert.Single(markers);
        Assert.Equal(new MonthMarker(2026, 5, "file-name"), marker);
    }

    [Fact]
    public void ExtractRecognizesMonthWithoutYear()
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract(
            "file-name",
            "4-统计表格-5月的数据.xlsx");

        MonthMarker marker = Assert.Single(markers);
        Assert.Equal(new MonthMarker(null, 5, "file-name"), marker);
    }

    [Fact]
    public void ExtractRecognizesPartialMonthAfterNonNumericYearWord()
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract("sheet-a", "今年5月");

        Assert.Equal(new MonthMarker(null, 5, "sheet-a"), Assert.Single(markers));
    }

    [Fact]
    public void ExtractIgnoresMalformedUnicodeYearWithoutFallingBackToPartialMonth()
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract(
            "file-name",
            "20２６年5月.xlsx");

        Assert.Empty(markers);
    }

    [Theory]
    [InlineData("２202605")]
    [InlineData("202605２")]
    [InlineData("20２６05月")]
    public void ExtractDoesNotStartOrEndMarkerInsideUnicodeDigitSequence(string text)
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract("file-name", text);

        Assert.Empty(markers);
    }

    [Fact]
    public void ExtractReturnsAReadOnlySnapshot()
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract("file-a", "2026年5月.xlsx");
        IList<MonthMarker> listView = Assert.IsAssignableFrom<IList<MonthMarker>>(markers);

        Assert.Throws<NotSupportedException>(() =>
            listView[0] = new MonthMarker(2026, 4, "changed"));
    }

    [Fact]
    public void ExtractIgnoresUnsupportedExplicitYearWithoutPartialFallback()
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract("file-a", "2100年5月.xlsx");

        Assert.Empty(markers);
    }

    [Fact]
    public void ExtractIgnoresExplicitInvalidMonthWithoutPartialFallback()
    {
        IReadOnlyList<MonthMarker> markers = MonthTextParser.Extract("file-a", "2026年13月.xlsx");

        Assert.Empty(markers);
    }

    [Theory]
    [InlineData(null, 0)]
    [InlineData(2026, 13)]
    [InlineData(2100, 5)]
    public void ResolveRejectsInvalidHandConstructedMarker(int? year, int month)
    {
        MonthResolution resolution = MonthResolver.Resolve(
            [new MonthMarker(year, month, "source-a")],
            new ReportMonth(2026, 5));

        Assert.False(resolution.IsValid);
        Assert.Null(resolution.Month);
        ValidationIssue issue = Assert.Single(resolution.Issues);
        Assert.Equal("MONTH_INVALID", issue.Code);
        Assert.Contains("source-a", issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ResolveRejectsDefaultReportMonthSelection()
    {
        MonthResolution resolution = MonthResolver.Resolve([], default(ReportMonth));

        Assert.False(resolution.IsValid);
        Assert.Null(resolution.Month);
        Assert.Equal("MONTH_INVALID", Assert.Single(resolution.Issues).Code);
    }

    [Fact]
    public void ResolveInvalidSelectionMessageIncludesAvailableMarkerSource()
    {
        MonthResolution resolution = MonthResolver.Resolve(
            [new MonthMarker(2026, 5, "source-a")],
            default(ReportMonth));

        ValidationIssue issue = Assert.Single(resolution.Issues);
        Assert.Equal("MONTH_INVALID", issue.Code);
        Assert.Contains("source-a", issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ResolveCombinesMatchingFullAndPartialMarkers()
    {
        MonthMarker[] markers =
        [
            new(2026, 5, "file-a"),
            new(null, 5, "sheet-a"),
        ];

        MonthResolution resolution = MonthResolver.Resolve(markers, selected: null);

        Assert.True(resolution.IsValid);
        Assert.Equal(new ReportMonth(2026, 5), resolution.Month);
        Assert.Empty(resolution.Issues);
    }

    [Fact]
    public void ResolveRejectsConflictingFullMarkersEvenWhenSelectedMatchesOne()
    {
        MonthMarker[] markers =
        [
            new(2026, 5, "file-a"),
            new(2026, 4, "file-b"),
        ];

        MonthResolution resolution = MonthResolver.Resolve(markers, new ReportMonth(2026, 5));

        Assert.False(resolution.IsValid);
        Assert.Null(resolution.Month);
        ValidationIssue issue = Assert.Single(resolution.Issues);
        Assert.Equal("MONTH_CONFLICT", issue.Code);
        Assert.Contains("file-a", issue.Message, StringComparison.Ordinal);
        Assert.Contains("file-b", issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ResolveUsesSelectedYearForMatchingPartialMarker()
    {
        MonthResolution resolution = MonthResolver.Resolve(
            [new MonthMarker(null, 5, "sheet-a")],
            new ReportMonth(2026, 5));

        Assert.True(resolution.IsValid);
        Assert.Equal(new ReportMonth(2026, 5), resolution.Month);
        Assert.Empty(resolution.Issues);
    }

    [Fact]
    public void ResolveRejectsConflictingPartialMarkers()
    {
        MonthResolution resolution = MonthResolver.Resolve(
            [new MonthMarker(null, 5, "sheet-a"), new MonthMarker(null, 4, "sheet-b")],
            new ReportMonth(2026, 5));

        Assert.False(resolution.IsValid);
        Assert.Null(resolution.Month);
        Assert.Equal("MONTH_CONFLICT", Assert.Single(resolution.Issues).Code);
    }

    [Fact]
    public void ResolveRejectsSelectedMonthThatConflictsWithFullMarker()
    {
        MonthResolution resolution = MonthResolver.Resolve(
            [new MonthMarker(2026, 5, "file-a")],
            new ReportMonth(2026, 4));

        Assert.False(resolution.IsValid);
        Assert.Null(resolution.Month);
        Assert.Equal("MONTH_CONFLICT", Assert.Single(resolution.Issues).Code);
    }

    [Fact]
    public void ResolveRequiresYearWhenOnlyPartialMarkersExist()
    {
        MonthResolution resolution = MonthResolver.Resolve(
            [new MonthMarker(null, 5, "sheet-a")],
            selected: null);

        Assert.False(resolution.IsValid);
        Assert.Null(resolution.Month);
        Assert.Equal("MONTH_YEAR_MISSING", Assert.Single(resolution.Issues).Code);
    }

    [Fact]
    public void ResolveUsesSelectedMonthWhenNoMarkersExist()
    {
        MonthResolution resolution = MonthResolver.Resolve([], new ReportMonth(2026, 5));

        Assert.True(resolution.IsValid);
        Assert.Equal(new ReportMonth(2026, 5), resolution.Month);
        Assert.Empty(resolution.Issues);
    }

    [Fact]
    public void ResolveReportsMissingMonthWhenNoMarkersOrSelectionExist()
    {
        MonthResolution resolution = MonthResolver.Resolve([], selected: null);

        Assert.False(resolution.IsValid);
        Assert.Null(resolution.Month);
        Assert.Equal("MONTH_MISSING", Assert.Single(resolution.Issues).Code);
    }

    [Fact]
    public void ReportMonthFormatsFolderAndFileNamesWithZeroPaddedMonth()
    {
        var month = new ReportMonth(2026, 5);

        Assert.Equal("2026年05月统计结果", month.FolderName);
        Assert.Equal("2026年05月", month.FileSuffix);
    }

    [Theory]
    [InlineData(1999, 5)]
    [InlineData(10000, 5)]
    [InlineData(2026, 0)]
    [InlineData(2026, 13)]
    public void ReportMonthRejectsOutOfRangeValues(int year, int month)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ReportMonth(year, month));
    }
}
