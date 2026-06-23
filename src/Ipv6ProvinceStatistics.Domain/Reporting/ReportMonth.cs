namespace Ipv6ProvinceStatistics.Domain.Reporting;

public readonly record struct ReportMonth
{
    public ReportMonth(int year, int month)
    {
        if (year is < 2000 or > 9999)
        {
            throw new ArgumentOutOfRangeException(nameof(year), year, "Year must be between 2000 and 9999.");
        }

        if (month is < 1 or > 12)
        {
            throw new ArgumentOutOfRangeException(nameof(month), month, "Month must be between 1 and 12.");
        }

        Year = year;
        Month = month;
    }

    public int Year { get; }

    public int Month { get; }

    public string FolderName => $"{Year:D4}年{Month:D2}月统计结果";

    public string FileSuffix => $"{Year:D4}年{Month:D2}月";
}
