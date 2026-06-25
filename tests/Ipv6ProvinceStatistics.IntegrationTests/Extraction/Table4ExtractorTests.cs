using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Extraction;

public sealed class Table4ExtractorTests
{
    [Fact]
    public async Task Reader_maps_human_and_iot_daily_pb_columns()
    {
        var path = TempFiles.Next("4.xlsx");
        TestWorkbookBuilder.Create(path,
            new TestSheet("人网统计", [TestCell.SharedText("G5", "北京"),
                TestCell.Number("R5", "8.25"), TestCell.Number("S5", "5.5")]),
            new TestSheet("物网统计", [TestCell.SharedText("G5", "北京市"),
                TestCell.Number("R5", "1.125"), TestCell.Number("S5", "0.25")]));
        var result = await new OpenXmlSourceWorkbookReader().ReadAsync(
            path, SourceWorkbookKind.Table4, CancellationToken.None);
        var values = result.Values[new Province("北京")];
        Assert.Equal(8.25m, values[MetricKey.HumanTotal]);
        Assert.Equal(5.5m, values[MetricKey.HumanIpv6]);
        Assert.Equal(1.125m, values[MetricKey.IotTotal]);
        Assert.Equal(0.25m, values[MetricKey.IotIpv6]);
        Assert.Empty(result.Issues);
    }

    [Fact]
    public async Task Reader_rejects_duplicate_province_rows()
    {
        var path = TempFiles.Next("4-duplicate.xlsx");
        TestWorkbookBuilder.Create(path,
            new TestSheet("人网统计", [
                TestCell.SharedText("G5", "北京"), TestCell.Number("R5", "8"), TestCell.Number("S5", "5"),
                TestCell.SharedText("G6", "北京市"), TestCell.Number("R6", "9"), TestCell.Number("S6", "6")
            ]),
            new TestSheet("物网统计", [
                TestCell.SharedText("G5", "北京"), TestCell.Number("R5", "1"), TestCell.Number("S5", "0.2")
            ]));
        var result = await new OpenXmlSourceWorkbookReader().ReadAsync(
            path, SourceWorkbookKind.Table4, CancellationToken.None);
        Assert.Contains(result.Issues, issue => issue.Code == "PROVINCE_DUPLICATE"
            && issue.Province == "北京");
        Assert.Empty(result.Values);
    }
}
