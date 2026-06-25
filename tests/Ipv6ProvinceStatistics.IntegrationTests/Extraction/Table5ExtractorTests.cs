using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Extraction;

public sealed class Table5ExtractorTests
{
    [Theory]
    [InlineData("IDC汇总 (客户)")]
    [InlineData("IDC汇总（客户）")]
    public async Task Reader_maps_internet_and_idc_columns_and_ignores_cloud_company(string idcSheet)
    {
        var path = TempFiles.Next("5.xlsx");
        TestWorkbookBuilder.Create(path,
            new TestSheet("互联网专线汇总", [
                TestCell.SharedText("M6", "北京市"),
                TestCell.Number("E6", "100.1"), TestCell.Number("F6", "40.2"),
                TestCell.Number("K6", "60.3"), TestCell.Number("L6", "20.4")
            ]),
            new TestSheet(idcSheet, [
                TestCell.SharedText("O4", "北京市"),
                TestCell.Number("F4", "200.5"), TestCell.Number("G4", "80.6"),
                TestCell.Number("M4", "150.7"), TestCell.Number("N4", "70.8"),
                TestCell.SharedText("O11", "云公司"),
                TestCell.Number("F11", "999"), TestCell.Number("G11", "888"),
                TestCell.Number("M11", "777"), TestCell.Number("N11", "666")
            ]));
        var result = await new OpenXmlSourceWorkbookReader().ReadAsync(
            path, SourceWorkbookKind.Table5, CancellationToken.None);
        var values = result.Values[new Province("北京")];
        Assert.Equal(100.1m, values[MetricKey.InternetTotal]);
        Assert.Equal(40.2m, values[MetricKey.InternetIpv6]);
        Assert.Equal(60.3m, values[MetricKey.InternetOneGTotal]);
        Assert.Equal(20.4m, values[MetricKey.InternetOneGIpv6]);
        Assert.Equal(200.5m, values[MetricKey.IdcTotal]);
        Assert.Equal(80.6m, values[MetricKey.IdcIpv6]);
        Assert.Equal(150.7m, values[MetricKey.IdcTenGTotal]);
        Assert.Equal(70.8m, values[MetricKey.IdcTenGIpv6]);
        Assert.Empty(result.Issues);
        Assert.Single(result.Values);
    }
}
