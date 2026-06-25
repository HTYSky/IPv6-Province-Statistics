using System.Globalization;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;

namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

public sealed record MonthlyFixture(IReadOnlyList<string> Paths,
    IReadOnlyDictionary<Province, ProvinceReportInput> Expected);

public static class MonthlyFixtureBuilder
{
    public static MonthlyFixture Create(string root, bool invalidIdcN)
    {
        Directory.CreateDirectory(root);
        var table1 = new List<TestCell> { TestCell.SharedText("A1", "2026年5月IPv6关键指标进展情况"),
            TestCell.SharedText("B2", "省份"), TestCell.SharedText("C2", "运营商"),
            TestCell.SharedText("D3", "城域网出口IPv4和IPv6总流量（Gbps）"),
            TestCell.SharedText("E3", "城域网出口IPv6总流量（Gbps）"),
            TestCell.SharedText("G3", "移动核心网出口IPv4和IPv6总流量（Gbps）"),
            TestCell.SharedText("H3", "移动核心网出口IPv6总流量（Gbps）") };
        var human = TrafficHeaders(); var iot = TrafficHeaders();
        var internet = new List<TestCell> { TestCell.SharedText("E2", "总流量（Gbps）"), TestCell.SharedText("F2", "IPv6流量（Gbps）"),
            TestCell.SharedText("K2", "总流量（Gbps）"), TestCell.SharedText("L2", "IPv6流量（Gbps）"), TestCell.SharedText("M2", "省") };
        var idc = new List<TestCell> { TestCell.SharedText("F2", "总流量（Gbps）"), TestCell.SharedText("G2", "V6日流量（Gbps）"),
            TestCell.SharedText("M2", "总流量（Gbps）"), TestCell.SharedText("N2", "V6日流量（Gbps）"), TestCell.SharedText("O2", "省") };
        var table8 = new List<TestCell> { TestCell.SharedText("A2", "月"), TestCell.SharedText("B2", "省份"),
            TestCell.SharedText("D2", "总流量(v4+v6，GB)"), TestCell.SharedText("J2", "IPv6总流量（GB）") };
        var expected = new Dictionary<Province, ProvinceReportInput>();
        for (var idx = 0; idx < ProvinceCatalog.All.Count; idx++)
        {
            var p = ProvinceCatalog.All[idx]; var v = 100m + idx;
            var r1 = 4 + idx * 4; var un = r1 + 2; var r = 5 + idx; var r5 = 3 + idx; var r8 = 4 + idx;
            table1.Add(TestCell.SharedText($"B{r1}", p.Name)); table1.Add(TestCell.SharedText($"C{un}", "中国联通"));
            table1.Add(TestCell.Number($"D{un}", F(v+1))); table1.Add(TestCell.Number($"E{un}", F(v+2)));
            table1.Add(TestCell.Number($"G{un}", F(v+3))); table1.Add(TestCell.Number($"H{un}", F(v+4)));
            AddTraffic(human, r, p.Name, v+5, v+6); AddTraffic(iot, r, p.Name, v+7, v+8);
            internet.Add(TestCell.SharedText($"M{r5}", p.Name)); internet.Add(TestCell.Number($"E{r5}", F(v+9)));
            internet.Add(TestCell.Number($"F{r5}", F(v+10))); internet.Add(TestCell.Number($"K{r5}", F(v+11)));
            internet.Add(TestCell.Number($"L{r5}", F(v+12)));
            idc.Add(TestCell.SharedText($"O{r5}", p.Name)); idc.Add(TestCell.Number($"F{r5}", F(v+13)));
            idc.Add(TestCell.Number($"G{r5}", F(v+14))); idc.Add(TestCell.Number($"M{r5}", F(v+15)));
            idc.Add(TestCell.Number($"N{r5}", invalidIdcN && idx == 0 ? "NULL" : F(v+16)));
            table8.Add(TestCell.Number($"A{r8}", "202605")); table8.Add(TestCell.SharedText($"B{r8}", p.Name));
            table8.Add(TestCell.Number($"D{r8}", F(v+17))); table8.Add(TestCell.Number($"J{r8}", F(v+18)));
            expected[p] = new ProvinceReportInput(new Dictionary<MetricKey, decimal> {
                [MetricKey.MetroTotal]=v+1,[MetricKey.MetroIpv6]=v+2,[MetricKey.MobileCoreTotal]=v+3,[MetricKey.MobileCoreIpv6]=v+4,
                [MetricKey.HumanTotal]=v+5,[MetricKey.HumanIpv6]=v+6,[MetricKey.IotTotal]=v+7,[MetricKey.IotIpv6]=v+8,
                [MetricKey.InternetTotal]=v+9,[MetricKey.InternetIpv6]=v+10,[MetricKey.InternetOneGTotal]=v+11,[MetricKey.InternetOneGIpv6]=v+12,
                [MetricKey.IdcTotal]=v+13,[MetricKey.IdcIpv6]=v+14,[MetricKey.IdcTenGTotal]=v+15,[MetricKey.IdcTenGIpv6]=v+16,
                [MetricKey.BroadbandTotal]=v+17,[MetricKey.BroadbandIpv6]=v+18});
        }
        var paths = new[] { Path.Combine(root,"1-2026年5月.xlsx"),Path.Combine(root,"4-5月.xlsx"),
            Path.Combine(root,"5-202605.xlsx"),Path.Combine(root,"8-202605.xlsx") };
        TestWorkbookBuilder.Create(paths[0], new TestSheet("分省统计表", table1));
        TestWorkbookBuilder.Create(paths[1], new TestSheet("人网统计", human), new TestSheet("物网统计", iot));
        TestWorkbookBuilder.Create(paths[2], new TestSheet("互联网专线汇总", internet), new TestSheet("IDC汇总 (客户)", idc));
        TestWorkbookBuilder.Create(paths[3], new TestSheet("1-省统计", table8));
        return new(paths, expected);
    }
    private static List<TestCell> TrafficHeaders() => [
        TestCell.SharedText("H4","prov_id"), TestCell.SharedText("R4","总计日均流量(PB)"), TestCell.SharedText("S4","IPV6日均流量(PB)") ];
    private static void AddTraffic(List<TestCell> cells, int row, string prov, decimal total, decimal ipv6) {
        cells.Add(TestCell.SharedText($"G{row}", prov)); cells.Add(TestCell.Number($"R{row}", F(total))); cells.Add(TestCell.Number($"S{row}", F(ipv6))); }
    private static string F(decimal value) => value.ToString(CultureInfo.InvariantCulture);
}
