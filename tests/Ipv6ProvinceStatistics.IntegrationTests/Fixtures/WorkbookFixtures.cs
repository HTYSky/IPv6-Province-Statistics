namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

internal static class WorkbookFixtures
{
    public static string CreateTable1(string fileName)
    {
        string path = TempFiles.Next(fileName);
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("A1", "2026年5月统计表"),
                    TestCell.SharedText("B2", "省份"),
                    TestCell.SharedText("C2", "运营商"),
                    TestCell.SharedText("D3", "城域网出口IPv4和IPv6总流量"),
                    TestCell.SharedText("E3", "城域网出口IPv6总流量"),
                    TestCell.SharedText("G3", "移动核心网出口IPv4和IPv6总流量"),
                    TestCell.SharedText("H3", "移动核心网出口IPv6总流量"),
                ]));
        return path;
    }

    public static string CreateTable4(string fileName)
    {
        string path = TempFiles.Next(fileName);
        TestCell[] headers =
        [
            TestCell.SharedText("H4", " prov\u2003_id "),
            TestCell.SharedText("R4", "总计日均流量（ PB ）"),
            TestCell.SharedText("S4", "ipv6日均流量 ( pb )"),
        ];
        TestWorkbookBuilder.Create(
            path,
            new TestSheet("人网统计", headers),
            new TestSheet("物网统计", headers));
        return path;
    }

    public static string CreateTable5(string idcName, string fileName)
    {
        string path = TempFiles.Next(fileName);
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "互联网专线汇总",
                [
                    TestCell.SharedText("E2", "总流量"),
                    TestCell.SharedText("F2", "IPv6流量"),
                    TestCell.SharedText("K2", "总流量"),
                    TestCell.SharedText("L2", "IPv6流量"),
                    TestCell.SharedText("M2", "省"),
                ]),
            new TestSheet(
                idcName,
                [
                    TestCell.SharedText("F2", "总流量"),
                    TestCell.SharedText("G2", "V6日流量"),
                    TestCell.SharedText("M2", "总流量"),
                    TestCell.SharedText("N2", "V6日流量"),
                    TestCell.SharedText("O2", "省"),
                ]));
        return path;
    }

    public static string CreateTable8(string sheetName, string fileName)
    {
        string path = TempFiles.Next(fileName);
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                sheetName,
                [
                    TestCell.SharedText("A2", "月"),
                    TestCell.SharedText("B2", "省份"),
                    TestCell.SharedText("D2", "总流量"),
                    TestCell.SharedText("J2", "IPv6总流量"),
                    TestCell.Number("A4", "202605"),
                ]));
        return path;
    }

    public static string CreateAmbiguous(string fileName = "ambiguous.xlsx")
    {
        string path = TempFiles.Next(fileName);
        TestCell[] table4Headers =
        [
            TestCell.SharedText("H4", "prov_id"),
            TestCell.SharedText("R4", "总计日均流量(PB)"),
            TestCell.SharedText("S4", "IPV6日均流量(PB)"),
        ];
        TestWorkbookBuilder.Create(
            path,
            new TestSheet(
                "分省统计表",
                [
                    TestCell.SharedText("B2", "省份"),
                    TestCell.SharedText("C2", "运营商"),
                    TestCell.SharedText("D3", "城域网出口IPv4和IPv6总流量"),
                    TestCell.SharedText("E3", "城域网出口IPv6总流量"),
                    TestCell.SharedText("G3", "移动核心网出口IPv4和IPv6总流量"),
                    TestCell.SharedText("H3", "移动核心网出口IPv6总流量"),
                ]),
            new TestSheet("人网统计", table4Headers),
            new TestSheet("物网统计", table4Headers));
        return path;
    }
}
