using Ipv6ProvinceStatistics.Domain.Reporting;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

public static class TemplateCellMap
{
    public static IReadOnlyDictionary<MetricKey, string> Inputs { get; } =
        new Dictionary<MetricKey, string>
        {
            [MetricKey.MetroTotal] = "C17", [MetricKey.MetroIpv6] = "E17",
            [MetricKey.MobileCoreTotal] = "C18", [MetricKey.MobileCoreIpv6] = "E18",
            [MetricKey.InternetTotal] = "C20", [MetricKey.InternetIpv6] = "E20",
            [MetricKey.InternetOneGTotal] = "C21", [MetricKey.InternetOneGIpv6] = "E21",
            [MetricKey.IdcTotal] = "C22", [MetricKey.IdcIpv6] = "E22",
            [MetricKey.IdcTenGTotal] = "C23", [MetricKey.IdcTenGIpv6] = "E23",
            [MetricKey.HumanTotal] = "C25", [MetricKey.HumanIpv6] = "E25",
            [MetricKey.IotTotal] = "C26", [MetricKey.IotIpv6] = "E26",
            [MetricKey.BroadbandTotal] = "C27", [MetricKey.BroadbandIpv6] = "E27"
        };
}
