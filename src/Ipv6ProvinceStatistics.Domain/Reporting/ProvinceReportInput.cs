using System.Collections.ObjectModel;

namespace Ipv6ProvinceStatistics.Domain.Reporting;

public sealed class ProvinceReportInput
{
    public ProvinceReportInput(IReadOnlyDictionary<MetricKey, decimal> values)
    {
        ArgumentNullException.ThrowIfNull(values);

        MetricKey[] missingMetrics = Enum.GetValues<MetricKey>()
            .Where(metric => !values.ContainsKey(metric))
            .ToArray();
        if (missingMetrics.Length > 0)
        {
            throw new ArgumentException(
                $"缺少必需指标：{string.Join(", ", missingMetrics)}。",
                nameof(values));
        }

        Values = new ReadOnlyDictionary<MetricKey, decimal>(
            new Dictionary<MetricKey, decimal>(values));
    }

    public decimal this[MetricKey key] => Values[key];

    public IReadOnlyDictionary<MetricKey, decimal> Values { get; }
}
