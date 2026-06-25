using System.Collections.ObjectModel;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Application.Services;

public sealed class ProvinceDataAssembler
{
    private static readonly IReadOnlyDictionary<SourceWorkbookKind, IReadOnlyList<MetricKey>>
        MetricsByKind = new ReadOnlyDictionary<SourceWorkbookKind, IReadOnlyList<MetricKey>>(
            new Dictionary<SourceWorkbookKind, IReadOnlyList<MetricKey>>
            {
                [SourceWorkbookKind.Table1] = Array.AsReadOnly<MetricKey>(
                [
                    MetricKey.MetroTotal,
                    MetricKey.MetroIpv6,
                    MetricKey.MobileCoreTotal,
                    MetricKey.MobileCoreIpv6,
                ]),
                [SourceWorkbookKind.Table4] = Array.AsReadOnly<MetricKey>(
                [
                    MetricKey.HumanTotal,
                    MetricKey.HumanIpv6,
                    MetricKey.IotTotal,
                    MetricKey.IotIpv6,
                ]),
                [SourceWorkbookKind.Table5] = Array.AsReadOnly<MetricKey>(
                [
                    MetricKey.InternetTotal,
                    MetricKey.InternetIpv6,
                    MetricKey.InternetOneGTotal,
                    MetricKey.InternetOneGIpv6,
                    MetricKey.IdcTotal,
                    MetricKey.IdcIpv6,
                    MetricKey.IdcTenGTotal,
                    MetricKey.IdcTenGIpv6,
                ]),
                [SourceWorkbookKind.Table8] = Array.AsReadOnly<MetricKey>(
                [
                    MetricKey.BroadbandTotal,
                    MetricKey.BroadbandIpv6,
                ]),
            });

    private static readonly IReadOnlyDictionary<Province, ProvinceReportInput> EmptyReports =
        new ReadOnlyDictionary<Province, ProvinceReportInput>(
            new Dictionary<Province, ProvinceReportInput>());

    public static IReadOnlyDictionary<SourceWorkbookKind, IReadOnlyList<MetricKey>>
        ExpectedMetricsByKind => MetricsByKind;

    public ProvinceAssemblyResult Assemble(IReadOnlyCollection<SourceReadResult> sources)
    {
        ArgumentNullException.ThrowIfNull(sources);

        var issues = sources.SelectMany(source => source.Issues).ToList();
        Dictionary<SourceWorkbookKind, SourceReadResult> sourcesByKind =
            ValidateSourceCounts(sources, issues);
        var valuesByProvince = ProvinceCatalog.All.ToDictionary(
            province => province,
            _ => new Dictionary<MetricKey, decimal>());

        foreach ((SourceWorkbookKind kind, SourceReadResult source) in sourcesByKind)
        {
            IReadOnlyList<MetricKey> expectedMetrics = MetricsByKind[kind];

            foreach (Province province in ProvinceCatalog.All)
            {
                if (!source.Values.TryGetValue(province, out IReadOnlyDictionary<MetricKey, decimal>? values))
                {
                    issues.Add(new ValidationIssue(
                        "PROVINCE_MISSING",
                        $"{kind} 缺少省份 {province.Name} 的数据。",
                        Province: province.Name));
                    continue;
                }

                foreach (MetricKey metric in expectedMetrics)
                {
                    if (!values.TryGetValue(metric, out decimal value))
                    {
                        issues.Add(new ValidationIssue(
                            "METRIC_MISSING",
                            $"{kind} 的 {province.Name} 缺少指标 {metric}。",
                            Province: province.Name));
                        continue;
                    }

                    valuesByProvince[province].Add(metric, value);
                }
            }
        }

        if (issues.Count > 0)
        {
            var partialReports = valuesByProvince
                .Where(kvp => kvp.Value.Count == Enum.GetValues<MetricKey>().Length)
                .ToDictionary(kvp => kvp.Key, kvp => new ProvinceReportInput(kvp.Value));
            return new ProvinceAssemblyResult(
                new ReadOnlyDictionary<Province, ProvinceReportInput>(partialReports),
                issues.AsReadOnly());
        }

        var reports = valuesByProvince.ToDictionary(
            entry => entry.Key,
            entry => new ProvinceReportInput(entry.Value));
        return new ProvinceAssemblyResult(
            new ReadOnlyDictionary<Province, ProvinceReportInput>(reports),
            issues.AsReadOnly());
    }

    private static Dictionary<SourceWorkbookKind, SourceReadResult> ValidateSourceCounts(
        IReadOnlyCollection<SourceReadResult> sources,
        ICollection<ValidationIssue> issues)
    {
        Dictionary<SourceWorkbookKind, SourceReadResult[]> groupedSources = sources
            .GroupBy(source => source.Kind)
            .ToDictionary(group => group.Key, group => group.ToArray());
        var validSources = new Dictionary<SourceWorkbookKind, SourceReadResult>();

        foreach (SourceWorkbookKind kind in MetricsByKind.Keys)
        {
            groupedSources.TryGetValue(kind, out SourceReadResult[]? matchingSources);
            int count = matchingSources?.Length ?? 0;
            if (count != 1)
            {
                issues.Add(new ValidationIssue(
                    "SOURCE_KIND_COUNT",
                    $"来源类型 {kind} 必须且仅出现一次，实际出现 {count} 次。"));
                continue;
            }

            validSources.Add(kind, matchingSources![0]);
        }

        foreach (SourceWorkbookKind kind in groupedSources.Keys.Except(MetricsByKind.Keys))
        {
            issues.Add(new ValidationIssue(
                "SOURCE_KIND_COUNT",
                $"来源类型 {kind} 不受支持，实际出现 {groupedSources[kind].Length} 次。"));
        }

        return validSources;
    }
}
