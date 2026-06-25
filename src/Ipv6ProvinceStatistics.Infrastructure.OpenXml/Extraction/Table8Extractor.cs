using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table8Extractor
{
    public static SourceReadResult Extract(
        OpenXmlWorkbookReader reader,
        string path,
        CancellationToken cancellationToken)
    {
        string sheet = FindProvinceSummarySheet(reader);
        var issues = new List<ValidationIssue>();
        var values = new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>();

        foreach (uint row in reader.GetPopulatedRows(sheet).Where(row => row >= 4).Order())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ProvinceCatalog.TryResolve(
                    reader.GetText(sheet, $"B{row}"),
                    out Province province)) continue;

            bool totalValid = ExtractorSupport.TryReadDecimal(
                reader, path, sheet, $"D{row}", province, issues, out decimal total);
            bool ipv6Valid = ExtractorSupport.TryReadDecimal(
                reader, path, sheet, $"J{row}", province, issues, out decimal ipv6);

            if (!(totalValid & ipv6Valid)) continue;

            var metrics = new ReadOnlyDictionary<MetricKey, decimal>(
                new Dictionary<MetricKey, decimal>
                {
                    [MetricKey.BroadbandTotal] = total,
                    [MetricKey.BroadbandIpv6] = ipv6,
                });
            ExtractorSupport.AddProvince(values, path, sheet, province, metrics, issues);
        }

        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyDictionary<Province, IReadOnlyDictionary<MetricKey, decimal>> resultValues =
            issues.Count == 0
                ? ExtractorSupport.AsReadOnly(values)
                : ExtractorSupport.AsReadOnly(new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>());
        return new SourceReadResult(SourceWorkbookKind.Table8, resultValues, issues.AsReadOnly());
    }

    private static string FindProvinceSummarySheet(OpenXmlWorkbookReader reader)
    {
        foreach (string name in reader.SheetNames)
            if (name == "省统计" || name.EndsWith("-省统计", StringComparison.Ordinal))
                return name;
        throw new InvalidOperationException("8 表中未找到省统计工作表。");
    }
}
