using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Application.Abstractions;

public interface ITemplateReportExporter
{
    Task<IReadOnlyList<ValidationIssue>> ValidateTemplateAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<ValidationIssue>> ExportAsync(Province province, ReportMonth month,
        ProvinceReportInput input, string outputPath, CancellationToken cancellationToken);
}
