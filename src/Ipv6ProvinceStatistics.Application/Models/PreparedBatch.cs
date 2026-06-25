using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record PreparedBatch(TaskWorkspace Workspace, ReportMonth Month,
    IReadOnlyDictionary<Province, ProvinceReportInput> Reports,
    IReadOnlyDictionary<SourceWorkbookKind, SourceFileSnapshot> IdentifiedSources);
