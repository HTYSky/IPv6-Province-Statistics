using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

public sealed class OpenXmlSourceWorkbookReader : ISourceWorkbookReader
{
    public Task<SourceReadResult> ReadAsync(
        string path,
        SourceWorkbookKind kind,
        CancellationToken cancellationToken) =>
    Task.Run(() => ReadCore(path, kind, cancellationToken), cancellationToken);

    private static SourceReadResult ReadCore(
        string path,
        SourceWorkbookKind kind,
        CancellationToken cancellationToken)
    {
        string fileName = Path.GetFileName(path);
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            using OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false);
            cancellationToken.ThrowIfCancellationRequested();

            return kind switch
            {
                SourceWorkbookKind.Table1 => Table1Extractor.Extract(reader, path, cancellationToken),
                _ => throw new NotSupportedException($"Extractor not implemented: {kind}"),
            };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (NotSupportedException)
        {
            throw;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or
            InvalidDataException or OpenXmlPackageException or XmlException or
            KeyNotFoundException or FormatException)
        {
            return new SourceReadResult(
                kind,
                ExtractorSupport.AsReadOnly(
                    new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>()),
                [new ValidationIssue(
                    "SOURCE_UNREADABLE",
                    $"无法读取源工作簿：{fileName}",
                    fileName)]);
        }
    }
}
