using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

public sealed class OpenXmlWorkbookInspector : IWorkbookInspector
{
    private const string Table1Sheet = "分省统计表";
    private const string Table4PeopleSheet = "人网统计";
    private const string Table4ThingsSheet = "物网统计";
    private const string Table5InternetSheet = "互联网专线汇总";
    private const string Table5IdcSheet = "IDC汇总(客户)";

    public Task<WorkbookInspection> InspectAsync(
        string path,
        CancellationToken cancellationToken) =>
        Task.Run(() => InspectCore(path, cancellationToken), cancellationToken);

    private static WorkbookInspection InspectCore(
        string path,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string fileName = Path.GetFileName(path);

        try
        {
            using OpenXmlWorkbookReader reader = OpenXmlWorkbookReader.Open(path, editable: false);
            cancellationToken.ThrowIfCancellationRequested();

            var kinds = new List<SourceWorkbookKind>();
            var markers = new List<MonthMarker>(MonthTextParser.Extract(fileName, fileName));

            bool isTable1 = MatchesTable1(reader);
            if (isTable1)
            {
                kinds.Add(SourceWorkbookKind.Table1);
                markers.AddRange(
                    MonthTextParser.Extract("1表标题", reader.GetText(Table1Sheet, "A1")));
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (MatchesTable4(reader))
            {
                kinds.Add(SourceWorkbookKind.Table4);
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (MatchesTable5(reader))
            {
                kinds.Add(SourceWorkbookKind.Table5);
            }

            cancellationToken.ThrowIfCancellationRequested();
            string[] table8Sheets = reader.SheetNames
                .Where(IsTable8SheetName)
                .Where(sheetName => MatchesTable8Sheet(reader, sheetName))
                .ToArray();
            if (table8Sheets.Length > 0)
            {
                kinds.Add(SourceWorkbookKind.Table8);
                AddTable8MonthMarkers(reader, table8Sheets, markers, cancellationToken);
            }

            cancellationToken.ThrowIfCancellationRequested();
            ValidationIssue[] issues = CreateStructureIssues(fileName, kinds.Count);
            return CreateInspection(path, kinds, markers, issues);
        }
        catch (IOException)
        {
            return CreateUnreadableInspection(path, fileName);
        }
        catch (UnauthorizedAccessException)
        {
            return CreateUnreadableInspection(path, fileName);
        }
        catch (OpenXmlPackageException)
        {
            return CreateUnreadableInspection(path, fileName);
        }
        catch (XmlException)
        {
            return CreateUnreadableInspection(path, fileName);
        }
        catch (FormatException)
        {
            return CreateUnreadableInspection(path, fileName);
        }
    }

    private static bool MatchesTable1(OpenXmlWorkbookReader reader) =>
        reader.HasSheet(Table1Sheet) &&
        HeaderContains(reader, Table1Sheet, "B2", "省份") &&
        HeaderContains(reader, Table1Sheet, "C2", "运营商") &&
        HeaderContains(reader, Table1Sheet, "D3", "城域网出口IPv4和IPv6总流量") &&
        HeaderContains(reader, Table1Sheet, "E3", "城域网出口IPv6总流量") &&
        HeaderContains(reader, Table1Sheet, "G3", "移动核心网出口IPv4和IPv6总流量") &&
        HeaderContains(reader, Table1Sheet, "H3", "移动核心网出口IPv6总流量");

    private static bool MatchesTable4(OpenXmlWorkbookReader reader) =>
        MatchesTable4Sheet(reader, Table4PeopleSheet) &&
        MatchesTable4Sheet(reader, Table4ThingsSheet);

    private static bool MatchesTable4Sheet(OpenXmlWorkbookReader reader, string sheetName) =>
        reader.HasSheet(sheetName) &&
        HeaderContains(reader, sheetName, "H4", "prov_id") &&
        HeaderContains(reader, sheetName, "R4", "总计日均流量(PB)") &&
        HeaderContains(reader, sheetName, "S4", "IPV6日均流量(PB)");

    private static bool MatchesTable5(OpenXmlWorkbookReader reader)
    {
        if (!reader.HasSheet(Table5InternetSheet) ||
            !HeaderContains(reader, Table5InternetSheet, "M2", "省") ||
            !IsOrdinaryTotalHeader(reader, Table5InternetSheet, "E2") ||
            !HeaderContains(reader, Table5InternetSheet, "F2", "IPv6流量") ||
            !IsOrdinaryTotalHeader(reader, Table5InternetSheet, "K2") ||
            !HeaderContains(reader, Table5InternetSheet, "L2", "IPv6流量"))
        {
            return false;
        }

        return reader.SheetNames
            .Where(IsTable5IdcSheetName)
            .Any(sheetName => MatchesTable5IdcSheet(reader, sheetName));
    }

    private static bool IsTable5IdcSheetName(string sheetName) =>
        string.Equals(
            HeaderText.Normalize(sheetName),
            Table5IdcSheet,
            StringComparison.OrdinalIgnoreCase);

    private static bool MatchesTable5IdcSheet(
        OpenXmlWorkbookReader reader,
        string sheetName) =>
        HeaderContains(reader, sheetName, "O2", "省") &&
        IsOrdinaryTotalHeader(reader, sheetName, "F2") &&
        HeaderContains(reader, sheetName, "G2", "V6日流量") &&
        IsOrdinaryTotalHeader(reader, sheetName, "M2") &&
        HeaderContains(reader, sheetName, "N2", "V6日流量");

    private static bool IsTable8SheetName(string sheetName) =>
        string.Equals(sheetName, "省统计", StringComparison.Ordinal) ||
        sheetName.EndsWith("-省统计", StringComparison.Ordinal);

    private static bool MatchesTable8Sheet(
        OpenXmlWorkbookReader reader,
        string sheetName) =>
        HeaderContains(reader, sheetName, "A2", "月") &&
        HeaderContains(reader, sheetName, "B2", "省份") &&
        IsOrdinaryTotalHeader(reader, sheetName, "D2") &&
        HeaderContains(reader, sheetName, "J2", "IPv6总流量");

    private static bool HeaderContains(
        OpenXmlWorkbookReader reader,
        string sheetName,
        string address,
        string expected) =>
        HeaderText.Contains(reader.GetText(sheetName, address), expected);

    private static bool IsOrdinaryTotalHeader(
        OpenXmlWorkbookReader reader,
        string sheetName,
        string address)
    {
        string header = HeaderText.Normalize(reader.GetText(sheetName, address));
        if (header.StartsWith("总流量", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        bool isIpv6Only = header.Contains("IPv6", StringComparison.OrdinalIgnoreCase) ||
                          header.Contains("V6", StringComparison.OrdinalIgnoreCase);
        return header.Contains("总流量", StringComparison.OrdinalIgnoreCase) && !isIpv6Only;
    }

    private static void AddTable8MonthMarkers(
        OpenXmlWorkbookReader reader,
        IEnumerable<string> sheetNames,
        ICollection<MonthMarker> markers,
        CancellationToken cancellationToken)
    {
        foreach (string sheetName in sheetNames)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (uint row in reader.GetPopulatedRows(sheetName).Where(row => row >= 4).Order())
            {
                string address = $"A{row}";
                string? value = reader.GetText(sheetName, address);
                if (HeaderText.Normalize(value).Length == 0)
                {
                    continue;
                }

                IReadOnlyList<MonthMarker> parsedMarkers = MonthTextParser.Extract(
                    $"8表 {sheetName}!{address}",
                    value);
                if (parsedMarkers.Count == 0)
                {
                    continue;
                }

                foreach (MonthMarker marker in parsedMarkers)
                {
                    markers.Add(marker);
                }

                break;
            }
        }
    }

    private static ValidationIssue[] CreateStructureIssues(string fileName, int kindCount) =>
        kindCount switch
        {
            0 =>
            [
                new ValidationIssue(
                    "WORKBOOK_STRUCTURE_UNKNOWN",
                    $"无法识别工作簿 '{fileName}' 的结构：缺少必要工作表或表头。",
                    fileName),
            ],
            > 1 =>
            [
                new ValidationIssue(
                    "WORKBOOK_KIND_AMBIGUOUS",
                    $"工作簿 '{fileName}' 同时匹配 {kindCount} 种来源表结构。",
                    fileName),
            ],
            _ => [],
        };

    private static WorkbookInspection CreateInspection(
        string path,
        IEnumerable<SourceWorkbookKind> kinds,
        IEnumerable<MonthMarker> markers,
        IEnumerable<ValidationIssue> issues) =>
        new(
            path,
            Array.AsReadOnly(kinds.ToArray()),
            Array.AsReadOnly(markers.Distinct().ToArray()),
            Array.AsReadOnly(issues.ToArray()));

    private static WorkbookInspection CreateUnreadableInspection(
        string path,
        string fileName)
    {
        var issue = new ValidationIssue(
            "WORKBOOK_UNREADABLE",
            $"无法读取工作簿 '{fileName}'：文件可能损坏、加密或不是有效的 OOXML 工作簿。",
            fileName);
        return CreateInspection(
            path,
            [],
            MonthTextParser.Extract(fileName, fileName),
            [issue]);
    }
}
