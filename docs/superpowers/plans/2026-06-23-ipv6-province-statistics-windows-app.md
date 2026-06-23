# IPv6 省级统计助手 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个完全离线、无需安装 .NET 或 Microsoft Office、可在 Windows 10/11 x64 绿色运行并从四张月度 Excel 表生成 31 份省级统计表的桌面软件。

**Architecture:** 使用五层 .NET 10 结构：WPF 只负责交互，Application 编排预检与生成，Domain 保存省份、月份、18 项指标和公式规则，OpenXml 基础设施读取与写入 OOXML，FileSystem 基础设施负责快照、日志和同卷原子发布。所有业务规则先以单元测试固定，再用合成工作簿做集成测试，最后在无 .NET、无 Office 的干净 Windows 10/11 x64 环境验收。

**Tech Stack:** .NET SDK 10.0.301、C#、WPF、DocumentFormat.OpenXml 3.5.1、xUnit 2.9.3、Microsoft.NET.Test.Sdk 18.7.0、xunit.runner.visualstudio 3.1.5、coverlet.collector 10.0.1、PowerShell 7/Windows PowerShell 5.1。

---

## 执行前提

1. 开始实施前必须调用 `superpowers:using-git-worktrees`，在隔离工作树中执行本计划。
2. 每个功能任务遵循红—绿—重构：先写失败测试，确认失败原因正确，再写最小实现。
3. 每次提交前运行该任务列出的完整验证命令；提交信息使用 `type(scope): subject`，正文说明动机和验证结果。
4. 任何提交主题或正文不得出现 AI Agent、代码生成工具或编码助手名称。
5. `5月统计数据/` 中的原始业务工作簿只用于本机最终验收，不得 `git add`、不得上传远端。
6. 若执行环境没有真实 Windows 10/11 x64，允许完成跨平台测试和 `win-x64` 发布，但不得宣称软件最终验收完成；必须在 Task 20 的 Windows 门槛处停止并说明阻塞条件。

## 最终文件结构

```text
Directory.Build.props
Directory.Packages.props
global.json
Ipv6ProvinceStatistics.sln
README.md
assets/
  templates/
    Ipv6ReportTemplate.xlsx
    Ipv6ReportTemplate.sha256
build/
  Publish-WinX64.ps1
  Verify-OfflinePolicy.ps1
  Verify-Portable.ps1
docs/
  superpowers/plans/2026-06-23-ipv6-province-statistics-windows-app.md
  superpowers/specs/2026-06-23-ipv6-province-statistics-windows-app-design.md
  user-guide.md
src/
  Ipv6ProvinceStatistics.Domain/
    Provinces/Province.cs
    Provinces/ProvinceCatalog.cs
    Reporting/ReportMonth.cs
    Reporting/MonthMarker.cs
    Reporting/MonthTextParser.cs
    Reporting/MonthResolver.cs
    Reporting/MetricKey.cs
    Reporting/ProvinceReportInput.cs
    Reporting/ReportCalculator.cs
    Validation/ValidationIssue.cs
  Ipv6ProvinceStatistics.Application/
    Abstractions/IWorkbookInspector.cs
    Abstractions/ISourceWorkbookReader.cs
    Abstractions/ITemplateReportExporter.cs
    Abstractions/ITaskWorkspaceManager.cs
    Abstractions/IOutputTransaction.cs
    Abstractions/IOperationLogger.cs
    Abstractions/IReportProcessingService.cs
    Models/SourceWorkbookKind.cs
    Models/WorkbookInspection.cs
    Models/SourceReadResult.cs
    Models/SourceFileSnapshot.cs
    Models/TaskWorkspace.cs
    Models/PreparedBatch.cs
    Models/ProcessingProgress.cs
    Models/PreflightResult.cs
    Models/GenerationResult.cs
    Services/ProvinceDataAssembler.cs
    Services/ReportProcessingService.cs
  Ipv6ProvinceStatistics.Infrastructure.OpenXml/
    AssemblyInfo.cs
    Reading/OpenXmlWorkbookReader.cs
    Reading/OpenXmlWorkbookInspector.cs
    Reading/OpenXmlSourceWorkbookReader.cs
    Reading/HeaderText.cs
    Extraction/ExtractorSupport.cs
    Extraction/Table1Extractor.cs
    Extraction/Table4Extractor.cs
    Extraction/Table5Extractor.cs
    Extraction/Table8Extractor.cs
    Templates/TemplateCellMap.cs
    Templates/FormulaManifest.cs
    Templates/TemplateResourceProvider.cs
    Templates/TemplateSanitizer.cs
    Templates/TemplateValidator.cs
    Templates/TemplateReportExporter.cs
    Templates/ReportWorkbookVerifier.cs
  Ipv6ProvinceStatistics.Infrastructure.FileSystem/
    AppPaths.cs
    TaskWorkspaceManager.cs
    OutputDirectoryNaming.cs
    OutputTransaction.cs
    JsonOperationLogger.cs
  Ipv6ProvinceStatistics.App/
    App.xaml
    App.xaml.cs
    MainWindow.xaml
    MainWindow.xaml.cs
    Themes/Colors.xaml
    Themes/Controls.xaml
    ViewModels/ObservableObject.cs
    ViewModels/AsyncRelayCommand.cs
    ViewModels/FileCardViewModel.cs
    ViewModels/MainWindowViewModel.cs
    Services/DialogService.cs
    Services/ShellService.cs
  Ipv6ProvinceStatistics.TemplateTool/
    Program.cs
tests/
  Ipv6ProvinceStatistics.UnitTests/
    Provinces/ProvinceCatalogTests.cs
    Reporting/MonthResolverTests.cs
    Reporting/ReportCalculatorTests.cs
    Services/ProvinceDataAssemblerTests.cs
    Services/ReportProcessingServiceTests.cs
    Services/ProcessingHarness.cs
  Ipv6ProvinceStatistics.IntegrationTests/
    Fixtures/TestWorkbookBuilder.cs
    Fixtures/TempTestPaths.cs
    Fixtures/WorkbookFixtures.cs
    Fixtures/TemplateFixtures.cs
    Fixtures/MonthlyFixtureBuilder.cs
    Reading/OpenXmlWorkbookReaderTests.cs
    Reading/OpenXmlWorkbookInspectorTests.cs
    Extraction/Table1ExtractorTests.cs
    Extraction/Table4ExtractorTests.cs
    Extraction/Table5ExtractorTests.cs
    Extraction/Table8ExtractorTests.cs
    Templates/TemplateSanitizerTests.cs
    Templates/TemplateReportExporterTests.cs
    FileSystem/TaskWorkspaceManagerTests.cs
    FileSystem/OutputTransactionTests.cs
    Logging/JsonOperationLoggerTests.cs
    EndToEnd/MonthlyGenerationTests.cs
  Ipv6ProvinceStatistics.WindowsTests/
    ViewModels/MainWindowViewModelTests.cs
    ViewModels/ViewModelHarness.cs
```

### Task 1: 搭建可复现的 .NET 10 解决方案

**Files:**
- Create: `global.json`
- Create: `Directory.Build.props`
- Create: `Directory.Packages.props`
- Create: `Ipv6ProvinceStatistics.sln`
- Create: `src/*/*.csproj`
- Create: `tests/*/*.csproj`
- Create: `src/Ipv6ProvinceStatistics.App/App.xaml`
- Create: `src/Ipv6ProvinceStatistics.App/App.xaml.cs`
- Create: `src/Ipv6ProvinceStatistics.App/MainWindow.xaml`
- Create: `src/Ipv6ProvinceStatistics.App/MainWindow.xaml.cs`

- [ ] **Step 1: 验证 SDK 版本**

Run:

```bash
dotnet --version
```

Expected: 输出 `10.0.301`，或 `global.json` 允许的更新 10.0 feature band；若没有 .NET 10 SDK，先安装官方 .NET 10 SDK，再继续。

- [ ] **Step 2: 创建解决方案和项目目录**

Run:

```bash
dotnet new sln -n Ipv6ProvinceStatistics --format sln
dotnet new classlib -n Ipv6ProvinceStatistics.Domain -o src/Ipv6ProvinceStatistics.Domain -f net10.0 --no-restore
dotnet new classlib -n Ipv6ProvinceStatistics.Application -o src/Ipv6ProvinceStatistics.Application -f net10.0 --no-restore
dotnet new classlib -n Ipv6ProvinceStatistics.Infrastructure.OpenXml -o src/Ipv6ProvinceStatistics.Infrastructure.OpenXml -f net10.0 --no-restore
dotnet new classlib -n Ipv6ProvinceStatistics.Infrastructure.FileSystem -o src/Ipv6ProvinceStatistics.Infrastructure.FileSystem -f net10.0 --no-restore
dotnet new console -n Ipv6ProvinceStatistics.TemplateTool -o src/Ipv6ProvinceStatistics.TemplateTool -f net10.0 --no-restore
dotnet new xunit -n Ipv6ProvinceStatistics.UnitTests -o tests/Ipv6ProvinceStatistics.UnitTests -f net10.0 --no-restore
dotnet new xunit -n Ipv6ProvinceStatistics.IntegrationTests -o tests/Ipv6ProvinceStatistics.IntegrationTests -f net10.0 --no-restore
dotnet new xunit -n Ipv6ProvinceStatistics.WindowsTests -o tests/Ipv6ProvinceStatistics.WindowsTests -f net10.0 --no-restore
mkdir -p src/Ipv6ProvinceStatistics.App
```

Expected: 命令全部退出 0，除 WPF 外的项目目录已生成。使用 `apply_patch` 删除模板生成的 `Class1.cs` 和 `UnitTest1.cs`，不要用 shell 删除命令。

- [ ] **Step 3: 固定 SDK、编译规则和包版本**

Create `global.json`:

```json
{
  "sdk": {
    "version": "10.0.301",
    "rollForward": "latestFeature"
  }
}
```

Create `Directory.Build.props`:

```xml
<Project>
  <PropertyGroup>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <LangVersion>latest</LangVersion>
    <TreatWarningsAsErrors>true</TreatWarningsAsErrors>
    <Deterministic>true</Deterministic>
    <EnableWindowsTargeting>true</EnableWindowsTargeting>
  </PropertyGroup>
</Project>
```

Create `Directory.Packages.props`:

```xml
<Project>
  <PropertyGroup>
    <ManagePackageVersionsCentrally>true</ManagePackageVersionsCentrally>
  </PropertyGroup>
  <ItemGroup>
    <PackageVersion Include="DocumentFormat.OpenXml" Version="3.5.1" />
    <PackageVersion Include="Microsoft.NET.Test.Sdk" Version="18.7.0" />
    <PackageVersion Include="xunit" Version="2.9.3" />
    <PackageVersion Include="xunit.runner.visualstudio" Version="3.1.5" />
    <PackageVersion Include="coverlet.collector" Version="10.0.1" />
  </ItemGroup>
</Project>
```

Create `src/Ipv6ProvinceStatistics.App/Ipv6ProvinceStatistics.App.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net10.0-windows</TargetFramework>
    <UseWPF>true</UseWPF>
    <AssemblyName>IPv6省级统计助手</AssemblyName>
    <RootNamespace>Ipv6ProvinceStatistics.App</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\Ipv6ProvinceStatistics.Application\Ipv6ProvinceStatistics.Application.csproj" />
    <ProjectReference Include="..\Ipv6ProvinceStatistics.Infrastructure.OpenXml\Ipv6ProvinceStatistics.Infrastructure.OpenXml.csproj" />
    <ProjectReference Include="..\Ipv6ProvinceStatistics.Infrastructure.FileSystem\Ipv6ProvinceStatistics.Infrastructure.FileSystem.csproj" />
  </ItemGroup>
</Project>
```

Configure references and the Open XML package:

```bash
dotnet add src/Ipv6ProvinceStatistics.Application reference src/Ipv6ProvinceStatistics.Domain
dotnet add src/Ipv6ProvinceStatistics.Infrastructure.OpenXml reference src/Ipv6ProvinceStatistics.Domain src/Ipv6ProvinceStatistics.Application
dotnet add src/Ipv6ProvinceStatistics.Infrastructure.OpenXml package DocumentFormat.OpenXml
dotnet add src/Ipv6ProvinceStatistics.Infrastructure.FileSystem reference src/Ipv6ProvinceStatistics.Domain src/Ipv6ProvinceStatistics.Application
dotnet add src/Ipv6ProvinceStatistics.TemplateTool reference src/Ipv6ProvinceStatistics.Infrastructure.OpenXml
dotnet add tests/Ipv6ProvinceStatistics.UnitTests reference src/Ipv6ProvinceStatistics.Domain src/Ipv6ProvinceStatistics.Application
dotnet add tests/Ipv6ProvinceStatistics.IntegrationTests reference src/Ipv6ProvinceStatistics.Domain src/Ipv6ProvinceStatistics.Application src/Ipv6ProvinceStatistics.Infrastructure.OpenXml src/Ipv6ProvinceStatistics.Infrastructure.FileSystem
dotnet add tests/Ipv6ProvinceStatistics.WindowsTests reference src/Ipv6ProvinceStatistics.App
```

For each test `.csproj`, retain these package references without inline versions:

```xml
<ItemGroup>
  <PackageReference Include="coverlet.collector" PrivateAssets="all" />
  <PackageReference Include="Microsoft.NET.Test.Sdk" />
  <PackageReference Include="xunit" />
  <PackageReference Include="xunit.runner.visualstudio" PrivateAssets="all" />
</ItemGroup>
```

Set `tests/Ipv6ProvinceStatistics.WindowsTests/Ipv6ProvinceStatistics.WindowsTests.csproj` to `<TargetFramework>net10.0-windows</TargetFramework>` and `<UseWPF>true</UseWPF>`.

- [ ] **Step 4: 创建最小 WPF 入口并加入解决方案**

Create `App.xaml`:

```xml
<Application x:Class="Ipv6ProvinceStatistics.App.App"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             StartupUri="MainWindow.xaml">
  <Application.Resources />
</Application>
```

Create `App.xaml.cs`:

```csharp
using System.Windows;
namespace Ipv6ProvinceStatistics.App;
public partial class App : Application { }
```

Create `MainWindow.xaml`:

```xml
<Window x:Class="Ipv6ProvinceStatistics.App.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IPv6 省级统计助手" Width="960" Height="680">
  <Grid>
    <TextBlock HorizontalAlignment="Center" VerticalAlignment="Center"
               FontSize="24" Text="IPv6 省级统计助手" />
  </Grid>
</Window>
```

Create `MainWindow.xaml.cs`:

```csharp
using System.Windows;
namespace Ipv6ProvinceStatistics.App;
public partial class MainWindow : Window
{
    public MainWindow() => InitializeComponent();
}
```

Run:

```bash
dotnet sln Ipv6ProvinceStatistics.sln add src/*/*.csproj tests/*/*.csproj
dotnet restore Ipv6ProvinceStatistics.sln
dotnet build Ipv6ProvinceStatistics.sln -c Debug --no-restore
```

Expected: `Build succeeded.`，0 warnings，0 errors。非 Windows 主机只构建 Windows 项目，不运行 `WindowsTests`。

- [ ] **Step 5: 提交解决方案骨架**

```bash
git add global.json Directory.Build.props Directory.Packages.props Ipv6ProvinceStatistics.sln src tests
git commit -m "chore(build): scaffold .NET solution" -m "Create the layered .NET 10 solution, pin package versions, and establish warning-free deterministic builds."
```

### Task 2: 实现 31 省标准化目录

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Domain/Provinces/Province.cs`
- Create: `src/Ipv6ProvinceStatistics.Domain/Provinces/ProvinceCatalog.cs`
- Create: `tests/Ipv6ProvinceStatistics.UnitTests/Provinces/ProvinceCatalogTests.cs`

- [ ] **Step 1: 写省份标准化失败测试**

```csharp
using Ipv6ProvinceStatistics.Domain.Provinces;

namespace Ipv6ProvinceStatistics.UnitTests.Provinces;

public sealed class ProvinceCatalogTests
{
    [Fact]
    public void All_contains_exactly_31_unique_provinces()
    {
        Assert.Equal(31, ProvinceCatalog.All.Count);
        Assert.Equal(31, ProvinceCatalog.All.Select(x => x.Name).Distinct().Count());
    }

    [Theory]
    [InlineData("北京市", "北京")]
    [InlineData("北   京", "北京")]
    [InlineData("北\u3000京", "北京")]
    [InlineData("内蒙古自治区", "内蒙古")]
    [InlineData("广西壮族自治区", "广西")]
    [InlineData("宁夏回族自治区", "宁夏")]
    [InlineData("新疆维吾尔族自治区", "新疆")]
    [InlineData("西藏自治区", "西藏")]
    public void TryResolve_maps_aliases_to_standard_names(string input, string expected)
    {
        Assert.True(ProvinceCatalog.TryResolve(input, out var province));
        Assert.Equal(expected, province.Name);
    }

    [Theory]
    [InlineData("云公司")]
    [InlineData("香港")]
    [InlineData("")]
    public void TryResolve_rejects_non_province_rows(string input)
    {
        Assert.False(ProvinceCatalog.TryResolve(input, out _));
    }
}
```

- [ ] **Step 2: 运行测试并确认因类型不存在而失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~ProvinceCatalogTests
```

Expected: FAIL，编译错误指出 `ProvinceCatalog` 或 `Province` 不存在。

- [ ] **Step 3: 写最小省份实现**

Create `Province.cs`:

```csharp
namespace Ipv6ProvinceStatistics.Domain.Provinces;

public readonly record struct Province(string Name)
{
    public override string ToString() => Name;
}
```

Create `ProvinceCatalog.cs`:

```csharp
using System.Collections.ObjectModel;
using System.Text;

namespace Ipv6ProvinceStatistics.Domain.Provinces;

public static class ProvinceCatalog
{
    private static readonly string[] Names =
    [
        "北京", "天津", "河北", "山西", "内蒙古", "辽宁", "吉林", "黑龙江",
        "上海", "江苏", "浙江", "安徽", "福建", "江西", "山东", "河南",
        "湖北", "湖南", "广东", "广西", "海南", "重庆", "四川", "贵州",
        "云南", "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆"
    ];

    private static readonly IReadOnlyDictionary<string, Province> Aliases = BuildAliases();

    public static IReadOnlyList<Province> All { get; } =
        new ReadOnlyCollection<Province>(Names.Select(name => new Province(name)).ToArray());

    public static bool TryResolve(string? raw, out Province province)
    {
        province = default;
        return !string.IsNullOrWhiteSpace(raw) && Aliases.TryGetValue(RemoveWhitespace(raw), out province);
    }

    private static IReadOnlyDictionary<string, Province> BuildAliases()
    {
        var suffixes = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["北京"] = "北京市", ["天津"] = "天津市", ["上海"] = "上海市", ["重庆"] = "重庆市",
            ["内蒙古"] = "内蒙古自治区", ["广西"] = "广西壮族自治区", ["宁夏"] = "宁夏回族自治区",
            ["新疆"] = "新疆维吾尔族自治区", ["西藏"] = "西藏自治区"
        };
        var map = new Dictionary<string, Province>(StringComparer.Ordinal);
        foreach (var name in Names)
        {
            var province = new Province(name);
            map[name] = province;
            map[suffixes.GetValueOrDefault(name, name + "省")] = province;
        }
        return new ReadOnlyDictionary<string, Province>(map);
    }

    private static string RemoveWhitespace(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var rune in value.EnumerateRunes())
            if (!Rune.IsWhiteSpace(rune)) builder.Append(rune.ToString());
        return builder.ToString();
    }
}
```

- [ ] **Step 4: 运行省份测试**

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~ProvinceCatalogTests
```

Expected: PASS，全部省份测试通过。

- [ ] **Step 5: 提交省份领域规则**

```bash
git add src/Ipv6ProvinceStatistics.Domain/Provinces tests/Ipv6ProvinceStatistics.UnitTests/Provinces
git commit -m "feat(domain): normalize province names" -m "Define the canonical 31-province catalog and explicit aliases for whitespace and administrative suffix variants."
```

### Task 3: 实现年月提取、共识与冲突检测

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Domain/Validation/ValidationIssue.cs`
- Create: `src/Ipv6ProvinceStatistics.Domain/Reporting/ReportMonth.cs`
- Create: `src/Ipv6ProvinceStatistics.Domain/Reporting/MonthMarker.cs`
- Create: `src/Ipv6ProvinceStatistics.Domain/Reporting/MonthTextParser.cs`
- Create: `src/Ipv6ProvinceStatistics.Domain/Reporting/MonthResolver.cs`
- Create: `tests/Ipv6ProvinceStatistics.UnitTests/Reporting/MonthResolverTests.cs`

- [ ] **Step 1: 写年月规则失败测试**

```csharp
using Ipv6ProvinceStatistics.Domain.Reporting;

namespace Ipv6ProvinceStatistics.UnitTests.Reporting;

public sealed class MonthResolverTests
{
    [Fact]
    public void Extract_recognizes_chinese_and_compact_year_months()
    {
        var markers = MonthTextParser.Extract("files", "1-2026年5月.xlsx 202605");
        Assert.Contains(markers, x => x.Year == 2026 && x.Month == 5);
    }

    [Fact]
    public void Resolve_combines_full_markers_with_month_only_marker()
    {
        MonthMarker[] markers =
        [
            new(2026, 5, "1表"), new(null, 5, "4表"),
            new(2026, 5, "5表"), new(2026, 5, "8表")
        ];
        var result = MonthResolver.Resolve(markers, null);
        Assert.True(result.IsValid);
        Assert.Equal(new ReportMonth(2026, 5), result.Month);
    }

    [Fact]
    public void Resolve_rejects_conflicting_complete_markers_even_with_override()
    {
        MonthMarker[] markers = [new(2026, 5, "1表"), new(2026, 4, "5表")];
        var result = MonthResolver.Resolve(markers, new ReportMonth(2026, 5));
        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, x => x.Code == "MONTH_CONFLICT");
    }

    [Fact]
    public void Resolve_uses_override_when_only_month_is_known()
    {
        MonthMarker[] markers = [new(null, 5, "4表")];
        var result = MonthResolver.Resolve(markers, new ReportMonth(2026, 5));
        Assert.True(result.IsValid);
        Assert.Equal(new ReportMonth(2026, 5), result.Month);
    }
}
```

- [ ] **Step 2: 运行测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~MonthResolverTests
```

Expected: FAIL，缺少 `MonthTextParser`、`MonthResolver` 和相关记录类型。

- [ ] **Step 3: 实现年月值对象和解析器**

```csharp
// ValidationIssue.cs
namespace Ipv6ProvinceStatistics.Domain.Validation;
public sealed record ValidationIssue(string Code, string Message, string? FileName = null,
    string? Sheet = null, string? Province = null, string? Cell = null);
```

```csharp
// ReportMonth.cs
namespace Ipv6ProvinceStatistics.Domain.Reporting;
public readonly record struct ReportMonth
{
    public ReportMonth(int year, int month)
    {
        if (year is < 2000 or > 9999) throw new ArgumentOutOfRangeException(nameof(year));
        if (month is < 1 or > 12) throw new ArgumentOutOfRangeException(nameof(month));
        Year = year;
        Month = month;
    }
    public int Year { get; }
    public int Month { get; }
    public string FolderName => $"{Year}年{Month:00}月统计结果";
    public string FileSuffix => $"{Year}年{Month:00}月";
}
```

```csharp
// MonthMarker.cs
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Domain.Reporting;
public sealed record MonthMarker(int? Year, int Month, string Source);
public sealed record MonthResolution(ReportMonth? Month, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid => Month is not null && Issues.Count == 0;
}
```

```csharp
// MonthTextParser.cs
using System.Text.RegularExpressions;
namespace Ipv6ProvinceStatistics.Domain.Reporting;
public static partial class MonthTextParser
{
    [GeneratedRegex(@"(?<!\d)(?<year>20\d{2})年(?<month>0?[1-9]|1[0-2])月", RegexOptions.CultureInvariant)]
    private static partial Regex ChinesePattern();
    [GeneratedRegex(@"(?<!\d)(?<year>20\d{2})(?<month>0[1-9]|1[0-2])(?!\d)", RegexOptions.CultureInvariant)]
    private static partial Regex CompactPattern();
    [GeneratedRegex(@"(?<!\d)(?<month>0?[1-9]|1[0-2])月", RegexOptions.CultureInvariant)]
    private static partial Regex MonthOnlyPattern();

    public static IReadOnlyList<MonthMarker> Extract(string source, string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return [];
        var result = new List<MonthMarker>();
        foreach (Match match in ChinesePattern().Matches(text))
            result.Add(new(int.Parse(match.Groups["year"].Value), int.Parse(match.Groups["month"].Value), source));
        foreach (Match match in CompactPattern().Matches(text))
            result.Add(new(int.Parse(match.Groups["year"].Value), int.Parse(match.Groups["month"].Value), source));
        if (result.Count == 0)
            foreach (Match match in MonthOnlyPattern().Matches(text))
                result.Add(new(null, int.Parse(match.Groups["month"].Value), source));
        return result.Distinct().ToArray();
    }
}
```

Create `MonthResolver.cs`:

```csharp
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Domain.Reporting;
public static class MonthResolver
{
    public static MonthResolution Resolve(IReadOnlyCollection<MonthMarker> markers, ReportMonth? selected)
    {
        var full = markers.Where(x => x.Year.HasValue)
            .Select(x => new ReportMonth(x.Year!.Value, x.Month)).Distinct().ToArray();
        if (full.Length > 1)
            return new(null, [new("MONTH_CONFLICT", "源工作簿包含互相冲突的完整统计年月。")]);
        var month = full.SingleOrDefault();
        if (month == default && selected is null)
            return new(null, [new("MONTH_YEAR_MISSING", "无法确定统计年份，请选择年月。")]);
        month = month == default ? selected!.Value : month;
        if (selected is not null && full.Length == 1 && selected.Value != month)
            return new(null, [new("MONTH_CONFLICT", "手动选择的年月与源工作簿明确年月冲突。")]);
        if (markers.Any(x => x.Month != month.Month))
            return new(null, [new("MONTH_CONFLICT", "源工作簿月份不一致。")]);
        return new(month, []);
    }
}
```

- [ ] **Step 4: 运行年月测试**

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~MonthResolverTests
```

Expected: PASS，4 个测试通过。

- [ ] **Step 5: 提交年月规则**

```bash
git add src/Ipv6ProvinceStatistics.Domain/Reporting src/Ipv6ProvinceStatistics.Domain/Validation tests/Ipv6ProvinceStatistics.UnitTests/Reporting
git commit -m "feat(domain): resolve reporting month" -m "Parse full and partial month markers, merge consistent sources, and reject explicit conflicts that manual selection cannot bypass."
```

### Task 4: 定义 18 项指标并组装 31 省数据

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Domain/Reporting/MetricKey.cs`
- Create: `src/Ipv6ProvinceStatistics.Domain/Reporting/ProvinceReportInput.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Models/SourceWorkbookKind.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Models/SourceReadResult.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Services/ProvinceDataAssembler.cs`
- Create: `tests/Ipv6ProvinceStatistics.UnitTests/Services/ProvinceDataAssemblerTests.cs`

- [ ] **Step 1: 写缺失指标和成功组装测试**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Application.Services;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;

namespace Ipv6ProvinceStatistics.UnitTests.Services;

public sealed class ProvinceDataAssemblerTests
{
    [Fact]
    public void Assemble_builds_all_31_inputs_when_every_metric_exists()
    {
        var result = ProvinceDataAssembler.Assemble(CreateComplete());
        Assert.Empty(result.Issues);
        Assert.Equal(31, result.Reports.Count);
        Assert.Equal(1m, result.Reports[new Province("北京")][MetricKey.MetroTotal]);
    }

    [Fact]
    public void Assemble_reports_missing_province_metric()
    {
        var sources = CreateComplete().ToList();
        var table8 = sources.Single(x => x.Kind == SourceWorkbookKind.Table8);
        var values = table8.Values.Where(x => x.Key.Name != "北京").ToDictionary(x => x.Key, x => x.Value);
        sources[sources.IndexOf(table8)] = table8 with { Values = values };
        var result = ProvinceDataAssembler.Assemble(sources);
        Assert.Contains(result.Issues, x => x.Code == "PROVINCE_MISSING" && x.Province == "北京");
        Assert.Empty(result.Reports);
    }

    private static SourceReadResult[] CreateComplete() =>
        Enum.GetValues<SourceWorkbookKind>().Select(kind =>
        {
            var provinces = ProvinceCatalog.All.ToDictionary(
                province => province,
                province => (IReadOnlyDictionary<MetricKey, decimal>)
                    ProvinceDataAssembler.ExpectedMetricsByKind[kind].ToDictionary(key => key, key => 1m));
            return new SourceReadResult(kind, provinces, []);
        }).ToArray();
}
```

- [ ] **Step 2: 运行测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~ProvinceDataAssemblerTests
```

Expected: FAIL，缺少指标、来源结果和组装器类型。

- [ ] **Step 3: 实现指标和来源模型**

Create `MetricKey.cs`:

```csharp
namespace Ipv6ProvinceStatistics.Domain.Reporting;
public enum MetricKey
{
    MetroTotal, MetroIpv6, MobileCoreTotal, MobileCoreIpv6,
    InternetTotal, InternetIpv6, InternetOneGTotal, InternetOneGIpv6,
    IdcTotal, IdcIpv6, IdcTenGTotal, IdcTenGIpv6,
    HumanTotal, HumanIpv6, IotTotal, IotIpv6,
    BroadbandTotal, BroadbandIpv6
}
```

Create `ProvinceReportInput.cs`:

```csharp
namespace Ipv6ProvinceStatistics.Domain.Reporting;
public sealed class ProvinceReportInput
{
    private readonly IReadOnlyDictionary<MetricKey, decimal> _values;
    public ProvinceReportInput(IReadOnlyDictionary<MetricKey, decimal> values)
    {
        var missing = Enum.GetValues<MetricKey>().Except(values.Keys).ToArray();
        if (missing.Length > 0)
            throw new ArgumentException($"Missing metrics: {string.Join(',', missing)}", nameof(values));
        _values = new Dictionary<MetricKey, decimal>(values);
    }
    public decimal this[MetricKey key] => _values[key];
    public IReadOnlyDictionary<MetricKey, decimal> Values => _values;
}
```

Create application models:

```csharp
// SourceWorkbookKind.cs
namespace Ipv6ProvinceStatistics.Application.Models;
public enum SourceWorkbookKind { Table1, Table4, Table5, Table8 }
```

```csharp
// SourceReadResult.cs
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record SourceReadResult(SourceWorkbookKind Kind,
    IReadOnlyDictionary<Province, IReadOnlyDictionary<MetricKey, decimal>> Values,
    IReadOnlyList<ValidationIssue> Issues);
public sealed record ProvinceAssemblyResult(IReadOnlyDictionary<Province, ProvinceReportInput> Reports,
    IReadOnlyList<ValidationIssue> Issues);
```

- [ ] **Step 4: 实现组装器并运行测试**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Application.Services;

public static class ProvinceDataAssembler
{
    public static IReadOnlyDictionary<SourceWorkbookKind, IReadOnlySet<MetricKey>> ExpectedMetricsByKind { get; } =
        new Dictionary<SourceWorkbookKind, IReadOnlySet<MetricKey>>
        {
            [SourceWorkbookKind.Table1] = new HashSet<MetricKey>
                { MetricKey.MetroTotal, MetricKey.MetroIpv6, MetricKey.MobileCoreTotal, MetricKey.MobileCoreIpv6 },
            [SourceWorkbookKind.Table4] = new HashSet<MetricKey>
                { MetricKey.HumanTotal, MetricKey.HumanIpv6, MetricKey.IotTotal, MetricKey.IotIpv6 },
            [SourceWorkbookKind.Table5] = new HashSet<MetricKey>
                { MetricKey.InternetTotal, MetricKey.InternetIpv6, MetricKey.InternetOneGTotal,
                  MetricKey.InternetOneGIpv6, MetricKey.IdcTotal, MetricKey.IdcIpv6,
                  MetricKey.IdcTenGTotal, MetricKey.IdcTenGIpv6 },
            [SourceWorkbookKind.Table8] = new HashSet<MetricKey>
                { MetricKey.BroadbandTotal, MetricKey.BroadbandIpv6 }
        };

    public static ProvinceAssemblyResult Assemble(IReadOnlyCollection<SourceReadResult> sources)
    {
        var issues = sources.SelectMany(x => x.Issues).ToList();
        foreach (var kind in Enum.GetValues<SourceWorkbookKind>())
            if (sources.Count(x => x.Kind == kind) != 1)
                issues.Add(new("SOURCE_KIND_COUNT", $"{kind} 必须且只能出现一次。"));
        if (issues.Count > 0) return new(new Dictionary<Province, ProvinceReportInput>(), issues);

        var reports = new Dictionary<Province, ProvinceReportInput>();
        foreach (var province in ProvinceCatalog.All)
        {
            var metrics = new Dictionary<MetricKey, decimal>();
            foreach (var source in sources)
            {
                if (!source.Values.TryGetValue(province, out var sourceMetrics))
                {
                    issues.Add(new("PROVINCE_MISSING", $"{source.Kind} 缺少省份 {province.Name}。",
                        Province: province.Name));
                    continue;
                }
                foreach (var expected in ExpectedMetricsByKind[source.Kind])
                    if (sourceMetrics.TryGetValue(expected, out var value)) metrics[expected] = value;
                    else issues.Add(new("METRIC_MISSING", $"{source.Kind} 的 {province.Name} 缺少 {expected}。",
                        Province: province.Name));
            }
            if (issues.Count == 0) reports[province] = new ProvinceReportInput(metrics);
        }
        return issues.Count == 0 ? new(reports, []) : new(new Dictionary<Province, ProvinceReportInput>(), issues);
    }
}
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~ProvinceDataAssemblerTests
```

Expected: PASS，完整数据得到 31 省；缺失省份时报告错误且不返回部分结果。

- [ ] **Step 5: 提交指标组装**

```bash
git add src/Ipv6ProvinceStatistics.Domain/Reporting src/Ipv6ProvinceStatistics.Application/Models src/Ipv6ProvinceStatistics.Application/Services tests/Ipv6ProvinceStatistics.UnitTests/Services
git commit -m "feat(application): assemble province metrics" -m "Define all 18 source metrics and require one complete record for every canonical province before generation."
```

### Task 5: 实现模板公式的独立计算器

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Domain/Reporting/ReportCalculator.cs`
- Create: `tests/Ipv6ProvinceStatistics.UnitTests/Reporting/ReportCalculatorTests.cs`

- [ ] **Step 1: 写公式和除零失败测试**

```csharp
using Ipv6ProvinceStatistics.Domain.Reporting;

namespace Ipv6ProvinceStatistics.UnitTests.Reporting;

public sealed class ReportCalculatorTests
{
    [Fact]
    public void Calculate_matches_template_formulas()
    {
        var input = CreateAll(100m);
        var result = ReportCalculator.Calculate(input);
        Assert.Empty(result.Issues);
        var factor = 3600m * 24m / 8m / 1024m / 1024m;
        Assert.Equal(100m * factor, result.Values["C2"]);
        Assert.Equal(result.Values["E2"] / result.Values["C2"], result.Values["F2"]);
        Assert.Equal(result.Values["G9"] * 0.5m, result.Values["G10"]);
        Assert.Equal(37, result.Values.Count);
    }

    [Fact]
    public void Calculate_rejects_zero_denominator()
    {
        var result = ReportCalculator.Calculate(CreateAll(0m));
        Assert.Contains(result.Issues, x => x.Code == "FORMULA_DIVIDE_BY_ZERO");
        Assert.Empty(result.Values);
    }

    private static ProvinceReportInput CreateAll(decimal value) =>
        new(Enum.GetValues<MetricKey>().ToDictionary(key => key, key => value));
}
```

- [ ] **Step 2: 运行测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~ReportCalculatorTests
```

Expected: FAIL，`ReportCalculator` 不存在。

- [ ] **Step 3: 实现全部 37 个缓存结果**

```csharp
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Domain.Reporting;
public sealed record FormulaCalculationResult(IReadOnlyDictionary<string, decimal> Values,
    IReadOnlyList<ValidationIssue> Issues);
public static class ReportCalculator
{
    private const decimal GbpsToPbPerDay = 3600m * 24m / 8m / 1024m / 1024m;
    public static FormulaCalculationResult Calculate(ProvinceReportInput input)
    {
        var v = new Dictionary<string, decimal>
        {
            ["C2"] = input[MetricKey.MetroTotal] * GbpsToPbPerDay,
            ["E2"] = input[MetricKey.MetroIpv6] * GbpsToPbPerDay,
            ["C9"] = input[MetricKey.MobileCoreTotal] * GbpsToPbPerDay,
            ["E9"] = input[MetricKey.MobileCoreIpv6] * GbpsToPbPerDay,
            ["D4"] = input[MetricKey.BroadbandTotal], ["E4"] = input[MetricKey.BroadbandIpv6],
            ["D5"] = input[MetricKey.InternetTotal] * GbpsToPbPerDay,
            ["E5"] = input[MetricKey.InternetIpv6] * GbpsToPbPerDay,
            ["D6"] = input[MetricKey.InternetOneGTotal] * GbpsToPbPerDay,
            ["E6"] = input[MetricKey.InternetOneGIpv6] * GbpsToPbPerDay,
            ["D7"] = input[MetricKey.IdcTotal] * GbpsToPbPerDay,
            ["E7"] = input[MetricKey.IdcIpv6] * GbpsToPbPerDay,
            ["D8"] = input[MetricKey.IdcTenGTotal] * GbpsToPbPerDay,
            ["E8"] = input[MetricKey.IdcTenGIpv6] * GbpsToPbPerDay,
            ["D10"] = input[MetricKey.HumanTotal], ["E10"] = input[MetricKey.HumanIpv6],
            ["D11"] = input[MetricKey.IotTotal], ["E11"] = input[MetricKey.IotIpv6]
        };
        v["C3"] = v["C2"] - v["C9"] * 0.9m;
        v["E3"] = v["E2"] - v["E9"] * 0.9m;
        decimal[] denominators =
        [
            v["C2"], v["C3"], v["D4"], v["D5"], v["D6"], v["D7"], v["D8"], v["C9"],
            v["D10"], v["D11"], v["C3"] + v["C9"], v["D4"] + v["D5"] + v["D7"],
            v["D10"] + v["D11"]
        ];
        if (denominators.Any(x => x == 0m))
            return new(new Dictionary<string, decimal>(),
                [new("FORMULA_DIVIDE_BY_ZERO", "模板公式存在分母为 0 的结果。")]);
        v["F2"] = v["E2"] / v["C2"]; v["F3"] = v["E3"] / v["C3"];
        v["F4"] = v["E4"] / v["D4"]; v["F5"] = v["E5"] / v["D5"];
        v["F6"] = v["E6"] / v["D6"]; v["F7"] = v["E7"] / v["D7"];
        v["F8"] = v["E8"] / v["D8"]; v["F9"] = v["E9"] / v["C9"];
        v["F10"] = v["E10"] / v["D10"]; v["F11"] = v["E11"] / v["D11"];
        v["G3"] = v["C3"] / (v["C3"] + v["C9"]);
        v["G4"] = v["D4"] / (v["D4"] + v["D5"] + v["D7"]);
        v["G5"] = v["D5"] / (v["D4"] + v["D5"] + v["D7"]);
        v["G7"] = v["D7"] / (v["D4"] + v["D5"] + v["D7"]);
        v["G9"] = v["C9"] / (v["C3"] + v["C9"]);
        v["G10"] = v["G9"] * (v["D10"] / (v["D10"] + v["D11"]));
        v["G11"] = v["G9"] * (v["D11"] / (v["D10"] + v["D11"]));
        return new(v, []);
    }
}
```

- [ ] **Step 4: 运行公式测试和全量单元测试**

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~ReportCalculatorTests
dotnet test tests/Ipv6ProvinceStatistics.UnitTests
```

Expected: 两条命令均 PASS，公式结果字典恰好包含 37 个地址。

- [ ] **Step 5: 提交公式计算器**

```bash
git add src/Ipv6ProvinceStatistics.Domain/Reporting/ReportCalculator.cs tests/Ipv6ProvinceStatistics.UnitTests/Reporting/ReportCalculatorTests.cs
git commit -m "feat(domain): calculate template formulas" -m "Reproduce all report formulas with decimal arithmetic and fail preflight when any required denominator is zero."
```

### Task 6: 建立 OOXML 单元格读取层和合成工作簿夹具

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/AssemblyInfo.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlWorkbookReader.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Fixtures/TestWorkbookBuilder.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Fixtures/TempTestPaths.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Reading/OpenXmlWorkbookReaderTests.cs`

- [ ] **Step 1: 写共享字符串、数字、公式缓存和 sheet 读取失败测试**

```csharp
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.IntegrationTests.Reading;

public sealed class OpenXmlWorkbookReaderTests
{
    [Fact]
    public void Reader_resolves_supported_cell_types_without_excel()
    {
        var path = TempFiles.Next("示例 workbook.xlsx");
        TestWorkbookBuilder.Create(path,
            new TestSheet("数据", [
                TestCell.SharedText("A1", "北京市"), TestCell.InlineText("B1", "中国联通"),
                TestCell.Number("C1", "6.5E3"), TestCell.Formula("D1", "C1/2", "3250")
            ]));
        using var reader = OpenXmlWorkbookReader.Open(path, false);
        Assert.Equal(["数据"], reader.SheetNames);
        Assert.Equal("北京市", reader.GetText("数据", "A1"));
        Assert.Equal("中国联通", reader.GetText("数据", "B1"));
        Assert.True(reader.TryGetDecimal("数据", "C1", out var number));
        Assert.Equal(6500m, number);
        Assert.Equal("C1/2", reader.GetFormula("数据", "D1"));
        Assert.Equal([1], reader.GetPopulatedRows("数据"));
    }
}
```

- [ ] **Step 2: 运行测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~OpenXmlWorkbookReaderTests
```

Expected: FAIL，缺少 `TestWorkbookBuilder` 和 `OpenXmlWorkbookReader`。

- [ ] **Step 3: 创建完整的合成工作簿生成器**

```csharp
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

public enum TestCellKind { SharedText, InlineText, Number, Formula }
public sealed record TestCell(string Address, TestCellKind Kind, string Value, string? CachedValue = null)
{
    public static TestCell SharedText(string address, string value) => new(address, TestCellKind.SharedText, value);
    public static TestCell InlineText(string address, string value) => new(address, TestCellKind.InlineText, value);
    public static TestCell Number(string address, string value) => new(address, TestCellKind.Number, value);
    public static TestCell Formula(string address, string formula, string cached) =>
        new(address, TestCellKind.Formula, formula, cached);
}
public sealed record TestSheet(string Name, IReadOnlyList<TestCell> Cells);

public static class TestWorkbookBuilder
{
    public static void Create(string path, params TestSheet[] sheets)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sharedPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedPart.SharedStringTable = new SharedStringTable();
        var sheetCollection = workbookPart.Workbook.AppendChild(new Sheets());
        uint id = 1;
        foreach (var spec in sheets)
        {
            var part = workbookPart.AddNewPart<WorksheetPart>();
            var data = new SheetData();
            part.Worksheet = new Worksheet(data);
            foreach (var group in spec.Cells.GroupBy(cell => RowNumber(cell.Address)).OrderBy(group => group.Key))
            {
                var row = new Row { RowIndex = group.Key };
                foreach (var specCell in group.OrderBy(cell => cell.Address, StringComparer.Ordinal))
                    row.Append(CreateCell(specCell, sharedPart.SharedStringTable));
                data.Append(row);
            }
            part.Worksheet.Save();
            sheetCollection.Append(new Sheet { Id = workbookPart.GetIdOfPart(part), SheetId = id++, Name = spec.Name });
        }
        sharedPart.SharedStringTable.Save();
        workbookPart.Workbook.Save();
    }

    private static Cell CreateCell(TestCell spec, SharedStringTable shared)
    {
        var cell = new Cell { CellReference = spec.Address };
        switch (spec.Kind)
        {
            case TestCellKind.SharedText:
                shared.AppendChild(new SharedStringItem(new Text(spec.Value)));
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue((shared.Count() - 1).ToString(CultureInfo.InvariantCulture));
                break;
            case TestCellKind.InlineText:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(spec.Value));
                break;
            case TestCellKind.Number:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(spec.Value);
                break;
            case TestCellKind.Formula:
                cell.CellFormula = new CellFormula(spec.Value);
                cell.CellValue = new CellValue(spec.CachedValue!);
                break;
        }
        return cell;
    }

    private static uint RowNumber(string address) =>
        uint.Parse(new string(address.SkipWhile(char.IsLetter).ToArray()), CultureInfo.InvariantCulture);
}
```

Create `TempTestPaths.cs`:

```csharp
namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

public static class TempDirectories
{
    private static readonly string SessionRoot = Path.Combine(Path.GetTempPath(),
        "Ipv6ProvinceStatistics.Tests", Guid.NewGuid().ToString("N"));

    static TempDirectories()
    {
        Directory.CreateDirectory(SessionRoot);
        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            try { if (Directory.Exists(SessionRoot)) Directory.Delete(SessionRoot, true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        };
    }

    public static string Next()
    {
        var path = Path.Combine(SessionRoot, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }
}

public static class TempFiles
{
    public static string Next(string fileName) => Path.Combine(TempDirectories.Next(), fileName);
}
```

- [ ] **Step 4: 实现只读 OOXML reader 并运行测试**

Create `AssemblyInfo.cs`:

```csharp
using System.Runtime.CompilerServices;
[assembly: InternalsVisibleTo("Ipv6ProvinceStatistics.IntegrationTests")]
```

Create `OpenXmlWorkbookReader.cs`:

```csharp
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

internal sealed class OpenXmlWorkbookReader : IDisposable
{
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbook;
    private readonly Dictionary<string, WorksheetPart> _sheets;

    private OpenXmlWorkbookReader(SpreadsheetDocument document)
    {
        _document = document;
        _workbook = document.WorkbookPart ?? throw new InvalidDataException("Workbook part missing.");
        _sheets = _workbook.Workbook.Sheets!.Elements<Sheet>().ToDictionary(
            sheet => sheet.Name!.Value!,
            sheet => (WorksheetPart)_workbook.GetPartById(sheet.Id!.Value!),
            StringComparer.Ordinal);
    }

    public static OpenXmlWorkbookReader Open(string path, bool editable) =>
        new(SpreadsheetDocument.Open(path, editable));
    public IReadOnlyList<string> SheetNames => _sheets.Keys.ToArray();
    public bool HasSheet(string name) => _sheets.ContainsKey(name);

    public string? GetText(string sheet, string address)
    {
        var cell = FindCell(sheet, address);
        if (cell is null) return null;
        var raw = cell.CellValue?.InnerText;
        return cell.DataType?.Value switch
        {
            CellValues.SharedString => ResolveSharedString(raw),
            CellValues.InlineString => cell.InlineString?.InnerText,
            CellValues.Boolean => raw == "1" ? "TRUE" : "FALSE",
            _ => raw
        };
    }

    public bool TryGetDecimal(string sheet, string address, out decimal value) =>
        decimal.TryParse(GetText(sheet, address), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    public string? GetFormula(string sheet, string address) => FindCell(sheet, address)?.CellFormula?.Text;
    public IReadOnlyList<uint> GetPopulatedRows(string sheet) =>
        GetPart(sheet).Worksheet.Descendants<Row>().Select(row => row.RowIndex!.Value).ToArray();
    public Cell? FindCell(string sheet, string address) => GetPart(sheet).Worksheet.Descendants<Cell>()
        .FirstOrDefault(cell => string.Equals(cell.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
    public void Dispose() => _document.Dispose();

    private WorksheetPart GetPart(string sheet) => _sheets.TryGetValue(sheet, out var part)
        ? part : throw new KeyNotFoundException($"Sheet not found: {sheet}");
    private string? ResolveSharedString(string? index) =>
        int.TryParse(index, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? _workbook.SharedStringTablePart?.SharedStringTable?.ElementAtOrDefault(value)?.InnerText
            : null;
}
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~OpenXmlWorkbookReaderTests
```

Expected: PASS；测试进程中没有启动 Excel 或 Office 组件。

- [ ] **Step 5: 提交 OOXML 读取基础**

```bash
git add src/Ipv6ProvinceStatistics.Infrastructure.OpenXml tests/Ipv6ProvinceStatistics.IntegrationTests/Fixtures tests/Ipv6ProvinceStatistics.IntegrationTests/Reading
git commit -m "feat(excel): read OOXML cell values" -m "Resolve shared, inline, numeric, and formula-cached cells directly from XLSX packages without Office automation."
```

### Task 7: 按结构识别 1、4、5、8 表并提取年月标记

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Application/Abstractions/IWorkbookInspector.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Models/WorkbookInspection.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/HeaderText.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlWorkbookInspector.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Fixtures/WorkbookFixtures.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Reading/OpenXmlWorkbookInspectorTests.cs`

- [ ] **Step 1: 写别名识别、表头签名和歧义失败测试**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Reading;

public sealed class OpenXmlWorkbookInspectorTests
{
[Theory]
[InlineData("IDC汇总 (客户)")]
[InlineData("IDC汇总（客户）")]
public async Task Inspector_identifies_table5_with_both_idc_aliases(string idcName)
{
    var path = WorkbookFixtures.CreateTable5(idcName, "5-202605.xlsx");
    var result = await new OpenXmlWorkbookInspector().InspectAsync(path, CancellationToken.None);
    Assert.Equal([SourceWorkbookKind.Table5], result.MatchingKinds);
    Assert.Contains(result.MonthMarkers, x => x.Year == 2026 && x.Month == 5);
}

[Theory]
[InlineData("省统计")]
[InlineData("1-省统计")]
public async Task Inspector_identifies_table8_sheet_alias(string sheetName)
{
    var path = WorkbookFixtures.CreateTable8(sheetName, "8-202605.xlsx");
    var result = await new OpenXmlWorkbookInspector().InspectAsync(path, CancellationToken.None);
    Assert.Equal([SourceWorkbookKind.Table8], result.MatchingKinds);
}

[Fact]
public async Task Inspector_reports_workbook_matching_more_than_one_kind()
{
    var path = WorkbookFixtures.CreateAmbiguous();
    var result = await new OpenXmlWorkbookInspector().InspectAsync(path, CancellationToken.None);
    Assert.Contains(result.Issues, x => x.Code == "WORKBOOK_KIND_AMBIGUOUS");
}

[Fact]
public async Task Inspector_reports_corrupt_or_encrypted_package_as_unreadable()
{
    var path = TempFiles.Next("corrupt.xlsx");
    await File.WriteAllBytesAsync(path, [0x01, 0x02, 0x03]);
    var result = await new OpenXmlWorkbookInspector().InspectAsync(path, CancellationToken.None);
    var issue = Assert.Single(result.Issues);
    Assert.Equal("WORKBOOK_UNREADABLE", issue.Code);
    Assert.Equal("corrupt.xlsx", issue.FileName);
}

[Fact]
public async Task Inspector_reports_missing_required_header_as_unknown_structure()
{
    var path = TempFiles.Next("8-missing-J-header.xlsx");
    TestWorkbookBuilder.Create(path, new TestSheet("1-省统计", [
        TestCell.SharedText("A2", "月"), TestCell.SharedText("B2", "省份"),
        TestCell.SharedText("D2", "总流量"), TestCell.Number("A4", "202605")
    ]));
    var result = await new OpenXmlWorkbookInspector().InspectAsync(path, CancellationToken.None);
    Assert.Contains(result.Issues, issue => issue.Code == "WORKBOOK_STRUCTURE_UNKNOWN");
}
}
```

Create `WorkbookFixtures.cs`:

```csharp
namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

public static class WorkbookFixtures
{
    public static string CreateTable5(string idcSheetName, string fileName)
    {
        var path = TempFiles.Next(fileName);
        TestWorkbookBuilder.Create(path,
            new TestSheet("互联网专线汇总", [
                TestCell.SharedText("E2", "总流量"), TestCell.SharedText("F2", "IPv6流量"),
                TestCell.SharedText("K2", "1G及以上总流量"), TestCell.SharedText("L2", "1G及以上IPv6流量"),
                TestCell.SharedText("M2", "省")
            ]),
            new TestSheet(idcSheetName, [
                TestCell.SharedText("F2", "总流量"), TestCell.SharedText("G2", "V6日流量"),
                TestCell.SharedText("M2", "10G及以上总流量"), TestCell.SharedText("N2", "V6日流量"),
                TestCell.SharedText("O2", "省")
            ]));
        return path;
    }

    public static string CreateTable8(string sheetName, string fileName)
    {
        var path = TempFiles.Next(fileName);
        TestWorkbookBuilder.Create(path, new TestSheet(sheetName, [
            TestCell.SharedText("A2", "月"), TestCell.SharedText("B2", "省份"), TestCell.SharedText("D2", "总流量"),
            TestCell.SharedText("J2", "IPv6总流量"), TestCell.Number("A4", "202605")
        ]));
        return path;
    }

    public static string CreateAmbiguous()
    {
        var path = TempFiles.Next("混合结构-202605.xlsx");
        TestWorkbookBuilder.Create(path,
            new TestSheet("分省统计表", [
                TestCell.SharedText("A1", "2026年5月"),
                TestCell.SharedText("B2", "省份"), TestCell.SharedText("C2", "运营商"),
                TestCell.SharedText("D3", "城域网出口IPv4和IPv6总流量"),
                TestCell.SharedText("E3", "城域网出口IPv6总流量"),
                TestCell.SharedText("G3", "移动核心网出口IPv4和IPv6总流量"),
                TestCell.SharedText("H3", "移动核心网出口IPv6总流量")
            ]),
            TrafficSheet("人网统计"), TrafficSheet("物网统计"));
        return path;
    }

    private static TestSheet TrafficSheet(string name) => new(name, [
        TestCell.SharedText("H4", "prov_id"), TestCell.SharedText("R4", "总计日均流量(PB)"),
        TestCell.SharedText("S4", "IPV6日均流量(PB)")
    ]);
}
```

- [ ] **Step 2: 运行识别测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~OpenXmlWorkbookInspectorTests
```

Expected: FAIL，缺少 inspector 接口、模型与实现。

- [ ] **Step 3: 定义识别契约和表头正规化**

```csharp
// IWorkbookInspector.cs
using Ipv6ProvinceStatistics.Application.Models;
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public interface IWorkbookInspector
{
    Task<WorkbookInspection> InspectAsync(string path, CancellationToken cancellationToken);
}
```

```csharp
// WorkbookInspection.cs
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record WorkbookInspection(string Path, IReadOnlyList<SourceWorkbookKind> MatchingKinds,
    IReadOnlyList<MonthMarker> MonthMarkers, IReadOnlyList<ValidationIssue> Issues);
```

```csharp
// HeaderText.cs
using System.Text;
namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
internal static class HeaderText
{
    public static string Normalize(string? value)
    {
        if (value is null) return string.Empty;
        var builder = new StringBuilder(value.Length);
        foreach (var rune in value.EnumerateRunes())
            if (!Rune.IsWhiteSpace(rune)) builder.Append(rune.ToString());
        return builder.ToString().Replace('（', '(').Replace('）', ')');
    }
    public static bool Contains(string? value, string expected) =>
        Normalize(value).Contains(Normalize(expected), StringComparison.OrdinalIgnoreCase);
}
```

- [ ] **Step 4: 实现结构识别和年月标记提取**

```csharp
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

public sealed class OpenXmlWorkbookInspector : IWorkbookInspector
{
    public Task<WorkbookInspection> InspectAsync(string path, CancellationToken cancellationToken)
    {
        try { return Task.FromResult(Inspect(path, cancellationToken)); }
        catch (OperationCanceledException) { throw; }
        catch (Exception)
        {
            return Task.FromResult(new WorkbookInspection(path, [], [],
                [new("WORKBOOK_UNREADABLE", $"无法读取工作簿 {Path.GetFileName(path)}。",
                    Path.GetFileName(path))]));
        }
    }

    private static WorkbookInspection Inspect(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var book = OpenXmlWorkbookReader.Open(path, false);
        var kinds = new List<SourceWorkbookKind>();
        if (IsTable1(book)) kinds.Add(SourceWorkbookKind.Table1);
        if (IsTable4(book)) kinds.Add(SourceWorkbookKind.Table4);
        if (IsTable5(book)) kinds.Add(SourceWorkbookKind.Table5);
        if (IsTable8(book)) kinds.Add(SourceWorkbookKind.Table8);
        var markers = MonthTextParser.Extract(Path.GetFileName(path), Path.GetFileName(path)).ToList();
        if (kinds.Contains(SourceWorkbookKind.Table1))
            markers.AddRange(MonthTextParser.Extract("1表标题", book.GetText("分省统计表", "A1")));
        if (kinds.Contains(SourceWorkbookKind.Table8))
        {
            var sheet = book.SheetNames.Single(name => name == "省统计" || name.EndsWith("-省统计", StringComparison.Ordinal));
            var monthCell = book.GetPopulatedRows(sheet).Where(row => row >= 4).Select(row => book.GetText(sheet, $"A{row}"))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            markers.AddRange(MonthTextParser.Extract("8表A列", monthCell));
        }
        var issues = new List<ValidationIssue>();
        if (kinds.Count == 0)
            issues.Add(new("WORKBOOK_STRUCTURE_UNKNOWN", "工作簿缺少必需 sheet 或表头签名。",
                Path.GetFileName(path)));
        else if (kinds.Count > 1)
            issues.Add(new("WORKBOOK_KIND_AMBIGUOUS", $"工作簿匹配 {kinds.Count} 种表结构。", Path.GetFileName(path)));
        return new WorkbookInspection(path, kinds, markers.Distinct().ToArray(), issues);
    }

    private static bool IsTable1(OpenXmlWorkbookReader b) => b.HasSheet("分省统计表") &&
        HeaderText.Contains(b.GetText("分省统计表", "B2"), "省份") &&
        HeaderText.Contains(b.GetText("分省统计表", "C2"), "运营商") &&
        HeaderText.Contains(b.GetText("分省统计表", "D3"), "城域网出口IPv4和IPv6总流量") &&
        HeaderText.Contains(b.GetText("分省统计表", "E3"), "城域网出口IPv6总流量") &&
        HeaderText.Contains(b.GetText("分省统计表", "G3"), "移动核心网出口IPv4和IPv6总流量") &&
        HeaderText.Contains(b.GetText("分省统计表", "H3"), "移动核心网出口IPv6总流量");

    private static bool IsTable4(OpenXmlWorkbookReader b) =>
        HasTrafficSheet(b, "人网统计") && HasTrafficSheet(b, "物网统计");
    private static bool HasTrafficSheet(OpenXmlWorkbookReader b, string sheet) => b.HasSheet(sheet) &&
        HeaderText.Contains(b.GetText(sheet, "H4"), "prov_id") &&
        HeaderText.Contains(b.GetText(sheet, "R4"), "总计日均流量(PB)") &&
        HeaderText.Contains(b.GetText(sheet, "S4"), "IPV6日均流量(PB)");

    private static bool IsTable5(OpenXmlWorkbookReader b)
    {
        var idc = b.SheetNames.FirstOrDefault(name => HeaderText.Normalize(name) == "IDC汇总(客户)");
        return b.HasSheet("互联网专线汇总") && idc is not null &&
            HeaderText.Contains(b.GetText("互联网专线汇总", "M2"), "省") &&
            HeaderText.Contains(b.GetText("互联网专线汇总", "E2"), "总流量") &&
            HeaderText.Contains(b.GetText("互联网专线汇总", "F2"), "IPv6流量") &&
            HeaderText.Contains(b.GetText("互联网专线汇总", "K2"), "总流量") &&
            HeaderText.Contains(b.GetText("互联网专线汇总", "L2"), "IPv6流量") &&
            HeaderText.Contains(b.GetText(idc, "O2"), "省") &&
            HeaderText.Contains(b.GetText(idc, "F2"), "总流量") &&
            HeaderText.Contains(b.GetText(idc, "G2"), "V6日流量") &&
            HeaderText.Contains(b.GetText(idc, "M2"), "总流量") &&
            HeaderText.Contains(b.GetText(idc, "N2"), "V6日流量");
    }

    private static bool IsTable8(OpenXmlWorkbookReader b)
    {
        var sheet = b.SheetNames.FirstOrDefault(name => name == "省统计" || name.EndsWith("-省统计", StringComparison.Ordinal));
        return sheet is not null && HeaderText.Contains(b.GetText(sheet, "A2"), "月") &&
            HeaderText.Contains(b.GetText(sheet, "B2"), "省份") &&
            HeaderText.Contains(b.GetText(sheet, "D2"), "总流量") &&
            HeaderText.Contains(b.GetText(sheet, "J2"), "IPv6总流量");
    }
}
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~OpenXmlWorkbookInspectorTests
```

Expected: PASS，半角/全角 IDC 别名和带前缀/不带前缀省统计均能识别；混合结构被拒绝。

- [ ] **Step 5: 提交工作簿识别**

```bash
git add src/Ipv6ProvinceStatistics.Application/Abstractions/IWorkbookInspector.cs src/Ipv6ProvinceStatistics.Application/Models/WorkbookInspection.cs src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading tests/Ipv6ProvinceStatistics.IntegrationTests/Reading tests/Ipv6ProvinceStatistics.IntegrationTests/Fixtures
git commit -m "feat(excel): identify source workbooks" -m "Classify the four source workbook types by sheet signatures and collect full or partial month markers from filenames and cells."
```

### Task 8: 提取 1 表省份分组中的中国联通数据

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Application/Abstractions/ISourceWorkbookReader.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction/ExtractorSupport.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction/Table1Extractor.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlSourceWorkbookReader.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Extraction/Table1ExtractorTests.cs`

- [ ] **Step 1: 写省份分组、运营商选择和非法数值失败测试**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Extraction;

public sealed class Table1ExtractorTests
{
    [Fact]
    public async Task Reader_carries_province_to_china_unicom_row()
    {
        var path = CreateTable1("61.5");
        var result = await new OpenXmlSourceWorkbookReader().ReadAsync(path, SourceWorkbookKind.Table1, CancellationToken.None);
        var values = result.Values[new Province("北京")];
        Assert.Equal(101.25m, values[MetricKey.MetroTotal]);
        Assert.Equal(61.5m, values[MetricKey.MetroIpv6]);
        Assert.Equal(44.75m, values[MetricKey.MobileCoreTotal]);
        Assert.Equal(22.125m, values[MetricKey.MobileCoreIpv6]);
        Assert.Empty(result.Issues);
    }

    [Fact]
    public async Task Reader_reports_invalid_required_number_with_cell_address()
    {
        var result = await new OpenXmlSourceWorkbookReader().ReadAsync(
            CreateTable1("NULL"), SourceWorkbookKind.Table1, CancellationToken.None);
        Assert.Contains(result.Issues, x => x.Code == "VALUE_INVALID" && x.Cell == "E6");
        Assert.Empty(result.Values);
    }

    private static string CreateTable1(string metroIpv6)
    {
        var path = TempFiles.Next("1.xlsx");
        TestWorkbookBuilder.Create(path, new TestSheet("分省统计表",
        [
            TestCell.SharedText("B4", "北   京"), TestCell.SharedText("C4", "中国电信"),
            TestCell.SharedText("C5", "中国移动"), TestCell.SharedText("C6", "中国联通"),
            TestCell.Number("D6", "101.25"), TestCell.SharedText("E6", metroIpv6),
            TestCell.Number("G6", "44.75"), TestCell.Number("H6", "22.125")
        ]));
        return path;
    }
}
```

- [ ] **Step 2: 运行测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~Table1ExtractorTests
```

Expected: FAIL，缺少 `ISourceWorkbookReader` 和提取实现。

- [ ] **Step 3: 定义读取契约和公共提取帮助器**

```csharp
// ISourceWorkbookReader.cs
using Ipv6ProvinceStatistics.Application.Models;
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public interface ISourceWorkbookReader
{
    Task<SourceReadResult> ReadAsync(string path, SourceWorkbookKind kind, CancellationToken cancellationToken);
}
```

```csharp
// ExtractorSupport.cs
using System.Globalization;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class ExtractorSupport
{
    public static bool TryReadDecimal(OpenXmlWorkbookReader book, string file, string sheet,
        string address, Province province, List<ValidationIssue> issues, out decimal value)
    {
        var raw = book.GetText(sheet, address);
        if (decimal.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands,
            CultureInfo.InvariantCulture, out value)) return true;
        issues.Add(new("VALUE_INVALID", $"期望数值，实际为“{raw ?? "空值"}”。",
            Path.GetFileName(file), sheet, province.Name, address));
        return false;
    }

    public static void AddProvince(Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>> target,
        Province province, IReadOnlyDictionary<MetricKey, decimal> metrics, string file,
        string sheet, List<ValidationIssue> issues)
    {
        if (!target.TryAdd(province, metrics))
            issues.Add(new("PROVINCE_DUPLICATE", $"省份 {province.Name} 重复。",
                Path.GetFileName(file), sheet, province.Name));
    }
}
```

- [ ] **Step 4: 实现 1 表提取和 dispatcher**

```csharp
// Table1Extractor.cs
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table1Extractor
{
    public static SourceReadResult Extract(string path, OpenXmlWorkbookReader book, CancellationToken token)
    {
        const string sheet = "分省统计表";
        var output = new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>();
        var issues = new List<ValidationIssue>();
        Province? current = null;
        foreach (var row in book.GetPopulatedRows(sheet).Where(row => row >= 4))
        {
            token.ThrowIfCancellationRequested();
            if (ProvinceCatalog.TryResolve(book.GetText(sheet, $"B{row}"), out var resolved)) current = resolved;
            if (HeaderText.Normalize(book.GetText(sheet, $"C{row}")) != "中国联通") continue;
            if (current is null)
            {
                issues.Add(new("PROVINCE_GROUP_MISSING", "中国联通行之前没有可识别省份。",
                    Path.GetFileName(path), sheet, Cell: $"C{row}"));
                continue;
            }
            var before = issues.Count;
            var metrics = new Dictionary<MetricKey, decimal>();
            Read("D", MetricKey.MetroTotal); Read("E", MetricKey.MetroIpv6);
            Read("G", MetricKey.MobileCoreTotal); Read("H", MetricKey.MobileCoreIpv6);
            if (issues.Count == before) ExtractorSupport.AddProvince(output, current.Value, metrics, path, sheet, issues);

            void Read(string column, MetricKey key)
            {
                if (ExtractorSupport.TryReadDecimal(book, path, sheet, $"{column}{row}", current.Value,
                    issues, out var value)) metrics[key] = value;
            }
        }
        return new(SourceWorkbookKind.Table1, issues.Count == 0 ? output :
            new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>(), issues);
    }
}
```

```csharp
// OpenXmlSourceWorkbookReader.cs
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;
namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
public sealed class OpenXmlSourceWorkbookReader : ISourceWorkbookReader
{
    public Task<SourceReadResult> ReadAsync(string path, SourceWorkbookKind kind, CancellationToken token)
    {
        using var book = OpenXmlWorkbookReader.Open(path, false);
        var result = kind == SourceWorkbookKind.Table1
            ? Table1Extractor.Extract(path, book, token)
            : throw new NotSupportedException($"Extractor not implemented: {kind}");
        return Task.FromResult(result);
    }
}
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~Table1ExtractorTests
```

Expected: PASS；只读取中国联通行，非法 E6 精确报错且不返回部分省份值。

- [ ] **Step 5: 提交 1 表提取器**

```bash
git add src/Ipv6ProvinceStatistics.Application/Abstractions/ISourceWorkbookReader.cs src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlSourceWorkbookReader.cs tests/Ipv6ProvinceStatistics.IntegrationTests/Extraction
git commit -m "feat(excel): extract table 1 metrics" -m "Carry province groups to China Unicom rows and read the four fixed and mobile source metrics with cell-level validation."
```

### Task 9: 提取 4 表人网和物网数据

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction/Table4Extractor.cs`
- Modify: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlSourceWorkbookReader.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Extraction/Table4ExtractorTests.cs`

- [ ] **Step 1: 写 G/R/S 映射失败测试**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Extraction;

public sealed class Table4ExtractorTests
{
[Fact]
public async Task Reader_maps_human_and_iot_daily_pb_columns()
{
    var path = TempFiles.Next("4.xlsx");
    TestWorkbookBuilder.Create(path,
        new TestSheet("人网统计", [TestCell.SharedText("G5", "北京"),
            TestCell.Number("R5", "8.25"), TestCell.Number("S5", "5.5")]),
        new TestSheet("物网统计", [TestCell.SharedText("G5", "北京市"),
            TestCell.Number("R5", "1.125"), TestCell.Number("S5", "0.25")]));
    var result = await new OpenXmlSourceWorkbookReader().ReadAsync(path, SourceWorkbookKind.Table4, CancellationToken.None);
    var values = result.Values[new Province("北京")];
    Assert.Equal(8.25m, values[MetricKey.HumanTotal]);
    Assert.Equal(5.5m, values[MetricKey.HumanIpv6]);
    Assert.Equal(1.125m, values[MetricKey.IotTotal]);
    Assert.Equal(0.25m, values[MetricKey.IotIpv6]);
}

[Fact]
public async Task Reader_rejects_duplicate_province_rows()
{
    var path = TempFiles.Next("4-duplicate.xlsx");
    TestWorkbookBuilder.Create(path,
        new TestSheet("人网统计", [
            TestCell.SharedText("G5", "北京"), TestCell.Number("R5", "8"), TestCell.Number("S5", "5"),
            TestCell.SharedText("G6", "北京市"), TestCell.Number("R6", "9"), TestCell.Number("S6", "6")
        ]),
        new TestSheet("物网统计", [
            TestCell.SharedText("G5", "北京"), TestCell.Number("R5", "1"), TestCell.Number("S5", "0.2")
        ]));
    var result = await new OpenXmlSourceWorkbookReader().ReadAsync(path,
        SourceWorkbookKind.Table4, CancellationToken.None);
    Assert.Contains(result.Issues, issue => issue.Code == "PROVINCE_DUPLICATE" && issue.Province == "北京");
    Assert.Empty(result.Values);
}
}
```

- [ ] **Step 2: 运行测试并确认 dispatcher 不支持 Table4**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~Table4ExtractorTests
```

Expected: FAIL，抛出 `NotSupportedException: Extractor not implemented: Table4`。

- [ ] **Step 3: 实现两个 sheet 的合并提取**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table4Extractor
{
    public static SourceReadResult Extract(string path, OpenXmlWorkbookReader book, CancellationToken token)
    {
        var issues = new List<ValidationIssue>();
        var merged = new Dictionary<Province, Dictionary<MetricKey, decimal>>();
        ReadSheet("人网统计", MetricKey.HumanTotal, MetricKey.HumanIpv6);
        ReadSheet("物网统计", MetricKey.IotTotal, MetricKey.IotIpv6);
        var values = merged.ToDictionary(x => x.Key,
            x => (IReadOnlyDictionary<MetricKey, decimal>)x.Value);
        return new(SourceWorkbookKind.Table4, issues.Count == 0 ? values :
            new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>(), issues);

        void ReadSheet(string sheet, MetricKey totalKey, MetricKey ipv6Key)
        {
            var seen = new HashSet<Province>();
            foreach (var row in book.GetPopulatedRows(sheet).Where(row => row >= 5))
            {
                token.ThrowIfCancellationRequested();
                if (!ProvinceCatalog.TryResolve(book.GetText(sheet, $"G{row}"), out var province)) continue;
                if (!seen.Add(province))
                {
                    issues.Add(new("PROVINCE_DUPLICATE", $"{sheet} 的 {province.Name} 重复。",
                        Path.GetFileName(path), sheet, province.Name));
                    continue;
                }
                var before = issues.Count;
                if (!merged.TryGetValue(province, out var metrics)) merged[province] = metrics = [];
                if (ExtractorSupport.TryReadDecimal(book, path, sheet, $"R{row}", province, issues, out var total))
                    metrics[totalKey] = total;
                if (ExtractorSupport.TryReadDecimal(book, path, sheet, $"S{row}", province, issues, out var ipv6))
                    metrics[ipv6Key] = ipv6;
                if (issues.Count > before) merged.Remove(province);
            }
        }
    }
}
```

- [ ] **Step 4: 增加 dispatcher 分支并运行测试**

Replace the reader dispatch expression with:

```csharp
var result = kind switch
{
    SourceWorkbookKind.Table1 => Table1Extractor.Extract(path, book, token),
    SourceWorkbookKind.Table4 => Table4Extractor.Extract(path, book, token),
    _ => throw new NotSupportedException($"Extractor not implemented: {kind}")
};
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~Table4ExtractorTests
```

Expected: PASS，两个 sheet 的四项指标合并到同一北京记录。

- [ ] **Step 5: 提交 4 表提取器**

```bash
git add src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction/Table4Extractor.cs src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlSourceWorkbookReader.cs tests/Ipv6ProvinceStatistics.IntegrationTests/Extraction/Table4ExtractorTests.cs
git commit -m "feat(excel): extract table 4 metrics" -m "Merge the human-network and IoT daily PB values from the fixed G, R, and S columns by normalized province."
```

### Task 10: 提取 5 表互联网专线和 IDC 数据

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction/Table5Extractor.cs`
- Modify: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlSourceWorkbookReader.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Extraction/Table5ExtractorTests.cs`

- [ ] **Step 1: 写两种 IDC sheet 名、八项映射和非省级行测试**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Extraction;

public sealed class Table5ExtractorTests
{
[Theory]
[InlineData("IDC汇总 (客户)")]
[InlineData("IDC汇总（客户）")]
public async Task Reader_maps_internet_and_idc_columns_and_ignores_cloud_company(string idcSheet)
{
    var path = TempFiles.Next("5.xlsx");
    TestWorkbookBuilder.Create(path,
        new TestSheet("互联网专线汇总",
        [
            TestCell.SharedText("M6", "北京市"), TestCell.Number("E6", "100.1"),
            TestCell.Number("F6", "40.2"), TestCell.Number("K6", "60.3"),
            TestCell.Number("L6", "20.4")
        ]),
        new TestSheet(idcSheet,
        [
            TestCell.SharedText("O4", "北京市"), TestCell.Number("F4", "200.5"),
            TestCell.Number("G4", "80.6"), TestCell.Number("M4", "150.7"),
            TestCell.Number("N4", "70.8"), TestCell.SharedText("O11", "云公司"),
            TestCell.Number("F11", "999"), TestCell.Number("G11", "888"),
            TestCell.Number("M11", "777"), TestCell.Number("N11", "666")
        ]));
    var result = await new OpenXmlSourceWorkbookReader().ReadAsync(path, SourceWorkbookKind.Table5, CancellationToken.None);
    var values = result.Values[new Province("北京")];
    Assert.Equal(100.1m, values[MetricKey.InternetTotal]);
    Assert.Equal(40.2m, values[MetricKey.InternetIpv6]);
    Assert.Equal(60.3m, values[MetricKey.InternetOneGTotal]);
    Assert.Equal(20.4m, values[MetricKey.InternetOneGIpv6]);
    Assert.Equal(200.5m, values[MetricKey.IdcTotal]);
    Assert.Equal(80.6m, values[MetricKey.IdcIpv6]);
    Assert.Equal(150.7m, values[MetricKey.IdcTenGTotal]);
    Assert.Equal(70.8m, values[MetricKey.IdcTenGIpv6]);
    Assert.Single(result.Values);
}
}
```

- [ ] **Step 2: 运行测试并确认 dispatcher 不支持 Table5**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~Table5ExtractorTests
```

Expected: FAIL，抛出 `NotSupportedException: Extractor not implemented: Table5`。

- [ ] **Step 3: 实现两个汇总 sheet 的通用列映射**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table5Extractor
{
    public static SourceReadResult Extract(string path, OpenXmlWorkbookReader book, CancellationToken token)
    {
        var idcSheet = book.SheetNames.Single(name => HeaderText.Normalize(name) == "IDC汇总(客户)");
        var issues = new List<ValidationIssue>();
        var merged = new Dictionary<Province, Dictionary<MetricKey, decimal>>();
        ReadSheet("互联网专线汇总", "M", 3, new Dictionary<string, MetricKey>
        {
            ["E"] = MetricKey.InternetTotal, ["F"] = MetricKey.InternetIpv6,
            ["K"] = MetricKey.InternetOneGTotal, ["L"] = MetricKey.InternetOneGIpv6
        });
        ReadSheet(idcSheet, "O", 3, new Dictionary<string, MetricKey>
        {
            ["F"] = MetricKey.IdcTotal, ["G"] = MetricKey.IdcIpv6,
            ["M"] = MetricKey.IdcTenGTotal, ["N"] = MetricKey.IdcTenGIpv6
        });
        var values = merged.ToDictionary(x => x.Key,
            x => (IReadOnlyDictionary<MetricKey, decimal>)x.Value);
        return new(SourceWorkbookKind.Table5, issues.Count == 0 ? values :
            new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>(), issues);

        void ReadSheet(string sheet, string provinceColumn, uint firstRow,
            IReadOnlyDictionary<string, MetricKey> columns)
        {
            var seen = new HashSet<Province>();
            foreach (var row in book.GetPopulatedRows(sheet).Where(row => row >= firstRow))
            {
                token.ThrowIfCancellationRequested();
                if (!ProvinceCatalog.TryResolve(book.GetText(sheet, $"{provinceColumn}{row}"), out var province)) continue;
                if (!seen.Add(province))
                {
                    issues.Add(new("PROVINCE_DUPLICATE", $"{sheet} 的 {province.Name} 重复。",
                        Path.GetFileName(path), sheet, province.Name));
                    continue;
                }
                if (!merged.TryGetValue(province, out var metrics)) merged[province] = metrics = [];
                foreach (var mapping in columns)
                    if (ExtractorSupport.TryReadDecimal(book, path, sheet, $"{mapping.Key}{row}",
                        province, issues, out var value)) metrics[mapping.Value] = value;
            }
        }
    }
}
```

- [ ] **Step 4: 增加 dispatcher 分支并运行测试**

The switch must now contain:

```csharp
SourceWorkbookKind.Table1 => Table1Extractor.Extract(path, book, token),
SourceWorkbookKind.Table4 => Table4Extractor.Extract(path, book, token),
SourceWorkbookKind.Table5 => Table5Extractor.Extract(path, book, token),
_ => throw new NotSupportedException($"Extractor not implemented: {kind}")
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~Table5ExtractorTests
```

Expected: PASS；两种 IDC sheet 名均映射八项指标，`云公司` 不进入省份字典。

- [ ] **Step 5: 提交 5 表提取器**

```bash
git add src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction/Table5Extractor.cs src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlSourceWorkbookReader.cs tests/Ipv6ProvinceStatistics.IntegrationTests/Extraction/Table5ExtractorTests.cs
git commit -m "feat(excel): extract table 5 metrics" -m "Read the four internet-line and four IDC metrics from fixed columns while supporting both verified IDC sheet aliases."
```

### Task 11: 提取 8 表省统计 D/J 数据

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction/Table8Extractor.cs`
- Modify: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlSourceWorkbookReader.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Extraction/Table8ExtractorTests.cs`

- [ ] **Step 1: 写带前缀和无前缀省统计映射测试**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Extraction;

public sealed class Table8ExtractorTests
{
[Theory]
[InlineData("省统计")]
[InlineData("1-省统计")]
public async Task Reader_maps_broadband_D_and_J_columns(string sheet)
{
    var path = TempFiles.Next("8.xlsx");
    TestWorkbookBuilder.Create(path, new TestSheet(sheet,
    [
        TestCell.SharedText("B7", "北京市"), TestCell.Number("D7", "6.5E3"),
        TestCell.Number("J7", "5.625E2")
    ]));
    var result = await new OpenXmlSourceWorkbookReader().ReadAsync(path, SourceWorkbookKind.Table8, CancellationToken.None);
    var values = result.Values[new Province("北京")];
    Assert.Equal(6500m, values[MetricKey.BroadbandTotal]);
    Assert.Equal(562.5m, values[MetricKey.BroadbandIpv6]);
}
}
```

- [ ] **Step 2: 运行测试并确认 dispatcher 不支持 Table8**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~Table8ExtractorTests
```

Expected: FAIL，抛出 `NotSupportedException: Extractor not implemented: Table8`。

- [ ] **Step 3: 实现省统计提取器**

```csharp
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;

namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Extraction;

internal static class Table8Extractor
{
    public static SourceReadResult Extract(string path, OpenXmlWorkbookReader book, CancellationToken token)
    {
        var sheet = book.SheetNames.Single(name => name == "省统计" || name.EndsWith("-省统计", StringComparison.Ordinal));
        var issues = new List<ValidationIssue>();
        var output = new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>();
        foreach (var row in book.GetPopulatedRows(sheet).Where(row => row >= 4))
        {
            token.ThrowIfCancellationRequested();
            if (!ProvinceCatalog.TryResolve(book.GetText(sheet, $"B{row}"), out var province)) continue;
            var metrics = new Dictionary<MetricKey, decimal>();
            if (ExtractorSupport.TryReadDecimal(book, path, sheet, $"D{row}", province, issues, out var total))
                metrics[MetricKey.BroadbandTotal] = total;
            if (ExtractorSupport.TryReadDecimal(book, path, sheet, $"J{row}", province, issues, out var ipv6))
                metrics[MetricKey.BroadbandIpv6] = ipv6;
            if (metrics.Count == 2) ExtractorSupport.AddProvince(output, province, metrics, path, sheet, issues);
        }
        return new(SourceWorkbookKind.Table8, issues.Count == 0 ? output :
            new Dictionary<Province, IReadOnlyDictionary<MetricKey, decimal>>(), issues);
    }
}
```

- [ ] **Step 4: 完成 dispatcher 并运行全部提取测试**

The final switch expression is:

```csharp
var result = kind switch
{
    SourceWorkbookKind.Table1 => Table1Extractor.Extract(path, book, token),
    SourceWorkbookKind.Table4 => Table4Extractor.Extract(path, book, token),
    SourceWorkbookKind.Table5 => Table5Extractor.Extract(path, book, token),
    SourceWorkbookKind.Table8 => Table8Extractor.Extract(path, book, token),
    _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null)
};
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~Extraction
```

Expected: PASS，四种工作簿的全部提取测试通过。

- [ ] **Step 5: 提交 8 表提取器**

```bash
git add src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Extraction/Table8Extractor.cs src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Reading/OpenXmlSourceWorkbookReader.cs tests/Ipv6ProvinceStatistics.IntegrationTests/Extraction/Table8ExtractorTests.cs
git commit -m "feat(excel): extract table 8 metrics" -m "Read the province broadband total and IPv6 values from D and J for prefixed and unprefixed province-summary sheets."
```

### Task 12: 脱敏、固定并校验内置模板

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Templates/TemplateCellMap.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Templates/FormulaManifest.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Templates/TemplateSanitizer.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Templates/TemplateValidator.cs`
- Modify: `src/Ipv6ProvinceStatistics.TemplateTool/Program.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Fixtures/TemplateFixtures.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Templates/TemplateSanitizerTests.cs`
- Create: `assets/templates/Ipv6ReportTemplate.xlsx`
- Create: `assets/templates/Ipv6ReportTemplate.sha256`

- [ ] **Step 1: 写模板地址、脱敏和公式保持失败测试**

```csharp
using System.IO.Compression;
using System.Security.Cryptography;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

namespace Ipv6ProvinceStatistics.IntegrationTests.Templates;

public sealed class TemplateSanitizerTests
{
    [Fact]
    public void Sanitizer_clears_inputs_and_cached_values_but_preserves_formulas_and_styles()
    {
        var source = TemplateFixtures.CreatePopulatedTemplate();
        var output = TempFiles.Next("sanitized.xlsx");
        var styleBefore = HashZipPart(source, "xl/styles.xml");
        TemplateSanitizer.Sanitize(source, output);
        using var book = OpenXmlWorkbookReader.Open(output, false);
        Assert.All(TemplateCellMap.Inputs.Values, address => Assert.Null(book.GetText("Sheet1", address)));
        Assert.All(FormulaManifest.Formulas, pair =>
        {
            Assert.Equal(pair.Value, book.GetFormula("Sheet1", pair.Key));
            Assert.Null(book.GetText("Sheet1", pair.Key));
        });
        Assert.Equal(styleBefore, HashZipPart(output, "xl/styles.xml"));
        Assert.Empty(TemplateValidator.Validate(output));
    }

    private static string HashZipPart(string path, string part)
    {
        using var archive = ZipFile.OpenRead(path);
        using var stream = archive.GetEntry(part)!.Open();
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
```

Create `TemplateFixtures.cs`:

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;

namespace Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

public static class TemplateFixtures
{
    public static string CreatePopulatedTemplate()
    {
        var path = TempFiles.Next("populated-template.xlsx");
        var inputs = TemplateCellMap.Inputs.Values.Select(address => TestCell.Number(address, "999"));
        var formulas = FormulaManifest.Formulas.Select(pair => TestCell.Formula(pair.Key, pair.Value, "1"));
        TestWorkbookBuilder.Create(path, new TestSheet("Sheet1", inputs.Concat(formulas).ToArray()));

        using var document = SpreadsheetDocument.Open(path, true);
        var styles = document.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1U },
            new Fills(new Fill()) { Count = 1U },
            new Borders(new Border()) { Count = 1U },
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            new CellFormats(new CellFormat()) { Count = 1U });
        styles.Stylesheet.Save();
        return path;
    }
}
```

- [ ] **Step 2: 运行测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~TemplateSanitizerTests
```

Expected: FAIL，缺少模板清单、脱敏器和验证器。

- [ ] **Step 3: 固定 18 个输入地址和 37 个公式文本**

Create `TemplateCellMap.cs`:

```csharp
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
```

Create `FormulaManifest.cs`:

```csharp
namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
public static class FormulaManifest
{
    public static IReadOnlyDictionary<string, string> Formulas { get; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["C2"]="C17*3600*24/8/1024/1024", ["E2"]="E17*3600*24/8/1024/1024",
            ["C3"]="C2-C9*0.9", ["E3"]="E2-E9*0.9",
            ["C9"]="C18*3600*24/8/1024/1024", ["E9"]="E18*3600*24/8/1024/1024",
            ["D4"]="C27", ["E4"]="E27", ["D5"]="C20*3600*24/8/1024/1024",
            ["E5"]="E20*3600*24/8/1024/1024", ["D6"]="C21*3600*24/8/1024/1024",
            ["E6"]="E21*3600*24/8/1024/1024", ["D7"]="C22*3600*24/8/1024/1024",
            ["E7"]="E22*3600*24/8/1024/1024", ["D8"]="C23*3600*24/8/1024/1024",
            ["E8"]="E23*3600*24/8/1024/1024", ["D10"]="C25", ["E10"]="E25",
            ["D11"]="C26", ["E11"]="E26",
            ["F2"]="E2/C2", ["F3"]="E3/C3", ["F4"]="E4/D4", ["F5"]="E5/D5",
            ["F6"]="E6/D6", ["F7"]="E7/D7", ["F8"]="E8/D8", ["F9"]="E9/C9",
            ["F10"]="E10/D10", ["F11"]="E11/D11",
            ["G3"]="C3/(C3+C9)", ["G4"]="D4/(D4+D5+D7)", ["G5"]="D5/(D4+D5+D7)",
            ["G7"]="D7/(D4+D5+D7)", ["G9"]="C9/(C3+C9)",
            ["G10"]="G9*(D10/(D10+D11))", ["G11"]="G9*(D11/(D10+D11))"
        };
}
```

- [ ] **Step 4: 实现脱敏器、验证器和 CLI**

```csharp
// TemplateValidator.cs
using System.Security.Cryptography;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
public static class TemplateValidator
{
    public static IReadOnlyList<ValidationIssue> Validate(string path, string? expectedSha256 = null)
    {
        var issues = new List<ValidationIssue>();
        if (expectedSha256 is not null)
        {
            using var stream = File.OpenRead(path);
            var actual = Convert.ToHexString(SHA256.HashData(stream));
            if (!actual.Equals(expectedSha256.Trim(), StringComparison.OrdinalIgnoreCase))
                issues.Add(new("TEMPLATE_HASH", "内置模板哈希不匹配。"));
        }
        using var book = OpenXmlWorkbookReader.Open(path, false);
        if (!book.HasSheet("Sheet1")) return [new("TEMPLATE_SHEET", "内置模板缺少 Sheet1。")];
        foreach (var address in TemplateCellMap.Inputs.Values)
        {
            if (book.FindCell("Sheet1", address) is null)
                issues.Add(new("TEMPLATE_INPUT_MISSING", $"模板缺少输入单元格 {address}。", Cell: address));
            else if (book.GetFormula("Sheet1", address) is not null)
                issues.Add(new("TEMPLATE_INPUT_FORMULA", $"输入单元格 {address} 不应包含公式。", Cell: address));
        }
        foreach (var expected in FormulaManifest.Formulas)
            if (!string.Equals(book.GetFormula("Sheet1", expected.Key), expected.Value, StringComparison.Ordinal))
                issues.Add(new("TEMPLATE_FORMULA", $"公式 {expected.Key} 与版本清单不一致。", Cell: expected.Key));
        return issues;
    }
}
```

```csharp
// TemplateSanitizer.cs
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
public static class TemplateSanitizer
{
    public static void Sanitize(string input, string output)
    {
        File.Copy(input, output, true);
        using (var document = SpreadsheetDocument.Open(output, true))
        {
            var workbook = document.WorkbookPart ?? throw new InvalidDataException("Workbook part missing.");
            var sheet = workbook.Workbook.Sheets!.Elements<Sheet>().Single(x => x.Name == "Sheet1");
            var part = (WorksheetPart)workbook.GetPartById(sheet.Id!.Value!);
            foreach (var address in TemplateCellMap.Inputs.Values.Concat(FormulaManifest.Formulas.Keys))
            {
                var cell = part.Worksheet.Descendants<Cell>().Single(x =>
                    string.Equals(x.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
                cell.CellValue = null;
                if (TemplateCellMap.Inputs.Values.Contains(address, StringComparer.OrdinalIgnoreCase))
                    cell.DataType = null;
            }
            part.Worksheet.Save();
        }
        var issues = TemplateValidator.Validate(output);
        if (issues.Count > 0) throw new InvalidDataException(string.Join(Environment.NewLine, issues.Select(x => x.Message)));
    }
}
```

Replace `TemplateTool/Program.cs` with:

```csharp
using System.Security.Cryptography;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
if (args.Length != 4 || args[0] != "sanitize")
{
    Console.Error.WriteLine("Usage: sanitize <input.xlsx> <output.xlsx> <output.sha256>");
    return 2;
}
Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(args[2]))!);
TemplateSanitizer.Sanitize(args[1], args[2]);
using var stream = File.OpenRead(args[2]);
await File.WriteAllTextAsync(args[3], Convert.ToHexString(SHA256.HashData(stream)) + Environment.NewLine);
return 0;
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~TemplateSanitizerTests
dotnet run --project src/Ipv6ProvinceStatistics.TemplateTool -- sanitize "5月统计数据/IPv6流量分类统计公式表 - 北京.xlsx" assets/templates/Ipv6ReportTemplate.xlsx assets/templates/Ipv6ReportTemplate.sha256
actual=$(shasum -a 256 assets/templates/Ipv6ReportTemplate.xlsx | awk '{print toupper($1)}')
expected=$(tr -d '\r\n' < assets/templates/Ipv6ReportTemplate.sha256)
test "$actual" = "$expected"
```

Expected: 测试 PASS；CLI 退出 0；模板文件与提交的 SHA-256 完全一致。`git status --short` 不得显示 `5月统计数据/` 中的文件。

- [ ] **Step 5: 提交脱敏模板和工具**

```bash
git add assets/templates src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Templates src/Ipv6ProvinceStatistics.TemplateTool tests/Ipv6ProvinceStatistics.IntegrationTests/Templates
git commit -m "feat(template): embed sanitized report template" -m "Lock the 18 input cells and 37 formula texts, strip prior business values, and generate a reproducible template hash."
```

### Task 13: 生成并回读验证单个省级统计表

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Application/Abstractions/ITemplateReportExporter.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Templates/TemplateResourceProvider.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Templates/TemplateReportExporter.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Templates/ReportWorkbookVerifier.cs`
- Modify: `src/Ipv6ProvinceStatistics.Infrastructure.OpenXml/Ipv6ProvinceStatistics.Infrastructure.OpenXml.csproj`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Templates/TemplateReportExporterTests.cs`

- [ ] **Step 1: 写 18 个输入、37 个缓存、公式保持和 OOXML 有效性失败测试**

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Templates;

public sealed class TemplateReportExporterTests
{
[Fact]
public async Task Exporter_writes_numeric_inputs_preserves_formulas_and_verifies_output()
{
    var output = TempFiles.Next("北京-2026年05月.xlsx");
    var template = TempFiles.Next("embedded-template.xlsx");
    var input = new ProvinceReportInput(Enum.GetValues<MetricKey>().ToDictionary(key => key, key => 100m));
    var exporter = new TemplateReportExporter(new TemplateResourceProvider());
    await File.WriteAllBytesAsync(template, new TemplateResourceProvider().ReadTemplate());
    var layoutBefore = LayoutFingerprint(template);
    Assert.Empty(await exporter.ValidateTemplateAsync(CancellationToken.None));
    var issues = await exporter.ExportAsync(new Province("北京"), new ReportMonth(2026, 5), input,
        output, CancellationToken.None);
    Assert.Empty(issues);
    using var book = OpenXmlWorkbookReader.Open(output, false);
    foreach (var mapping in TemplateCellMap.Inputs)
    {
        Assert.True(book.TryGetDecimal("Sheet1", mapping.Value, out var actual));
        Assert.Equal(input[mapping.Key], actual);
    }
    foreach (var formula in FormulaManifest.Formulas)
        Assert.Equal(formula.Value, book.GetFormula("Sheet1", formula.Key));
    Assert.Equal(37, FormulaManifest.Formulas.Count);
    Assert.Equal(layoutBefore, LayoutFingerprint(output));
}

private static string LayoutFingerprint(string path)
{
    using var document = SpreadsheetDocument.Open(path, false);
    var workbook = document.WorkbookPart!;
    var sheet = workbook.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == "Sheet1");
    var worksheet = ((WorksheetPart)workbook.GetPartById(sheet.Id!.Value!)).Worksheet;
    var rows = string.Join("|", worksheet.Descendants<Row>().Select(row =>
        $"{row.RowIndex}:{row.Height}:{row.CustomHeight}:{row.Hidden}"));
    return string.Join("\n", workbook.WorkbookStylesPart?.Stylesheet.OuterXml,
        worksheet.Elements<Columns>().SingleOrDefault()?.OuterXml,
        worksheet.Elements<MergeCells>().SingleOrDefault()?.OuterXml, rows);
}
}
```

- [ ] **Step 2: 运行测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~TemplateReportExporterTests
```

Expected: FAIL，缺少 exporter 接口、资源提供器和验证器。

- [ ] **Step 3: 嵌入模板并定义 exporter 契约**

Add to the OpenXml project file:

```xml
<ItemGroup>
  <EmbeddedResource Include="..\..\assets\templates\Ipv6ReportTemplate.xlsx"
                    LogicalName="Ipv6ProvinceStatistics.Template.xlsx" />
  <EmbeddedResource Include="..\..\assets\templates\Ipv6ReportTemplate.sha256"
                    LogicalName="Ipv6ProvinceStatistics.Template.sha256" />
</ItemGroup>
```

Create `ITemplateReportExporter.cs`:

```csharp
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
```

Create `TemplateResourceProvider.cs`:

```csharp
using System.Reflection;
namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
public sealed class TemplateResourceProvider
{
    private readonly Assembly _assembly = typeof(TemplateResourceProvider).Assembly;
    public byte[] ReadTemplate()
    {
        using var stream = _assembly.GetManifestResourceStream("Ipv6ProvinceStatistics.Template.xlsx")
            ?? throw new InvalidDataException("Embedded template missing.");
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }
    public string ReadExpectedSha256()
    {
        using var stream = _assembly.GetManifestResourceStream("Ipv6ProvinceStatistics.Template.sha256")
            ?? throw new InvalidDataException("Embedded template hash missing.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd().Trim();
    }
}
```

- [ ] **Step 4: 实现写入、公式缓存和回读验证**

Create `ReportWorkbookVerifier.cs`:

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
public static class ReportWorkbookVerifier
{
    public static IReadOnlyList<ValidationIssue> Verify(string path, ProvinceReportInput input,
        IReadOnlyDictionary<string, decimal> calculated)
    {
        var issues = new List<ValidationIssue>();
        using (var book = OpenXmlWorkbookReader.Open(path, false))
        {
            foreach (var mapping in TemplateCellMap.Inputs)
                if (!book.TryGetDecimal("Sheet1", mapping.Value, out var actual) || actual != input[mapping.Key])
                    issues.Add(new("OUTPUT_INPUT", $"输出单元格 {mapping.Value} 与输入不一致。", Cell: mapping.Value));
            foreach (var formula in FormulaManifest.Formulas)
            {
                if (book.GetFormula("Sheet1", formula.Key) != formula.Value)
                    issues.Add(new("OUTPUT_FORMULA", $"输出公式 {formula.Key} 已改变。", Cell: formula.Key));
                if (!book.TryGetDecimal("Sheet1", formula.Key, out var actual) ||
                    Math.Abs(actual - calculated[formula.Key]) > Math.Max(1m, Math.Abs(calculated[formula.Key])) * 0.000000000001m)
                    issues.Add(new("OUTPUT_CACHE", $"输出缓存 {formula.Key} 不正确。", Cell: formula.Key));
            }
        }
        using var document = SpreadsheetDocument.Open(path, false);
        foreach (var error in new OpenXmlValidator().Validate(document).Take(20))
            issues.Add(new("OUTPUT_OPENXML", error.Description));
        return issues;
    }
}
```

Create `TemplateReportExporter.cs`:

```csharp
using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
public sealed class TemplateReportExporter(TemplateResourceProvider resources) : ITemplateReportExporter
{
    public async Task<IReadOnlyList<ValidationIssue>> ValidateTemplateAsync(CancellationToken token)
    {
        var bytes = resources.ReadTemplate();
        var hash = Convert.ToHexString(SHA256.HashData(bytes));
        if (!hash.Equals(resources.ReadExpectedSha256(), StringComparison.OrdinalIgnoreCase))
            return [new("TEMPLATE_HASH", "内置模板哈希不匹配。")];
        var path = Path.Combine(Path.GetTempPath(), $"template-{Guid.NewGuid():N}.xlsx");
        try
        {
            await File.WriteAllBytesAsync(path, bytes, token);
            return TemplateValidator.Validate(path, resources.ReadExpectedSha256());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    public async Task<IReadOnlyList<ValidationIssue>> ExportAsync(Province province, ReportMonth month,
        ProvinceReportInput input, string outputPath, CancellationToken token)
    {
        var calculation = ReportCalculator.Calculate(input);
        if (calculation.Issues.Count > 0) return calculation.Issues;
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        await File.WriteAllBytesAsync(outputPath, resources.ReadTemplate(), token);
        using (var document = SpreadsheetDocument.Open(outputPath, true))
        {
            var workbook = document.WorkbookPart!;
            var sheet = workbook.Workbook.Sheets!.Elements<Sheet>().Single(x => x.Name == "Sheet1");
            var part = (WorksheetPart)workbook.GetPartById(sheet.Id!.Value!);
            foreach (var mapping in TemplateCellMap.Inputs)
                SetNumber(part, mapping.Value, input[mapping.Key]);
            foreach (var cached in calculation.Values)
                SetNumber(part, cached.Key, cached.Value, preserveFormula: true);
            workbook.Workbook.CalculationProperties ??= new CalculationProperties();
            workbook.Workbook.CalculationProperties.CalculationMode = CalculateModeValues.Auto;
            workbook.Workbook.CalculationProperties.ForceFullCalculation = true;
            workbook.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            part.Worksheet.Save();
            workbook.Workbook.Save();
        }
        token.ThrowIfCancellationRequested();
        return ReportWorkbookVerifier.Verify(outputPath, input, calculation.Values);
    }

    private static void SetNumber(WorksheetPart part, string address, decimal value, bool preserveFormula = false)
    {
        var cell = part.Worksheet.Descendants<Cell>().Single(x =>
            string.Equals(x.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
        if (!preserveFormula) cell.CellFormula = null;
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
    }
}
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~TemplateReportExporterTests
```

Expected: PASS，输出可重新打开，18 个输入和 37 个缓存正确，公式文本不变，Open XML validator 无错误。

- [ ] **Step 5: 提交单省输出引擎**

```bash
git add src/Ipv6ProvinceStatistics.Application/Abstractions/ITemplateReportExporter.cs src/Ipv6ProvinceStatistics.Infrastructure.OpenXml tests/Ipv6ProvinceStatistics.IntegrationTests/Templates
git commit -m "feat(excel): export verified province report" -m "Populate the embedded template, preserve formulas, write cached results, and reject any workbook that fails structural readback."
```

### Task 14: 实现源文件快照、本地日志和同卷原子发布

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Application/Models/SourceFileSnapshot.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Models/TaskWorkspace.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Abstractions/ITaskWorkspaceManager.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Abstractions/IOutputTransaction.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Abstractions/IOperationLogger.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.FileSystem/AppPaths.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.FileSystem/TaskWorkspaceManager.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.FileSystem/OutputDirectoryNaming.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.FileSystem/OutputTransaction.cs`
- Create: `src/Ipv6ProvinceStatistics.Infrastructure.FileSystem/JsonOperationLogger.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/FileSystem/TaskWorkspaceManagerTests.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/FileSystem/OutputTransactionTests.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Logging/JsonOperationLoggerTests.cs`

- [ ] **Step 1: 写快照、编号目录、原子发布和脱敏日志失败测试**

Create `TaskWorkspaceManagerTests.cs`:

```csharp
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.FileSystem;

public sealed class TaskWorkspaceManagerTests
{
[Fact]
public async Task Workspace_copies_sources_and_records_sha256()
{
    var root = TempDirectories.Next();
    var source = Path.Combine(root, "1.xlsx");
    await File.WriteAllTextAsync(source, "source-bytes");
    var manager = new TaskWorkspaceManager(Path.Combine(root, "local-app-data"));
    var workspace = await manager.CreateAsync([source], CancellationToken.None);
    Assert.Single(workspace.Sources);
    Assert.True(File.Exists(workspace.Sources[0].SnapshotPath));
    Assert.Equal(64, workspace.Sources[0].Sha256.Length);
    await manager.CleanupAsync(workspace);
    Assert.False(Directory.Exists(workspace.Root));
}

[Fact]
public async Task Workspace_accepts_readonly_source_with_chinese_spaces_and_long_name()
{
    var root = TempDirectories.Next();
    var source = Path.Combine(root, "1-中文 空格-这是一个用于验证长文件名和绿色便携处理的合成源工作簿-2026年05月.xlsx");
    await File.WriteAllTextAsync(source, "readonly-source");
    File.SetAttributes(source, File.GetAttributes(source) | FileAttributes.ReadOnly);
    try
    {
        var manager = new TaskWorkspaceManager(Path.Combine(root, "local-app-data"));
        var workspace = await manager.CreateAsync([source], CancellationToken.None);
        Assert.Equal(Path.GetFileName(source), workspace.Sources[0].FileName);
        Assert.True(File.Exists(workspace.Sources[0].SnapshotPath));
        await manager.CleanupAsync(workspace);
    }
    finally { File.SetAttributes(source, FileAttributes.Normal); }
}

[Fact]
public async Task Workspace_snapshots_source_held_open_for_shared_read()
{
    var root = TempDirectories.Next();
    var source = Path.Combine(root, "1-open.xlsx");
    await File.WriteAllTextAsync(source, "open-source");
    await using var held = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    var manager = new TaskWorkspaceManager(Path.Combine(root, "local-app-data"));
    var workspace = await manager.CreateAsync([source], CancellationToken.None);
    Assert.True(File.Exists(workspace.Sources[0].SnapshotPath));
    await manager.CleanupAsync(workspace);
}
}
```

Create `OutputTransactionTests.cs`:

```csharp
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.FileSystem;

public sealed class OutputTransactionTests
{
[Fact]
public async Task Output_publishes_31_files_to_numbered_same_volume_directory()
{
    var parent = TempDirectories.Next();
    Directory.CreateDirectory(Path.Combine(parent, "2026年05月统计结果"));
    var transaction = new OutputTransaction();
    var staging = await transaction.CreateStagingAsync(parent, Guid.Parse("11111111-1111-1111-1111-111111111111"));
    foreach (var province in ProvinceCatalog.All)
        await File.WriteAllTextAsync(Path.Combine(staging, $"{province.Name}-2026年05月.xlsx"), province.Name);
    var published = await transaction.PublishAsync(staging, parent, new ReportMonth(2026, 5), CancellationToken.None);
    Assert.EndsWith("2026年05月统计结果 (2)", published, StringComparison.Ordinal);
    Assert.Equal(31, Directory.GetFiles(published, "*.xlsx").Length);
    Assert.False(new DirectoryInfo(published).Attributes.HasFlag(FileAttributes.Hidden));
}

[Fact]
public async Task Output_rejects_31_files_when_any_expected_name_is_missing()
{
    var parent = TempDirectories.Next();
    var transaction = new OutputTransaction();
    var staging = await transaction.CreateStagingAsync(parent, Guid.NewGuid());
    foreach (var province in ProvinceCatalog.All.Take(30))
        await File.WriteAllTextAsync(Path.Combine(staging, $"{province.Name}-2026年05月.xlsx"), province.Name);
    await File.WriteAllTextAsync(Path.Combine(staging, "错误命名-2026年05月.xlsx"), "wrong");
    await Assert.ThrowsAsync<InvalidDataException>(() => transaction.PublishAsync(staging, parent,
        new ReportMonth(2026, 5), CancellationToken.None));
    Assert.False(Directory.Exists(Path.Combine(parent, "2026年05月统计结果")));
    await transaction.CleanupAsync(staging);
}
}
```

Create `JsonOperationLoggerTests.cs`:

```csharp
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.Logging;

public sealed class JsonOperationLoggerTests
{
[Fact]
public async Task Logger_never_serializes_business_values_and_keeps_30_files()
{
    var root = TempDirectories.Next();
    var logger = new JsonOperationLogger(root);
    for (var index = 0; index < 31; index++)
        await logger.WriteAsync(new OperationLogEntry(Guid.NewGuid(), DateTimeOffset.UtcNow, "1.0.0",
            "Succeeded", 2026, 5, [new("1.xlsx", 100, DateTime.UnixEpoch, "ABC", "Table1")],
            "C:\\output", 31, 1000, []), CancellationToken.None);
    Assert.Equal(30, Directory.GetFiles(root, "*.json").Length);
    Assert.DoesNotContain("CellValue", await File.ReadAllTextAsync(Directory.GetFiles(root).First()));
}
}
```

- [ ] **Step 2: 运行三个测试类并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter "FullyQualifiedName~TaskWorkspaceManagerTests|FullyQualifiedName~OutputTransactionTests|FullyQualifiedName~JsonOperationLoggerTests"
```

Expected: FAIL，缺少文件系统模型、接口和实现。

- [ ] **Step 3: 定义文件系统契约和模型**

```csharp
// SourceFileSnapshot.cs
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record SourceFileSnapshot(string OriginalPath, string SnapshotPath, string FileName,
    long Length, DateTime LastWriteTimeUtc, string Sha256);
```

```csharp
// TaskWorkspace.cs
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record TaskWorkspace(Guid TaskId, string Root, IReadOnlyList<SourceFileSnapshot> Sources);
public sealed class SourceChangedException(string path) : IOException($"Source changed while copying: {path}");
```

```csharp
// ITaskWorkspaceManager.cs
using Ipv6ProvinceStatistics.Application.Models;
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public interface ITaskWorkspaceManager
{
    Task<TaskWorkspace> CreateAsync(IReadOnlyList<string> sourcePaths, CancellationToken cancellationToken);
    Task CleanupAsync(TaskWorkspace workspace);
}
```

```csharp
// IOutputTransaction.cs
using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public interface IOutputTransaction
{
    Task<string> CreateStagingAsync(string outputParent, Guid taskId);
    Task<string> PublishAsync(string stagingPath, string outputParent, ReportMonth month, CancellationToken token);
    Task CleanupAsync(string stagingPath);
}
```

```csharp
// IOperationLogger.cs
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public sealed record LogSource(string FileName, long Length, DateTime LastWriteTimeUtc, string Sha256,
    string? IdentifiedKind);
public sealed record OperationLogEntry(Guid TaskId, DateTimeOffset Timestamp, string Version, string Status,
    int? Year, int? Month, IReadOnlyList<LogSource> Sources, string? OutputDirectory,
    int OutputCount, long DurationMilliseconds,
    IReadOnlyList<string> Errors);
public interface IOperationLogger
{
    Task WriteAsync(OperationLogEntry entry, CancellationToken cancellationToken);
}
```

- [ ] **Step 4: 实现快照、发布和日志**

```csharp
// AppPaths.cs
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public static class AppPaths
{
    public static string Root => Path.Combine(Environment.GetFolderPath(
        Environment.SpecialFolder.LocalApplicationData), "IPv6ProvinceStatistics");
    public static string Workspaces => Path.Combine(Root, "Workspaces");
    public static string Logs => Path.Combine(Root, "Logs");
}
```

```csharp
// TaskWorkspaceManager.cs
using System.Security.Cryptography;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public sealed class TaskWorkspaceManager(string? root = null) : ITaskWorkspaceManager
{
    private readonly string _root = root ?? AppPaths.Workspaces;
    public async Task<TaskWorkspace> CreateAsync(IReadOnlyList<string> paths, CancellationToken token)
    {
        var id = Guid.NewGuid();
        var workspaceRoot = Path.Combine(_root, id.ToString("N"));
        Directory.CreateDirectory(workspaceRoot);
        var snapshots = new List<SourceFileSnapshot>();
        try
        {
            foreach (var source in paths)
            {
                token.ThrowIfCancellationRequested();
                var before = new FileInfo(source);
                var beforeLength = before.Length; var beforeWrite = before.LastWriteTimeUtc;
                var destination = Path.Combine(workspaceRoot, $"{snapshots.Count + 1}-{Path.GetFileName(source)}");
                await using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                await using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    await input.CopyToAsync(output, token);
                var after = new FileInfo(source);
                if (beforeLength != after.Length || beforeWrite != after.LastWriteTimeUtc)
                    throw new SourceChangedException(source);
                await using var copied = File.OpenRead(destination);
                snapshots.Add(new(source, destination, Path.GetFileName(source), beforeLength, beforeWrite,
                    Convert.ToHexString(await SHA256.HashDataAsync(copied, token))));
            }
            return new(id, workspaceRoot, snapshots);
        }
        catch { if (Directory.Exists(workspaceRoot)) Directory.Delete(workspaceRoot, true); throw; }
    }
    public Task CleanupAsync(TaskWorkspace workspace)
    {
        if (Directory.Exists(workspace.Root)) Directory.Delete(workspace.Root, true);
        return Task.CompletedTask;
    }
}
```

```csharp
// OutputDirectoryNaming.cs
using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public static class OutputDirectoryNaming
{
    public static string NextAvailable(string parent, ReportMonth month)
    {
        var basePath = Path.Combine(parent, month.FolderName);
        if (!Directory.Exists(basePath)) return basePath;
        for (var suffix = 2; suffix < int.MaxValue; suffix++)
        {
            var candidate = $"{basePath} ({suffix})";
            if (!Directory.Exists(candidate)) return candidate;
        }
        throw new IOException("无法分配输出目录名称。");
    }
}
```

```csharp
// OutputTransaction.cs
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public sealed class OutputTransaction : IOutputTransaction
{
    public Task<string> CreateStagingAsync(string parent, Guid id)
    {
        Directory.CreateDirectory(parent);
        var path = Path.Combine(parent, $".ipv6stats-{id:N}");
        Directory.CreateDirectory(path);
        if (OperatingSystem.IsWindows()) new DirectoryInfo(path).Attributes |= FileAttributes.Hidden;
        return Task.FromResult(path);
    }
    public Task<string> PublishAsync(string staging, string parent, ReportMonth month, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var expected = ProvinceCatalog.All.Select(province =>
            $"{province.Name}-{month.FileSuffix}.xlsx").ToHashSet(StringComparer.OrdinalIgnoreCase);
        var actual = Directory.GetFiles(staging, "*.xlsx").Select(path => Path.GetFileName(path)!)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!actual.SetEquals(expected))
            throw new InvalidDataException("Staging directory must contain the exact 31 province workbook names.");
        var target = OutputDirectoryNaming.NextAvailable(parent, month);
        Directory.Move(staging, target);
        if (OperatingSystem.IsWindows()) new DirectoryInfo(target).Attributes &= ~FileAttributes.Hidden;
        return Task.FromResult(target);
    }
    public Task CleanupAsync(string staging)
    {
        if (Directory.Exists(staging)) Directory.Delete(staging, true);
        return Task.CompletedTask;
    }
}
```

```csharp
// JsonOperationLogger.cs
using System.Text.Json;
using Ipv6ProvinceStatistics.Application.Abstractions;
namespace Ipv6ProvinceStatistics.Infrastructure.FileSystem;
public sealed class JsonOperationLogger(string? root = null) : IOperationLogger
{
    private readonly string _root = root ?? AppPaths.Logs;
    public async Task WriteAsync(OperationLogEntry entry, CancellationToken token)
    {
        Directory.CreateDirectory(_root);
        var path = Path.Combine(_root, $"{entry.Timestamp:yyyyMMdd-HHmmss}-{entry.TaskId:N}.json");
        await File.WriteAllTextAsync(path, JsonSerializer.Serialize(entry,
            new JsonSerializerOptions { WriteIndented = true }), token);
        foreach (var stale in Directory.GetFiles(_root, "*.json").OrderByDescending(File.GetCreationTimeUtc).Skip(30))
            File.Delete(stale);
    }
}
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter "FullyQualifiedName~TaskWorkspaceManagerTests|FullyQualifiedName~OutputTransactionTests|FullyQualifiedName~JsonOperationLoggerTests"
```

Expected: PASS；快照可清理，正式目录自动编号，发布后恰好 31 个文件，日志只保留 30 份。

- [ ] **Step 5: 提交文件事务和日志**

```bash
git add src/Ipv6ProvinceStatistics.Application/Abstractions src/Ipv6ProvinceStatistics.Application/Models src/Ipv6ProvinceStatistics.Infrastructure.FileSystem tests/Ipv6ProvinceStatistics.IntegrationTests/FileSystem tests/Ipv6ProvinceStatistics.IntegrationTests/Logging
git commit -m "feat(files): add atomic output transaction" -m "Snapshot mutable source files, stage outputs beside the destination for same-volume publication, and retain sanitized local operation logs."
```

### Task 15: 编排严格预检、31 省生成、取消与清理

**Files:**
- Create: `src/Ipv6ProvinceStatistics.Application/Models/PreparedBatch.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Models/ProcessingProgress.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Models/PreflightResult.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Models/GenerationResult.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Abstractions/IReportProcessingService.cs`
- Create: `src/Ipv6ProvinceStatistics.Application/Services/ReportProcessingService.cs`
- Create: `tests/Ipv6ProvinceStatistics.UnitTests/Services/ReportProcessingServiceTests.cs`
- Create: `tests/Ipv6ProvinceStatistics.UnitTests/Services/ProcessingHarness.cs`

- [ ] **Step 1: 写成功生成 31 份和无效预检零输出测试**

```csharp
using Ipv6ProvinceStatistics.UnitTests.Services;

namespace Ipv6ProvinceStatistics.UnitTests.Services;

public sealed class ReportProcessingServiceTests
{
[Fact]
public async Task Service_preflights_and_generates_exactly_31_files()
{
    using var harness = ProcessingHarness.Create(["1.xlsx", "4.xlsx", "5.xlsx", "8.xlsx"]);
    var preflight = await harness.Service.PreflightAsync(harness.Paths, null, null, CancellationToken.None);
    Assert.True(preflight.IsValid);
    var result = await harness.Service.GenerateAsync(preflight.Batch!, harness.OutputParent, null, CancellationToken.None);
    Assert.True(result.Succeeded);
    Assert.Equal(31, result.OutputCount);
    Assert.Equal(31, Directory.GetFiles(result.OutputDirectory!, "*.xlsx").Length);
    Assert.Contains("北京-2026年05月.xlsx", Directory.GetFiles(result.OutputDirectory!).Select(Path.GetFileName));
}

[Fact]
public async Task Service_rejects_duplicate_kind_and_cleans_workspace()
{
    using var harness = ProcessingHarness.Create(["1-a.xlsx", "1-b.xlsx", "4.xlsx", "5.xlsx"]);
    var result = await harness.Service.PreflightAsync(harness.Paths, null, null, CancellationToken.None);
    Assert.False(result.IsValid);
    Assert.Contains(result.Issues, x => x.Code == "SOURCE_KIND_COUNT");
    Assert.True(harness.Workspace.CleanupCalled);
    Assert.Empty(Directory.GetDirectories(harness.OutputParent));
}

[Fact]
public async Task Service_cancellation_removes_staging_and_workspace()
{
    using var harness = ProcessingHarness.Create(["1.xlsx", "4.xlsx", "5.xlsx", "8.xlsx"]);
    var preflight = await harness.Service.PreflightAsync(harness.Paths, null, null, CancellationToken.None);
    using var cancellation = new CancellationTokenSource();
    cancellation.Cancel();
    await Assert.ThrowsAsync<OperationCanceledException>(() => harness.Service.GenerateAsync(
        preflight.Batch!, harness.OutputParent, null, cancellation.Token));
    Assert.True(harness.Workspace.CleanupCalled);
    Assert.Empty(Directory.GetDirectories(harness.OutputParent));
}
}
```

Create `ProcessingHarness.cs`:

```csharp
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Application.Services;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.UnitTests.Services;

public sealed class ProcessingHarness : IDisposable
{
    private readonly string _root;
    public IReadOnlyList<string> Paths { get; }
    public string OutputParent { get; }
    public StubWorkspaceManager Workspace { get; }
    public IReportProcessingService Service { get; }

    private ProcessingHarness(IReadOnlyList<string> fileNames)
    {
        _root = Path.Combine(Path.GetTempPath(), "Ipv6ProvinceStatistics.UnitTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_root);
        OutputParent = Path.Combine(_root, "output");
        Directory.CreateDirectory(OutputParent);
        Paths = fileNames.Select(name => Path.Combine(_root, name)).ToArray();
        foreach (var path in Paths) File.WriteAllText(path, "fixture");
        Workspace = new StubWorkspaceManager(Path.Combine(_root, "workspace"));
        Service = new ReportProcessingService(Workspace, new StubInspector(), new StubReader(),
            new StubExporter(), new StubOutput(), new StubLogger(), "test");
    }

    public static ProcessingHarness Create(IReadOnlyList<string> fileNames) => new(fileNames);
    public void Dispose() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }

    public sealed class StubWorkspaceManager(string workspaceRoot) : ITaskWorkspaceManager
    {
        public bool CleanupCalled { get; private set; }
        public Task<TaskWorkspace> CreateAsync(IReadOnlyList<string> sourcePaths, CancellationToken token)
        {
            Directory.CreateDirectory(workspaceRoot);
            var sources = sourcePaths.Select(path =>
            {
                var snapshot = Path.Combine(workspaceRoot, Path.GetFileName(path));
                File.Copy(path, snapshot, true);
                return new SourceFileSnapshot(path, snapshot, Path.GetFileName(path),
                    new FileInfo(path).Length, File.GetLastWriteTimeUtc(path), "TEST-HASH");
            }).ToArray();
            return Task.FromResult(new TaskWorkspace(Guid.NewGuid(), workspaceRoot, sources));
        }
        public Task CleanupAsync(TaskWorkspace workspace)
        {
            CleanupCalled = true;
            if (Directory.Exists(workspace.Root)) Directory.Delete(workspace.Root, true);
            return Task.CompletedTask;
        }
    }

    private sealed class StubInspector : IWorkbookInspector
    {
        public Task<WorkbookInspection> InspectAsync(string path, CancellationToken token)
        {
            var kind = Path.GetFileName(path)[0] switch
            {
                '1' => SourceWorkbookKind.Table1, '4' => SourceWorkbookKind.Table4,
                '5' => SourceWorkbookKind.Table5, '8' => SourceWorkbookKind.Table8,
                _ => throw new InvalidDataException("Unknown test source.")
            };
            return Task.FromResult(new WorkbookInspection(path, [kind],
                [new MonthMarker(2026, 5, Path.GetFileName(path))], []));
        }
    }

    private sealed class StubReader : ISourceWorkbookReader
    {
        public Task<SourceReadResult> ReadAsync(string path, SourceWorkbookKind kind, CancellationToken token)
        {
            var values = ProvinceCatalog.All.ToDictionary(
                province => province,
                province => (IReadOnlyDictionary<MetricKey, decimal>)ProvinceDataAssembler
                    .ExpectedMetricsByKind[kind].ToDictionary(metric => metric, _ => 100m));
            return Task.FromResult(new SourceReadResult(kind, values, []));
        }
    }

    private sealed class StubExporter : ITemplateReportExporter
    {
        public Task<IReadOnlyList<ValidationIssue>> ValidateTemplateAsync(CancellationToken token) =>
            Task.FromResult<IReadOnlyList<ValidationIssue>>([]);
        public async Task<IReadOnlyList<ValidationIssue>> ExportAsync(Province province, ReportMonth month,
            ProvinceReportInput input, string outputPath, CancellationToken token)
        {
            await File.WriteAllTextAsync(outputPath, province.Name, token);
            return [];
        }
    }

    private sealed class StubOutput : IOutputTransaction
    {
        public Task<string> CreateStagingAsync(string parent, Guid taskId)
        {
            var path = Path.Combine(parent, $".stage-{taskId:N}");
            Directory.CreateDirectory(path);
            return Task.FromResult(path);
        }
        public Task<string> PublishAsync(string staging, string parent, ReportMonth month, CancellationToken token)
        {
            if (Directory.GetFiles(staging, "*.xlsx").Length != 31)
                throw new InvalidDataException("Staging must contain 31 workbooks.");
            var destination = Path.Combine(parent, $"{month.FileSuffix}统计结果");
            Directory.Move(staging, destination);
            return Task.FromResult(destination);
        }
        public Task CleanupAsync(string path)
        {
            if (Directory.Exists(path)) Directory.Delete(path, true);
            return Task.CompletedTask;
        }
    }

    private sealed class StubLogger : IOperationLogger
    {
        public List<OperationLogEntry> Entries { get; } = [];
        public Task WriteAsync(OperationLogEntry entry, CancellationToken token)
        {
            Entries.Add(entry);
            return Task.CompletedTask;
        }
    }
}
```

- [ ] **Step 2: 运行测试并确认失败**

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~ReportProcessingServiceTests
```

Expected: FAIL，缺少处理服务、批次和结果类型。

- [ ] **Step 3: 定义状态、结果和处理接口**

```csharp
// PreparedBatch.cs
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record PreparedBatch(TaskWorkspace Workspace, ReportMonth Month,
    IReadOnlyDictionary<Province, ProvinceReportInput> Reports,
    IReadOnlyDictionary<SourceWorkbookKind, SourceFileSnapshot> IdentifiedSources);
```

```csharp
// ProcessingProgress.cs
namespace Ipv6ProvinceStatistics.Application.Models;
public enum ProcessingStage { Snapshotting, Identifying, Validating, Generating, Verifying, Publishing }
public sealed record ProcessingProgress(ProcessingStage Stage, int Completed, int Total, string Message);
```

```csharp
// PreflightResult.cs
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record PreflightResult(PreparedBatch? Batch, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsValid => Batch is not null && Issues.Count == 0;
}
```

```csharp
// GenerationResult.cs
using Ipv6ProvinceStatistics.Domain.Validation;
namespace Ipv6ProvinceStatistics.Application.Models;
public sealed record GenerationResult(bool Succeeded, string? OutputDirectory, int OutputCount,
    TimeSpan Duration, IReadOnlyList<ValidationIssue> Issues);
```

```csharp
// IReportProcessingService.cs
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Reporting;
namespace Ipv6ProvinceStatistics.Application.Abstractions;
public interface IReportProcessingService
{
    Task<PreflightResult> PreflightAsync(IReadOnlyList<string> paths, ReportMonth? selectedMonth,
        IProgress<ProcessingProgress>? progress, CancellationToken token);
    Task<GenerationResult> GenerateAsync(PreparedBatch batch, string outputParent,
        IProgress<ProcessingProgress>? progress, CancellationToken token);
    Task DiscardAsync(PreparedBatch batch);
}
```

- [ ] **Step 4: 实现处理状态机并运行测试**

```csharp
using System.Diagnostics;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.Application.Services;

public sealed class ReportProcessingService(
    ITaskWorkspaceManager workspaces, IWorkbookInspector inspector, ISourceWorkbookReader reader,
    ITemplateReportExporter exporter, IOutputTransaction output, IOperationLogger logger,
    string applicationVersion) : IReportProcessingService
{
    public async Task<PreflightResult> PreflightAsync(IReadOnlyList<string> paths, ReportMonth? selected,
        IProgress<ProcessingProgress>? progress, CancellationToken token)
    {
        if (paths.Count != 4 || paths.Distinct(StringComparer.OrdinalIgnoreCase).Count() != 4)
            return new(null, [new("SOURCE_COUNT", "必须选择四个不同的 XLSX 文件。")]);
        TaskWorkspace? workspace = null;
        var inspections = new List<WorkbookInspection>();
        MonthResolution? resolvedMonth = null;
        try
        {
            progress?.Report(new(ProcessingStage.Snapshotting, 0, 4, "正在创建源文件快照"));
            workspace = await workspaces.CreateAsync(paths, token);
            foreach (var source in workspace.Sources)
            {
                progress?.Report(new(ProcessingStage.Identifying, inspections.Count, 4, source.FileName));
                var inspection = await inspector.InspectAsync(source.SnapshotPath, token);
                inspections.Add(inspection with
                {
                    Issues = inspection.Issues.Select(issue => issue with { FileName = source.FileName }).ToArray()
                });
            }
            var issues = inspections.SelectMany(x => x.Issues).ToList();
            foreach (var kind in Enum.GetValues<SourceWorkbookKind>())
                if (inspections.Count(x => x.MatchingKinds.Count == 1 && x.MatchingKinds[0] == kind) != 1)
                    issues.Add(new("SOURCE_KIND_COUNT", $"{kind} 必须且只能识别出一个文件。"));
            resolvedMonth = MonthResolver.Resolve(inspections.SelectMany(x => x.MonthMarkers).ToArray(), selected);
            issues.AddRange(resolvedMonth.Issues);
            if (issues.Count == 0)
            {
                var reads = new List<SourceReadResult>();
                foreach (var inspection in inspections)
                {
                    var read = await reader.ReadAsync(inspection.Path, inspection.MatchingKinds.Single(), token);
                    var original = workspace.Sources.Single(source => source.SnapshotPath == inspection.Path);
                    reads.Add(read with
                    {
                        Issues = read.Issues.Select(issue => issue with { FileName = original.FileName }).ToArray()
                    });
                }
                var assembly = ProvinceDataAssembler.Assemble(reads);
                issues.AddRange(assembly.Issues);
                issues.AddRange(await exporter.ValidateTemplateAsync(token));
                foreach (var report in assembly.Reports.Values)
                    issues.AddRange(ReportCalculator.Calculate(report).Issues);
                if (issues.Count == 0)
                {
                    var identified = inspections.ToDictionary(
                        inspection => inspection.MatchingKinds.Single(),
                        inspection => workspace.Sources.Single(source => source.SnapshotPath == inspection.Path));
                    return new(new(workspace, resolvedMonth!.Month!.Value, assembly.Reports, identified), []);
                }
            }
            var recognized = BuildIdentified(inspections, workspace);
            await workspaces.CleanupAsync(workspace);
            await LogAsync(workspace, "PreflightFailed", resolvedMonth?.Month, recognized,
                null, 0, TimeSpan.Zero, issues, token);
            return new(null, issues);
        }
        catch (OperationCanceledException)
        {
            if (workspace is not null) await workspaces.CleanupAsync(workspace);
            throw;
        }
        catch (Exception exception)
        {
            var issue = new ValidationIssue("INTERNAL_ERROR", exception.Message);
            if (workspace is not null)
            {
                await LogAsync(workspace, "PreflightFailed", resolvedMonth?.Month,
                    BuildIdentified(inspections, workspace), null, 0, TimeSpan.Zero,
                    [issue], CancellationToken.None);
                await workspaces.CleanupAsync(workspace);
            }
            return new(null, [issue]);
        }
    }

    public async Task<GenerationResult> GenerateAsync(PreparedBatch batch, string parent,
        IProgress<ProcessingProgress>? progress, CancellationToken token)
    {
        var watch = Stopwatch.StartNew();
        string? staging = null;
        var issues = new List<ValidationIssue>();
        try
        {
            staging = await output.CreateStagingAsync(parent, batch.Workspace.TaskId);
            for (var index = 0; index < ProvinceCatalog.All.Count; index++)
            {
                token.ThrowIfCancellationRequested();
                var province = ProvinceCatalog.All[index];
                progress?.Report(new(ProcessingStage.Generating, index, 31, $"正在生成 {province.Name}"));
                var path = Path.Combine(staging, $"{province.Name}-{batch.Month.FileSuffix}.xlsx");
                issues.AddRange(await exporter.ExportAsync(province, batch.Month, batch.Reports[province], path, token));
                if (issues.Count > 0) break;
            }
            if (issues.Count > 0)
            {
                await output.CleanupAsync(staging);
                await workspaces.CleanupAsync(batch.Workspace);
                await LogAsync(batch.Workspace, "Failed", batch.Month, batch.IdentifiedSources,
                    null, 0, watch.Elapsed, issues, token);
                return new(false, null, 0, watch.Elapsed, issues);
            }
            progress?.Report(new(ProcessingStage.Publishing, 31, 31, "正在发布完整结果"));
            var published = await output.PublishAsync(staging, parent, batch.Month, token);
            staging = null;
            await workspaces.CleanupAsync(batch.Workspace);
            await LogAsync(batch.Workspace, "Succeeded", batch.Month, batch.IdentifiedSources,
                published, 31, watch.Elapsed, [], token);
            return new(true, published, 31, watch.Elapsed, []);
        }
        catch (OperationCanceledException)
        {
            if (staging is not null) await output.CleanupAsync(staging);
            await workspaces.CleanupAsync(batch.Workspace);
            await LogAsync(batch.Workspace, "Canceled", batch.Month, batch.IdentifiedSources,
                null, 0, watch.Elapsed, [], CancellationToken.None);
            throw;
        }
        catch (Exception exception)
        {
            if (staging is not null) await output.CleanupAsync(staging);
            await workspaces.CleanupAsync(batch.Workspace);
            issues.Add(new("INTERNAL_ERROR", exception.Message));
            await LogAsync(batch.Workspace, "Failed", batch.Month, batch.IdentifiedSources,
                null, 0, watch.Elapsed, issues, CancellationToken.None);
            return new(false, null, 0, watch.Elapsed, issues);
        }
    }

    public Task DiscardAsync(PreparedBatch batch) => workspaces.CleanupAsync(batch.Workspace);

    private static IReadOnlyDictionary<SourceWorkbookKind, SourceFileSnapshot> BuildIdentified(
        IReadOnlyList<WorkbookInspection> inspections, TaskWorkspace workspace) => inspections
        .Where(inspection => inspection.MatchingKinds.Count == 1)
        .Select(inspection => new
        {
            Kind = inspection.MatchingKinds[0],
            Source = workspace.Sources.Single(source => source.SnapshotPath == inspection.Path)
        })
        .GroupBy(item => item.Kind)
        .Where(group => group.Count() == 1)
        .ToDictionary(group => group.Key, group => group.Single().Source);

    private Task LogAsync(TaskWorkspace workspace, string status, ReportMonth? month,
        IReadOnlyDictionary<SourceWorkbookKind, SourceFileSnapshot> identified,
        string? directory, int count, TimeSpan duration, IReadOnlyList<ValidationIssue> issues,
        CancellationToken token)
    {
        var kindsByPath = identified.ToDictionary(pair => pair.Value.SnapshotPath,
            pair => pair.Key.ToString(), StringComparer.OrdinalIgnoreCase);
        var sources = workspace.Sources.Select(source => new LogSource(source.FileName, source.Length,
            source.LastWriteTimeUtc, source.Sha256,
            kindsByPath.TryGetValue(source.SnapshotPath, out var kind) ? kind : null)).ToArray();
        return logger.WriteAsync(new(workspace.TaskId, DateTimeOffset.UtcNow, applicationVersion, status,
            month?.Year, month?.Month, sources, directory, count, (long)duration.TotalMilliseconds,
            issues.Select(issue => $"{issue.Code}: {issue.Message}").ToArray()), token);
    }
}
```

Run:

```bash
dotnet test tests/Ipv6ProvinceStatistics.UnitTests --filter FullyQualifiedName~ReportProcessingServiceTests
```

Expected: PASS；成功路径产生 31 个文件，无效预检清理 workspace 且正式输出为 0。

- [ ] **Step 5: 提交应用状态机**

```bash
git add src/Ipv6ProvinceStatistics.Application tests/Ipv6ProvinceStatistics.UnitTests/Services/ReportProcessingServiceTests.cs
git commit -m "feat(application): orchestrate monthly generation" -m "Run snapshot, identification, month resolution, extraction, template validation, 31-file generation, cleanup, logging, and atomic publication as one workflow."
```

### Task 16: 以测试驱动实现单页工作台 ViewModel

**Files:**
- Create: `src/Ipv6ProvinceStatistics.App/ViewModels/ObservableObject.cs`
- Create: `src/Ipv6ProvinceStatistics.App/ViewModels/AsyncRelayCommand.cs`
- Create: `src/Ipv6ProvinceStatistics.App/ViewModels/FileCardViewModel.cs`
- Create: `src/Ipv6ProvinceStatistics.App/ViewModels/MainWindowViewModel.cs`
- Create: `src/Ipv6ProvinceStatistics.App/Services/DialogService.cs`
- Create: `src/Ipv6ProvinceStatistics.App/Services/ShellService.cs`
- Create: `tests/Ipv6ProvinceStatistics.WindowsTests/ViewModels/MainWindowViewModelTests.cs`
- Create: `tests/Ipv6ProvinceStatistics.WindowsTests/ViewModels/ViewModelHarness.cs`

- [ ] **Step 1: 写初始、就绪、失败和成功状态测试**

```csharp
namespace Ipv6ProvinceStatistics.WindowsTests.ViewModels;

public sealed class MainWindowViewModelTests
{
[Fact]
public void Initial_state_disables_generation()
{
    using var harness = ViewModelHarness.Create();
    Assert.False(harness.ViewModel.GenerateCommand.CanExecute(null));
    Assert.Equal("请选择本月四张源表", harness.ViewModel.StatusText);
}

[Fact]
public async Task Valid_preflight_shows_four_identified_files_and_enables_generation()
{
    using var harness = ViewModelHarness.Create(valid: true);
    await harness.ViewModel.LoadFilesAsync(["1.xlsx", "4.xlsx", "5.xlsx", "8.xlsx"]);
    Assert.True(harness.ViewModel.GenerateCommand.CanExecute(null));
    Assert.Equal(4, harness.ViewModel.Files.Count);
    Assert.Equal(2026, harness.ViewModel.SelectedYear);
    Assert.Equal(5, harness.ViewModel.SelectedMonth);
    Assert.Equal("31/31 个省份数据完整", harness.ViewModel.StatusText);
}

[Fact]
public async Task Invalid_preflight_disables_generation_and_lists_errors()
{
    using var harness = ViewModelHarness.Create(valid: false);
    await harness.ViewModel.LoadFilesAsync(["1.xlsx", "4.xlsx", "5.xlsx", "8.xlsx"]);
    Assert.False(harness.ViewModel.GenerateCommand.CanExecute(null));
    Assert.True(harness.ViewModel.RevalidateMonthCommand.CanExecute(null));
    Assert.Contains(harness.ViewModel.Issues, text => text.Contains("IDC汇总", StringComparison.Ordinal));
}

[Fact]
public async Task Successful_generation_exposes_output_folder()
{
    using var harness = ViewModelHarness.Create(valid: true);
    await harness.ViewModel.LoadFilesAsync(["1.xlsx", "4.xlsx", "5.xlsx", "8.xlsx"]);
    harness.ViewModel.OutputParent = harness.OutputParent;
    await harness.ViewModel.GenerateCommand.ExecuteAsync();
    Assert.Equal("已生成 31 个统计表", harness.ViewModel.StatusText);
    Assert.Contains("2026年05月统计结果", harness.ViewModel.ResultDetail, StringComparison.Ordinal);
    Assert.Contains("1.0 秒", harness.ViewModel.ResultDetail, StringComparison.Ordinal);
    Assert.True(harness.ViewModel.OpenOutputCommand.CanExecute(null));
}
}
```

Create `ViewModelHarness.cs`:

```csharp
using Ipv6ProvinceStatistics.App.Services;
using Ipv6ProvinceStatistics.App.ViewModels;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.WindowsTests.ViewModels;

public sealed class ViewModelHarness : IDisposable
{
    private readonly string _root;
    public string OutputParent { get; }
    public MainWindowViewModel ViewModel { get; }

    private ViewModelHarness(bool valid)
    {
        _root = Path.Combine(Path.GetTempPath(), "Ipv6ProvinceStatistics.WindowsTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_root);
        OutputParent = Path.Combine(_root, "output");
        Directory.CreateDirectory(OutputParent);
        ViewModel = new MainWindowViewModel(new StubProcessingService(valid, _root),
            new StubDialogs(), new StubShell());
    }

    public static ViewModelHarness Create(bool valid = true) => new(valid);
    public void Dispose() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }

    private sealed class StubProcessingService(bool valid, string root) : IReportProcessingService
    {
        public Task<PreflightResult> PreflightAsync(IReadOnlyList<string> paths, ReportMonth? selected,
            IProgress<ProcessingProgress>? progress, CancellationToken token)
        {
            if (!valid)
                return Task.FromResult(new PreflightResult(null,
                    [new ValidationIssue("VALUE_INVALID", "期望数值，实际为“NULL”。",
                        "5.xlsx", "IDC汇总 (客户)", "西藏", "N31")]));
            var taskId = Guid.NewGuid();
            var snapshots = Enum.GetValues<SourceWorkbookKind>().Select((kind, index) =>
                new KeyValuePair<SourceWorkbookKind, SourceFileSnapshot>(kind,
                    new SourceFileSnapshot(paths[index], paths[index], Path.GetFileName(paths[index]),
                        1, DateTime.UnixEpoch, "TEST-HASH")))
                .ToDictionary(pair => pair.Key, pair => pair.Value);
            var workspace = new TaskWorkspace(taskId, Path.Combine(root, "workspace"), snapshots.Values.ToArray());
            var reports = ProvinceCatalog.All.ToDictionary(province => province,
                province => new ProvinceReportInput(Enum.GetValues<MetricKey>().ToDictionary(key => key, _ => 100m)));
            var month = selected ?? new ReportMonth(2026, 5);
            return Task.FromResult(new PreflightResult(new PreparedBatch(workspace, month, reports, snapshots), []));
        }

        public Task<GenerationResult> GenerateAsync(PreparedBatch batch, string outputParent,
            IProgress<ProcessingProgress>? progress, CancellationToken token)
        {
            var output = Path.Combine(outputParent, $"{batch.Month.FileSuffix}统计结果");
            Directory.CreateDirectory(output);
            return Task.FromResult(new GenerationResult(true, output, 31, TimeSpan.FromSeconds(1), []));
        }

        public Task DiscardAsync(PreparedBatch batch) => Task.CompletedTask;
    }

    private sealed class StubDialogs : IUiDialogService
    {
        public IReadOnlyList<string> SelectExcelFiles() => [];
        public string? SelectOutputFolder() => null;
    }

    private sealed class StubShell : IShellService
    {
        public void OpenFolder(string path) { }
    }
}
```

- [ ] **Step 2: 在 Windows 上运行测试并确认失败**

```powershell
dotnet test tests/Ipv6ProvinceStatistics.WindowsTests --filter FullyQualifiedName~MainWindowViewModelTests
```

Expected: FAIL，缺少 ViewModel、命令和 UI 服务接口。非 Windows 主机执行 `dotnet build tests/Ipv6ProvinceStatistics.WindowsTests` 验证编译，测试留到 Task 19 的真实 Windows 门槛。

- [ ] **Step 3: 实现可观察对象、异步命令和文件卡片**

```csharp
// ObservableObject.cs
using System.ComponentModel;
using System.Runtime.CompilerServices;
namespace Ipv6ProvinceStatistics.App.ViewModels;
public abstract class ObservableObject : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;
    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value; PropertyChanged?.Invoke(this, new(name)); return true;
    }
    protected void Raise([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new(name));
}
```

```csharp
// AsyncRelayCommand.cs
using System.Windows.Input;
namespace Ipv6ProvinceStatistics.App.ViewModels;
public sealed class AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null) : ICommand
{
    private bool _running;
    public event EventHandler? CanExecuteChanged;
    public bool CanExecute(object? parameter) => !_running && (canExecute?.Invoke() ?? true);
    public async void Execute(object? parameter) => await ExecuteAsync();
    public async Task ExecuteAsync()
    {
        if (!CanExecute(null)) return;
        _running = true; RaiseCanExecuteChanged();
        try { await execute(); }
        finally { _running = false; RaiseCanExecuteChanged(); }
    }
    public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}
```

Create the full-path file-card model while presenting only its base name:

```csharp
// FileCardViewModel.cs
using Ipv6ProvinceStatistics.Application.Models;
namespace Ipv6ProvinceStatistics.App.ViewModels;
public sealed record FileCardViewModel(SourceWorkbookKind Kind, string FullPath, bool IsValid)
{
    public string FileName => Path.GetFileName(FullPath);
}
```

Create `DialogService.cs`:

```csharp
namespace Ipv6ProvinceStatistics.App.Services;
public interface IUiDialogService
{
    IReadOnlyList<string> SelectExcelFiles();
    string? SelectOutputFolder();
}
```

Create `ShellService.cs`:

```csharp
namespace Ipv6ProvinceStatistics.App.Services;
public interface IShellService
{
    void OpenFolder(string path);
}
```

- [ ] **Step 4: 实现 MainWindowViewModel 并运行 Windows 测试**

```csharp
using System.Collections.ObjectModel;
using Ipv6ProvinceStatistics.App.Services;
using Ipv6ProvinceStatistics.Application.Abstractions;
using Ipv6ProvinceStatistics.Application.Models;
using Ipv6ProvinceStatistics.Domain.Reporting;
using Ipv6ProvinceStatistics.Domain.Validation;

namespace Ipv6ProvinceStatistics.App.ViewModels;

public sealed class MainWindowViewModel : ObservableObject
{
    private readonly IReportProcessingService _processing;
    private readonly IUiDialogService _dialogs;
    private readonly IShellService _shell;
    private PreparedBatch? _batch;
    private IReadOnlyList<string> _selectedPaths = [];
    private CancellationTokenSource? _cancellation;
    private string _statusText = "请选择本月四张源表";
    private string _resultDetail = string.Empty;
    private string _outputParent = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    private int _selectedYear = DateTime.Today.Year;
    private int _selectedMonth = DateTime.Today.Month;
    private bool _isBusy;
    private string? _lastOutput;

    public MainWindowViewModel(IReportProcessingService processing, IUiDialogService dialogs, IShellService shell)
    {
        _processing = processing; _dialogs = dialogs; _shell = shell;
        SelectFilesCommand = new(async () => await LoadFilesAsync(_dialogs.SelectExcelFiles()), () => !IsBusy);
        SelectOutputCommand = new(() => { OutputParent = _dialogs.SelectOutputFolder() ?? OutputParent; return Task.CompletedTask; }, () => !IsBusy);
        RevalidateMonthCommand = new(async () => await RevalidateAsync(), () => _selectedPaths.Count == 4 && !IsBusy);
        GenerateCommand = new(GenerateAsync, () => _batch is not null && !IsBusy && Directory.Exists(OutputParent));
        CancelCommand = new(() => { _cancellation?.Cancel(); return Task.CompletedTask; }, () => IsBusy);
        OpenOutputCommand = new(() => { if (_lastOutput is not null) _shell.OpenFolder(_lastOutput); return Task.CompletedTask; },
            () => _lastOutput is not null && Directory.Exists(_lastOutput));
    }

    public ObservableCollection<FileCardViewModel> Files { get; } = [];
    public ObservableCollection<string> Issues { get; } = [];
    public AsyncRelayCommand SelectFilesCommand { get; }
    public AsyncRelayCommand SelectOutputCommand { get; }
    public AsyncRelayCommand RevalidateMonthCommand { get; }
    public AsyncRelayCommand GenerateCommand { get; }
    public AsyncRelayCommand CancelCommand { get; }
    public AsyncRelayCommand OpenOutputCommand { get; }
    public string StatusText { get => _statusText; private set => Set(ref _statusText, value); }
    public string ResultDetail { get => _resultDetail; private set => Set(ref _resultDetail, value); }
    public string OutputParent { get => _outputParent; set { if (Set(ref _outputParent, value)) RefreshCommands(); } }
    public int SelectedYear { get => _selectedYear; set => Set(ref _selectedYear, value); }
    public int SelectedMonth { get => _selectedMonth; set => Set(ref _selectedMonth, value); }
    public bool IsBusy { get => _isBusy; private set { if (Set(ref _isBusy, value)) RefreshCommands(); } }

    public async Task LoadFilesAsync(IReadOnlyList<string> paths)
    {
        _selectedPaths = paths.ToArray();
        await PreflightAsync(_selectedPaths, null);
    }
    private async Task RevalidateAsync()
    {
        await PreflightAsync(_selectedPaths, new ReportMonth(SelectedYear, SelectedMonth));
    }

    private async Task PreflightAsync(IReadOnlyList<string> paths, ReportMonth? selected)
    {
        if (_batch is not null) { await _processing.DiscardAsync(_batch); _batch = null; }
        Files.Clear(); Issues.Clear(); _lastOutput = null; ResultDetail = string.Empty; IsBusy = true;
        _cancellation = new();
        try
        {
            var progress = new Progress<ProcessingProgress>(p => StatusText = p.Message);
            var result = await Task.Run(() => _processing.PreflightAsync(paths, selected,
                progress, _cancellation.Token), _cancellation.Token);
            foreach (var issue in result.Issues) Issues.Add(FormatIssue(issue));
            _batch = result.Batch;
            if (_batch is not null)
            {
                SelectedYear = _batch.Month.Year; SelectedMonth = _batch.Month.Month;
                foreach (var source in _batch.IdentifiedSources)
                    Files.Add(new(source.Key, source.Value.OriginalPath, true));
                StatusText = "31/31 个省份数据完整";
            }
            else StatusText = "输入数据存在问题，未生成结果";
        }
        catch (OperationCanceledException) { StatusText = "操作已取消"; }
        finally { IsBusy = false; RefreshCommands(); }
    }

    private async Task GenerateAsync()
    {
        if (_batch is null) return;
        IsBusy = true; _cancellation = new();
        try
        {
            var progress = new Progress<ProcessingProgress>(p => StatusText = p.Message);
            var result = await Task.Run(() => _processing.GenerateAsync(_batch, OutputParent,
                progress, _cancellation.Token), _cancellation.Token);
            Issues.Clear(); foreach (var issue in result.Issues) Issues.Add(FormatIssue(issue));
            if (result.Succeeded)
            {
                _lastOutput = result.OutputDirectory;
                ResultDetail = $"输出：{result.OutputDirectory} · 耗时 {result.Duration.TotalSeconds:F1} 秒";
                StatusText = "已生成 31 个统计表"; _batch = null;
            }
            else { StatusText = "生成失败，未留下正式结果"; _batch = null; }
        }
        catch (OperationCanceledException) { StatusText = "生成已取消，未留下正式结果"; _batch = null; }
        finally { IsBusy = false; RefreshCommands(); }
    }

    private void RefreshCommands()
    {
        SelectFilesCommand.RaiseCanExecuteChanged(); SelectOutputCommand.RaiseCanExecuteChanged();
        RevalidateMonthCommand.RaiseCanExecuteChanged(); GenerateCommand.RaiseCanExecuteChanged();
        CancelCommand.RaiseCanExecuteChanged(); OpenOutputCommand.RaiseCanExecuteChanged();
    }

    private static string FormatIssue(ValidationIssue issue)
    {
        var location = new[] { issue.FileName, issue.Sheet, issue.Province, issue.Cell }
            .Where(value => !string.IsNullOrWhiteSpace(value));
        var prefix = string.Join(" → ", location);
        return string.IsNullOrEmpty(prefix) ? issue.Message : $"{prefix}：{issue.Message}";
    }

    public async Task CloseAsync()
    {
        _cancellation?.Cancel();
        if (_batch is not null)
        {
            await _processing.DiscardAsync(_batch);
            _batch = null;
        }
    }
}
```

Run on Windows:

```powershell
dotnet test tests/Ipv6ProvinceStatistics.WindowsTests --filter FullyQualifiedName~MainWindowViewModelTests
```

Expected: PASS，初始、有效、无效和成功四种状态均符合测试。

- [ ] **Step 5: 提交 ViewModel 状态机**

```bash
git add src/Ipv6ProvinceStatistics.Application src/Ipv6ProvinceStatistics.App/ViewModels src/Ipv6ProvinceStatistics.App/Services tests/Ipv6ProvinceStatistics.WindowsTests
git commit -m "feat(ui): add workspace view model" -m "Expose file selection, month revalidation, strict preflight, generation, cancellation, issue display, and output-folder actions through testable commands."
```

### Task 17: 实现精美简洁的 WPF 单页工作台

**Files:**
- Modify: `src/Ipv6ProvinceStatistics.App/App.xaml`
- Modify: `src/Ipv6ProvinceStatistics.App/App.xaml.cs`
- Modify: `src/Ipv6ProvinceStatistics.App/MainWindow.xaml`
- Modify: `src/Ipv6ProvinceStatistics.App/MainWindow.xaml.cs`
- Create: `src/Ipv6ProvinceStatistics.App/Themes/Colors.xaml`
- Create: `src/Ipv6ProvinceStatistics.App/Themes/Controls.xaml`
- Modify: `src/Ipv6ProvinceStatistics.App/Services/DialogService.cs`
- Modify: `src/Ipv6ProvinceStatistics.App/Services/ShellService.cs`

- [ ] **Step 1: 构建 Windows 项目并记录当前最小界面基线**

```powershell
dotnet build src/Ipv6ProvinceStatistics.App -c Debug
dotnet run --project src/Ipv6ProvinceStatistics.App
```

Expected: 仅显示 Task 1 的居中文字窗口；截取基线图保存到本地临时目录，不提交。

- [ ] **Step 2: 创建颜色、按钮、卡片和输入框设计系统**

Create `Colors.xaml`:

```xml
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Color x:Key="PrimaryColor">#176BD8</Color><Color x:Key="PrimaryDarkColor">#1256AE</Color>
  <Color x:Key="SurfaceColor">#FFFFFF</Color><Color x:Key="CanvasColor">#F5F7FA</Color>
  <Color x:Key="TextColor">#203A58</Color><Color x:Key="MutedColor">#718097</Color>
  <Color x:Key="BorderColor">#DDE5EE</Color><Color x:Key="SuccessColor">#2F9E6D</Color>
  <Color x:Key="ErrorColor">#C25B3F</Color>
  <SolidColorBrush x:Key="PrimaryBrush" Color="{StaticResource PrimaryColor}" />
  <SolidColorBrush x:Key="PrimaryDarkBrush" Color="{StaticResource PrimaryDarkColor}" />
  <SolidColorBrush x:Key="SurfaceBrush" Color="{StaticResource SurfaceColor}" />
  <SolidColorBrush x:Key="CanvasBrush" Color="{StaticResource CanvasColor}" />
  <SolidColorBrush x:Key="TextBrush" Color="{StaticResource TextColor}" />
  <SolidColorBrush x:Key="MutedBrush" Color="{StaticResource MutedColor}" />
  <SolidColorBrush x:Key="BorderBrush" Color="{StaticResource BorderColor}" />
  <SolidColorBrush x:Key="SuccessBrush" Color="{StaticResource SuccessColor}" />
  <SolidColorBrush x:Key="ErrorBrush" Color="{StaticResource ErrorColor}" />
</ResourceDictionary>
```

Create `Controls.xaml`:

```xml
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Style TargetType="Window"><Setter Property="FontFamily" Value="Segoe UI" />
    <Setter Property="Background" Value="{StaticResource CanvasBrush}" />
    <Setter Property="Foreground" Value="{StaticResource TextBrush}" /></Style>
  <Style x:Key="CardStyle" TargetType="Border"><Setter Property="Background" Value="{StaticResource SurfaceBrush}" />
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}" /><Setter Property="BorderThickness" Value="1" />
    <Setter Property="CornerRadius" Value="12" /><Setter Property="Padding" Value="16" /></Style>
  <Style x:Key="PrimaryButtonStyle" TargetType="Button"><Setter Property="Foreground" Value="White" />
    <Setter Property="Background" Value="{StaticResource PrimaryBrush}" /><Setter Property="BorderThickness" Value="0" />
    <Setter Property="Padding" Value="18,11" /><Setter Property="FontWeight" Value="SemiBold" />
    <Setter Property="Cursor" Value="Hand" /><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
      <Border x:Name="Border" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" /></Border>
      <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Border" Property="Background" Value="{StaticResource PrimaryDarkBrush}" /></Trigger>
        <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Border" Property="Opacity" Value="0.45" /></Trigger></ControlTemplate.Triggers>
    </ControlTemplate></Setter.Value></Setter></Style>
  <Style x:Key="SecondaryButtonStyle" TargetType="Button" BasedOn="{StaticResource PrimaryButtonStyle}">
    <Setter Property="Foreground" Value="{StaticResource PrimaryBrush}" /><Setter Property="Background" Value="#EAF3FF" /></Style>
  <Style TargetType="TextBox"><Setter Property="Padding" Value="10,8" />
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}" /><Setter Property="BorderThickness" Value="1" />
    <Setter Property="Background" Value="White" /></Style>
</ResourceDictionary>
```

- [ ] **Step 3: 实现主窗口、拖放和本地服务**

Update `App.xaml` to remove `StartupUri` and merge both dictionaries:

```xml
<Application x:Class="Ipv6ProvinceStatistics.App.App"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Application.Resources><ResourceDictionary><ResourceDictionary.MergedDictionaries>
    <ResourceDictionary Source="Themes/Colors.xaml" />
    <ResourceDictionary Source="Themes/Controls.xaml" />
  </ResourceDictionary.MergedDictionaries></ResourceDictionary></Application.Resources>
</Application>
```

Replace `MainWindow.xaml` with:

```xml
<Window x:Class="Ipv6ProvinceStatistics.App.MainWindow"
 xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="IPv6 省级统计助手" Width="1060" Height="760" MinWidth="920" MinHeight="680"
 WindowStartupLocation="CenterScreen" AllowDrop="True" Drop="OnFilesDropped" Closed="OnClosed">
 <Grid><Grid.RowDefinitions><RowDefinition Height="64"/><RowDefinition/><RowDefinition Height="74"/></Grid.RowDefinitions>
  <Border Background="White" BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,0,0,1">
   <Grid Margin="24,0"><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
    <StackPanel Orientation="Horizontal" VerticalAlignment="Center"><Border Width="34" Height="34" CornerRadius="9" Background="{StaticResource PrimaryBrush}">
      <TextBlock Text="IP" Foreground="White" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
      <TextBlock Text="IPv6 省级统计助手" FontSize="17" FontWeight="SemiBold" Margin="12,0,0,0" VerticalAlignment="Center"/></StackPanel>
    <Border Grid.Column="1" Background="#EAF7F1" CornerRadius="16" Padding="12,6" VerticalAlignment="Center">
      <TextBlock Text="●  完全离线 · 数据仅在本机处理" Foreground="{StaticResource SuccessBrush}" FontSize="12"/></Border>
   </Grid>
  </Border>
  <Grid Grid.Row="1" Margin="24,20"><Grid.ColumnDefinitions><ColumnDefinition Width="1.55*"/><ColumnDefinition Width="16"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
   <Border Style="{StaticResource CardStyle}"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <Grid><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <StackPanel><TextBlock Text="导入本月数据" FontSize="16" FontWeight="SemiBold"/><TextBlock Text="可把四个 Excel 文件一起拖入窗口" Foreground="{StaticResource MutedBrush}" Margin="0,4,0,0"/></StackPanel>
      <Button Grid.Column="1" Content="选择四个文件" Command="{Binding SelectFilesCommand}" Style="{StaticResource SecondaryButtonStyle}"/></Grid>
    <ItemsControl Grid.Row="1" ItemsSource="{Binding Files}" Margin="0,16,0,16"><ItemsControl.ItemsPanel><ItemsPanelTemplate><UniformGrid Columns="2"/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
      <ItemsControl.ItemTemplate><DataTemplate><Border BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" CornerRadius="9" Margin="5" Padding="12">
        <StackPanel><TextBlock Text="{Binding Kind}" Foreground="{StaticResource PrimaryBrush}" FontWeight="Bold"/><TextBlock Text="{Binding FileName}" TextTrimming="CharacterEllipsis" ToolTip="{Binding FullPath}" Margin="0,6,0,0"/></StackPanel>
      </Border></DataTemplate></ItemsControl.ItemTemplate></ItemsControl>
    <Grid Grid.Row="2"><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="12"/><ColumnDefinition Width="2*"/></Grid.ColumnDefinitions>
      <StackPanel><TextBlock Text="统计年月" Foreground="{StaticResource MutedBrush}"/><StackPanel Orientation="Horizontal" Margin="0,6,0,0">
        <TextBox Width="72" Text="{Binding SelectedYear, UpdateSourceTrigger=PropertyChanged}"/><TextBlock Text="年" VerticalAlignment="Center" Margin="5,0"/>
        <TextBox Width="48" Text="{Binding SelectedMonth, UpdateSourceTrigger=PropertyChanged}"/><TextBlock Text="月" VerticalAlignment="Center" Margin="5,0"/>
        <Button Content="重新校验" Command="{Binding RevalidateMonthCommand}" Style="{StaticResource SecondaryButtonStyle}" Padding="10,7"/></StackPanel></StackPanel>
      <StackPanel Grid.Column="2"><TextBlock Text="输出位置" Foreground="{StaticResource MutedBrush}"/><DockPanel Margin="0,6,0,0"><Button DockPanel.Dock="Right" Content="选择" Command="{Binding SelectOutputCommand}" Style="{StaticResource SecondaryButtonStyle}" Padding="10,7"/>
        <TextBox Text="{Binding OutputParent}" IsReadOnly="True" Margin="0,0,8,0"/></DockPanel></StackPanel></Grid>
   </Grid></Border>
   <Border Grid.Column="2" Style="{StaticResource CardStyle}"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TextBlock Text="生成前检查" FontSize="16" FontWeight="SemiBold"/><TextBlock Grid.Row="1" Text="{Binding StatusText}" FontSize="18" FontWeight="SemiBold" Margin="0,24,0,12" TextWrapping="Wrap"/>
    <TextBlock Grid.Row="2" Text="{Binding ResultDetail}" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,0,0,12"/>
    <Expander Grid.Row="3" Header="查看错误详情" IsExpanded="False">
      <ListBox ItemsSource="{Binding Issues}" BorderThickness="0" Foreground="{StaticResource ErrorBrush}" Background="Transparent"/>
    </Expander>
    <StackPanel Grid.Row="4"><Button Content="生成 31 个统计表" Command="{Binding GenerateCommand}" Style="{StaticResource PrimaryButtonStyle}"/>
      <Button Content="取消当前操作" Command="{Binding CancelCommand}" Style="{StaticResource SecondaryButtonStyle}" Margin="0,8,0,0"/>
      <Button Content="打开输出文件夹" Command="{Binding OpenOutputCommand}" Style="{StaticResource SecondaryButtonStyle}" Margin="0,8,0,0"/></StackPanel>
   </Grid></Border>
  </Grid>
  <Border Grid.Row="2" Background="White" BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,1,0,0"><TextBlock Margin="24,0" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}"
    Text="流程：导入文件  →  自动识别  →  严格校验  →  生成并回读验证  →  一次性交付"/></Border>
 </Grid>
</Window>
```

Implement `MainWindow.xaml.cs`:

```csharp
using System.Windows;
using Ipv6ProvinceStatistics.App.ViewModels;
namespace Ipv6ProvinceStatistics.App;
public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel;
    public MainWindow(MainWindowViewModel viewModel)
    {
        InitializeComponent(); DataContext = _viewModel = viewModel;
    }
    private async void OnFilesDropped(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        var paths = ((string[])e.Data.GetData(DataFormats.FileDrop)).Where(path =>
            Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)).ToArray();
        await _viewModel.LoadFilesAsync(paths);
    }
    private async void OnClosed(object? sender, EventArgs e) => await _viewModel.CloseAsync();
}
```

Complete `DialogService.cs` below `IUiDialogService`:

```csharp
using Microsoft.Win32;
namespace Ipv6ProvinceStatistics.App.Services;
public sealed class DialogService : IUiDialogService
{
    public IReadOnlyList<string> SelectExcelFiles()
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", Multiselect = true };
        return dialog.ShowDialog() == true ? dialog.FileNames : [];
    }
    public string? SelectOutputFolder()
    {
        var dialog = new OpenFolderDialog { Title = "选择统计结果保存位置", Multiselect = false };
        return dialog.ShowDialog() == true ? dialog.FolderName : null;
    }
}
```

Complete `ShellService.cs` below `IShellService`:

```csharp
using System.Diagnostics;
namespace Ipv6ProvinceStatistics.App.Services;
public sealed class ShellService : IShellService
{
    public void OpenFolder(string path) => Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
}
```

- [ ] **Step 4: 在 App 启动时组合依赖并进行视觉检查**

Replace `App.xaml.cs` with:

```csharp
using System.Windows;
using Ipv6ProvinceStatistics.App.Services;
using Ipv6ProvinceStatistics.App.ViewModels;
using Ipv6ProvinceStatistics.Application.Services;
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
namespace Ipv6ProvinceStatistics.App;
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, args) =>
        {
            MessageBox.Show("程序遇到未预期错误。未生成任何正式结果，请查看本地日志。", "IPv6 省级统计助手",
                MessageBoxButton.OK, MessageBoxImage.Error); args.Handled = true;
        };
        var processing = new ReportProcessingService(new TaskWorkspaceManager(), new OpenXmlWorkbookInspector(),
            new OpenXmlSourceWorkbookReader(), new TemplateReportExporter(new TemplateResourceProvider()),
            new OutputTransaction(), new JsonOperationLogger(), typeof(App).Assembly.GetName().Version?.ToString() ?? "1.0.0");
        var viewModel = new MainWindowViewModel(processing, new DialogService(), new ShellService());
        new MainWindow(viewModel).Show();
    }
}
```

Run on Windows:

```powershell
dotnet build src/Ipv6ProvinceStatistics.App -c Debug
dotnet run --project src/Ipv6ProvinceStatistics.App
```

Expected: 窗口在 100% 和 150% 缩放下无裁切；四张文件卡、年月、输出位置、错误列表和操作按钮清晰；拖入非四个 `.xlsx` 时进入 Invalid，而不是崩溃。

- [ ] **Step 5: 提交 WPF 界面**

```bash
git add src/Ipv6ProvinceStatistics.App
git commit -m "feat(ui): build single-page workspace" -m "Create the polished offline WPF workspace with four-file drag and drop, month controls, output selection, validation issues, progress, cancellation, and result actions."
```

### Task 18: 用脱敏 31 省夹具跑通端到端生成

**Files:**
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/Fixtures/MonthlyFixtureBuilder.cs`
- Create: `tests/Ipv6ProvinceStatistics.IntegrationTests/EndToEnd/MonthlyGenerationTests.cs`

- [ ] **Step 1: 写真实组件端到端成功和非法值零输出测试**

```csharp
using Ipv6ProvinceStatistics.Application.Services;
using Ipv6ProvinceStatistics.Domain.Provinces;
using Ipv6ProvinceStatistics.Infrastructure.FileSystem;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Reading;
using Ipv6ProvinceStatistics.Infrastructure.OpenXml.Templates;
using Ipv6ProvinceStatistics.IntegrationTests.Fixtures;

namespace Ipv6ProvinceStatistics.IntegrationTests.EndToEnd;

public sealed class MonthlyGenerationTests
{
    [Fact]
    public async Task Real_pipeline_generates_and_verifies_31_named_workbooks()
    {
        var root = TempDirectories.Next();
        var fixture = MonthlyFixtureBuilder.Create(Path.Combine(root, "inputs"), false);
        var output = Path.Combine(root, "outputs");
        var service = CreateService(root);
        var preflight = await service.PreflightAsync(fixture.Paths, null, null, CancellationToken.None);
        Assert.True(preflight.IsValid, string.Join(Environment.NewLine, preflight.Issues.Select(x => x.Message)));
        var result = await service.GenerateAsync(preflight.Batch!, output, null, CancellationToken.None);
        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Issues.Select(x => x.Message)));
        Assert.Equal(31, Directory.GetFiles(result.OutputDirectory!, "*.xlsx").Length);
        foreach (var province in ProvinceCatalog.All)
        {
            var path = Path.Combine(result.OutputDirectory!, $"{province.Name}-2026年05月.xlsx");
            Assert.True(File.Exists(path));
            using var book = OpenXmlWorkbookReader.Open(path, false);
            foreach (var mapping in TemplateCellMap.Inputs)
            {
                Assert.True(book.TryGetDecimal("Sheet1", mapping.Value, out var actual));
                Assert.Equal(fixture.Expected[province][mapping.Key], actual);
            }
            var calculation = ReportCalculator.Calculate(fixture.Expected[province]);
            Assert.Empty(calculation.Issues);
            foreach (var cached in calculation.Values)
            {
                Assert.True(book.TryGetDecimal("Sheet1", cached.Key, out var actual));
                Assert.Equal(cached.Value, actual);
                Assert.Equal(FormulaManifest.Formulas[cached.Key], book.GetFormula("Sheet1", cached.Key));
            }
        }
    }

    [Fact]
    public async Task Invalid_required_value_produces_no_formal_output()
    {
        var root = TempDirectories.Next();
        var fixture = MonthlyFixtureBuilder.Create(Path.Combine(root, "inputs"), true);
        var output = Path.Combine(root, "outputs");
        var service = CreateService(root);
        var preflight = await service.PreflightAsync(fixture.Paths, null, null, CancellationToken.None);
        Assert.False(preflight.IsValid);
        Assert.Contains(preflight.Issues, issue => issue.Code == "VALUE_INVALID" && issue.Cell == "N3");
        Assert.False(Directory.Exists(output) && Directory.EnumerateFileSystemEntries(output).Any());
    }

    private static ReportProcessingService CreateService(string root) => new(
        new TaskWorkspaceManager(Path.Combine(root, "local", "workspaces")), new OpenXmlWorkbookInspector(),
        new OpenXmlSourceWorkbookReader(), new TemplateReportExporter(new TemplateResourceProvider()),
        new OutputTransaction(), new JsonOperationLogger(Path.Combine(root, "local", "logs")), "1.0.0-test");
}
```

- [ ] **Step 2: 运行测试并确认缺少完整月度夹具**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~MonthlyGenerationTests
```

Expected: FAIL，`MonthlyFixtureBuilder` 不存在。

- [ ] **Step 3: 创建与实际布局一致的 31 省合成四表**

```csharp
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
        var table1 = new List<TestCell>
        {
            TestCell.SharedText("A1", "2026年5月IPv6关键指标进展情况"),
            TestCell.SharedText("B2", "省份"), TestCell.SharedText("C2", "运营商"),
            TestCell.SharedText("D3", "城域网出口IPv4和IPv6总流量（Gbps）"),
            TestCell.SharedText("E3", "城域网出口IPv6总流量（Gbps）"),
            TestCell.SharedText("G3", "移动核心网出口IPv4和IPv6总流量（Gbps）"),
            TestCell.SharedText("H3", "移动核心网出口IPv6总流量（Gbps）")
        };
        var human = TrafficHeaders(); var iot = TrafficHeaders();
        var internet = new List<TestCell>
        {
            TestCell.SharedText("E2", "总流量（Gbps）"), TestCell.SharedText("F2", "IPv6流量（Gbps）"),
            TestCell.SharedText("K2", "总流量（Gbps）"), TestCell.SharedText("L2", "IPv6流量（Gbps）"),
            TestCell.SharedText("M2", "省")
        };
        var idc = new List<TestCell>
        {
            TestCell.SharedText("F2", "总流量（Gbps）"), TestCell.SharedText("G2", "V6日流量（Gbps）"),
            TestCell.SharedText("M2", "总流量（Gbps）"), TestCell.SharedText("N2", "V6日流量（Gbps）"),
            TestCell.SharedText("O2", "省")
        };
        var table8 = new List<TestCell>
        {
            TestCell.SharedText("A2", "月"), TestCell.SharedText("B2", "省份"),
            TestCell.SharedText("D2", "总流量(v4+v6，GB)"),
            TestCell.SharedText("J2", "IPv6总流量（GB）")
        };
        var expected = new Dictionary<Province, ProvinceReportInput>();
        for (var index = 0; index < ProvinceCatalog.All.Count; index++)
        {
            var province = ProvinceCatalog.All[index]; var value = 100m + index;
            var row1 = 4 + index * 4; var unionRow = row1 + 2; var row = 5 + index;
            var row5 = 3 + index; var row8 = 4 + index;
            table1.Add(TestCell.SharedText($"B{row1}", province.Name));
            table1.Add(TestCell.SharedText($"C{unionRow}", "中国联通"));
            table1.Add(TestCell.Number($"D{unionRow}", F(value + 1)));
            table1.Add(TestCell.Number($"E{unionRow}", F(value + 2)));
            table1.Add(TestCell.Number($"G{unionRow}", F(value + 3)));
            table1.Add(TestCell.Number($"H{unionRow}", F(value + 4)));
            AddTraffic(human, row, province.Name, value + 5, value + 6);
            AddTraffic(iot, row, province.Name, value + 7, value + 8);
            internet.Add(TestCell.SharedText($"M{row5}", province.Name));
            internet.Add(TestCell.Number($"E{row5}", F(value + 9)));
            internet.Add(TestCell.Number($"F{row5}", F(value + 10)));
            internet.Add(TestCell.Number($"K{row5}", F(value + 11)));
            internet.Add(TestCell.Number($"L{row5}", F(value + 12)));
            idc.Add(TestCell.SharedText($"O{row5}", province.Name));
            idc.Add(TestCell.Number($"F{row5}", F(value + 13)));
            idc.Add(TestCell.Number($"G{row5}", F(value + 14)));
            idc.Add(TestCell.Number($"M{row5}", F(value + 15)));
            idc.Add(TestCell.Number($"N{row5}", invalidIdcN && index == 0 ? "NULL" : F(value + 16)));
            table8.Add(TestCell.Number($"A{row8}", "202605"));
            table8.Add(TestCell.SharedText($"B{row8}", province.Name));
            table8.Add(TestCell.Number($"D{row8}", F(value + 17)));
            table8.Add(TestCell.Number($"J{row8}", F(value + 18)));
            expected[province] = new ProvinceReportInput(new Dictionary<MetricKey, decimal>
            {
                [MetricKey.MetroTotal]=value+1, [MetricKey.MetroIpv6]=value+2,
                [MetricKey.MobileCoreTotal]=value+3, [MetricKey.MobileCoreIpv6]=value+4,
                [MetricKey.HumanTotal]=value+5, [MetricKey.HumanIpv6]=value+6,
                [MetricKey.IotTotal]=value+7, [MetricKey.IotIpv6]=value+8,
                [MetricKey.InternetTotal]=value+9, [MetricKey.InternetIpv6]=value+10,
                [MetricKey.InternetOneGTotal]=value+11, [MetricKey.InternetOneGIpv6]=value+12,
                [MetricKey.IdcTotal]=value+13, [MetricKey.IdcIpv6]=value+14,
                [MetricKey.IdcTenGTotal]=value+15, [MetricKey.IdcTenGIpv6]=value+16,
                [MetricKey.BroadbandTotal]=value+17, [MetricKey.BroadbandIpv6]=value+18
            });
        }
        var paths = new[]
        {
            Path.Combine(root, "1-2026年5月.xlsx"), Path.Combine(root, "4-5月.xlsx"),
            Path.Combine(root, "5-202605.xlsx"), Path.Combine(root, "8-202605.xlsx")
        };
        TestWorkbookBuilder.Create(paths[0], new TestSheet("分省统计表", table1));
        TestWorkbookBuilder.Create(paths[1], new TestSheet("人网统计", human), new TestSheet("物网统计", iot));
        TestWorkbookBuilder.Create(paths[2], new TestSheet("互联网专线汇总", internet),
            new TestSheet("IDC汇总 (客户)", idc));
        TestWorkbookBuilder.Create(paths[3], new TestSheet("1-省统计", table8));
        return new(paths, expected);
    }

    private static List<TestCell> TrafficHeaders() =>
    [
        TestCell.SharedText("H4", "prov_id"), TestCell.SharedText("R4", "总计日均流量（PB）"),
        TestCell.SharedText("S4", "IPV6日均流量（PB）")
    ];
    private static void AddTraffic(List<TestCell> cells, int row, string province, decimal total, decimal ipv6)
    {
        cells.Add(TestCell.SharedText($"G{row}", province));
        cells.Add(TestCell.Number($"R{row}", F(total)));
        cells.Add(TestCell.Number($"S{row}", F(ipv6)));
    }
    private static string F(decimal value) => value.ToString(CultureInfo.InvariantCulture);
}
```

- [ ] **Step 4: 运行完整端到端与全部非 Windows 测试**

```bash
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests --filter FullyQualifiedName~MonthlyGenerationTests
dotnet test tests/Ipv6ProvinceStatistics.UnitTests
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests
```

Expected: 三条命令均 PASS；成功夹具生成 31 份，每份的 18 个输入、37 个公式文本和公式缓存都逐格回读相等；非法 `N3=NULL` 预检失败且正式输出为 0。

- [ ] **Step 5: 提交端到端夹具和测试**

```bash
git add tests/Ipv6ProvinceStatistics.IntegrationTests/Fixtures/MonthlyFixtureBuilder.cs tests/Ipv6ProvinceStatistics.IntegrationTests/EndToEnd/MonthlyGenerationTests.cs
git commit -m "test(e2e): verify 31-province export" -m "Exercise the real offline pipeline with four synthetic 31-province workbooks and prove invalid required data leaves no formal output."
```

### Task 19: 创建绿色发布脚本和用户文档

**Files:**
- Modify: `Directory.Build.props`
- Create: `build/Verify-OfflinePolicy.ps1`
- Create: `build/Verify-Portable.ps1`
- Create: `build/Publish-WinX64.ps1`
- Create: `README.md`
- Create: `docs/user-guide.md`
- Create: `**/packages.lock.json` through `dotnet restore`

- [ ] **Step 1: 写用户文档和明确的未签名提示**

Create `README.md`:

````markdown
# IPv6 省级统计助手

完全离线的 Windows 10/11 x64 桌面工具，从每月 1、4、5、8 四张 Excel 表生成 31 份省级统计表。

## 开发验证

```powershell
dotnet restore --locked-mode
dotnet test tests/Ipv6ProvinceStatistics.UnitTests -c Release --no-restore
dotnet test tests/Ipv6ProvinceStatistics.IntegrationTests -c Release --no-restore
dotnet test tests/Ipv6ProvinceStatistics.WindowsTests -c Release --no-restore
```

原始业务工作簿不得加入 Git。发布命令见 `build/Publish-WinX64.ps1`。
````

Create `docs/user-guide.md`:

```markdown
# IPv6 省级统计助手使用说明

1. 解压 `IPv6省级统计助手-v1.0.0-win-x64.zip`，不要直接在压缩包内运行。
2. 双击 `IPv6省级统计助手.exe`。
3. 点击“选择四个文件”或把当月 1、4、5、8 表一起拖入窗口。
4. 确认软件识别出的年月；存在明确月份冲突时必须更换错误源表。
5. 选择输出父目录，等待“31/31 个省份数据完整”。
6. 点击“生成 31 个统计表”。成功后可直接打开结果目录。

任何必填值、sheet、省份或月份异常都会阻止生成，软件不会用 0 替代缺失值。日志位于 `%LocalAppData%\IPv6ProvinceStatistics\Logs`，日志不含业务单元格数值。

若发布包未使用代码签名证书，Windows SmartScreen 或安全软件仍可能提示；请核对随包提供的 SHA-256，不要关闭安全软件绕过检查。
```

- [ ] **Step 2: 创建离线策略和绿色启动验证脚本**

Create `build/Verify-OfflinePolicy.ps1`:

```powershell
param([string]$Root = (Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference = 'Stop'
$patterns = @('HttpClient', 'System.Net.Http', 'WebRequest', 'TcpClient', 'UdpClient', 'System.Net.Sockets')
$matches = Get-ChildItem (Join-Path $Root 'src') -Recurse -Filter '*.cs' |
  Select-String -Pattern $patterns -SimpleMatch
if ($matches) { $matches | ForEach-Object { Write-Error $_.Line }; throw 'Offline policy violation.' }
Write-Host 'Offline policy source scan passed.'
```

Create `build/Verify-Portable.ps1`:

```powershell
param([Parameter(Mandatory=$true)][string]$PublishDirectory)
$ErrorActionPreference = 'Stop'
$exe = Join-Path $PublishDirectory 'IPv6省级统计助手.exe'
if (-not (Test-Path $exe)) { throw "Executable missing: $exe" }
if (Get-ChildItem $PublishDirectory -Recurse -Include '*.xlsx','*.ps1','*.cmd','*.bat') {
  throw 'Publish directory contains forbidden external template or script files.'
}
$process = Start-Process $exe -PassThru
try {
  $deadline = (Get-Date).AddSeconds(15)
  do { Start-Sleep -Milliseconds 250; $process.Refresh() }
  while (-not $process.HasExited -and $process.MainWindowHandle -eq 0 -and (Get-Date) -lt $deadline)
  if ($process.HasExited) { throw "Application exited early with code $($process.ExitCode)." }
  if ($process.MainWindowHandle -eq 0) { throw 'Main window did not appear within 15 seconds.' }
}
finally { if (-not $process.HasExited) { Stop-Process -Id $process.Id -Force } }
Write-Host 'Portable launch verification passed.'
```

- [ ] **Step 3: 锁定传递依赖并创建 win-x64 发布脚本**

Add to `Directory.Build.props`:

```xml
<RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>
```

Run once and retain every generated `packages.lock.json`:

```powershell
dotnet restore Ipv6ProvinceStatistics.sln --force-evaluate
dotnet restore src/Ipv6ProvinceStatistics.App -r win-x64 --force-evaluate
```

Create `build/Publish-WinX64.ps1`:

```powershell
param([string]$Version = '1.0.0')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$artifacts = Join-Path $root 'artifacts'
$publish = Join-Path $artifacts 'publish-win-x64'
$package = Join-Path $artifacts "IPv6省级统计助手-v$Version-win-x64.zip"
Remove-Item $publish -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $package -Force -ErrorAction SilentlyContinue
dotnet restore (Join-Path $root 'Ipv6ProvinceStatistics.sln') --locked-mode
dotnet test (Join-Path $root 'tests/Ipv6ProvinceStatistics.UnitTests') -c Release --no-restore
dotnet test (Join-Path $root 'tests/Ipv6ProvinceStatistics.IntegrationTests') -c Release --no-restore
dotnet test (Join-Path $root 'tests/Ipv6ProvinceStatistics.WindowsTests') -c Release --no-restore
& (Join-Path $PSScriptRoot 'Verify-OfflinePolicy.ps1') -Root $root
dotnet restore (Join-Path $root 'src/Ipv6ProvinceStatistics.App') -r win-x64 --locked-mode
dotnet publish (Join-Path $root 'src/Ipv6ProvinceStatistics.App') -c Release -r win-x64 `
  --self-contained true --no-restore -p:PublishSingleFile=false -p:PublishTrimmed=false `
  -p:Version=$Version -o $publish
Copy-Item (Join-Path $root 'docs/user-guide.md') (Join-Path $publish '使用说明.md')
Compress-Archive -Path (Join-Path $publish '*') -DestinationPath $package -CompressionLevel Optimal
$hash = Get-FileHash $package -Algorithm SHA256
"$($hash.Hash)  $([IO.Path]::GetFileName($package))" | Set-Content "$package.sha256" -Encoding ascii
Write-Host "Package: $package"
Write-Host "SHA256: $($hash.Hash)"
```

- [ ] **Step 4: 在 Windows 上运行完整发布构建**

```powershell
pwsh -File build/Publish-WinX64.ps1 -Version 1.0.0
pwsh -File build/Verify-Portable.ps1 -PublishDirectory artifacts/publish-win-x64
Get-ChildItem artifacts/publish-win-x64 -Recurse | Select-String '5月统计数据|IPv6流量分类统计公式表 - 北京'
```

Expected: 三个测试项目全部 PASS；离线扫描和便携启动通过；最后一条搜索没有输出；生成 ZIP 和对应 `.sha256`。

- [ ] **Step 5: 提交发布基础设施和文档**

```bash
git add Directory.Build.props build README.md docs/user-guide.md src tests
git commit -m "build(release): add win-x64 portable pipeline" -m "Lock dependencies, run all Windows tests, enforce the offline policy, publish a self-contained untrimmed folder, and package documentation with a SHA-256 checksum."
```

### Task 20: 在干净 Windows 与真实样例上执行最终验收

**Files:**
- Verify only: `artifacts/IPv6省级统计助手-v1.0.0-win-x64.zip`
- Verify only: local ignored `5月统计数据/*.xlsx`
- Verify: Git history and working tree

- [ ] **Step 1: 在无 .NET、无 Office 的干净 Windows 10 x64 验证启动**

Copy only the ZIP and `.sha256` to the VM, verify the hash, extract, then run:

```powershell
Get-FileHash .\IPv6省级统计助手-v1.0.0-win-x64.zip -Algorithm SHA256
$publish = Resolve-Path '.\IPv6省级统计助手'
$exe = Join-Path $publish 'IPv6省级统计助手.exe'
$process = Start-Process $exe -PassThru
$deadline = (Get-Date).AddSeconds(15)
do { Start-Sleep -Milliseconds 250; $process.Refresh() }
while (-not $process.HasExited -and $process.MainWindowHandle -eq 0 -and (Get-Date) -lt $deadline)
if ($process.HasExited) { throw "Application exited early with code $($process.ExitCode)." }
if ($process.MainWindowHandle -eq 0) { throw 'Main window did not appear within 15 seconds.' }
Stop-Process -Id $process.Id -Force
```

Expected: 哈希与 `.sha256` 一致；主窗口 15 秒内出现；系统未安装 .NET 和 Microsoft Office 仍可启动。

- [ ] **Step 2: 在干净 Windows 11 x64 重复启动、150% 缩放和中文路径验证**

Extract to `C:\测试目录\IPv6 统计助手\`, set display scaling to 150%, launch, drag four synthetic fixtures, and generate.

Expected: 无裁切、无崩溃；结果目录包含 31 个文件；中文和空格路径均正常。

- [ ] **Step 3: 验证运行期没有网络连接**

```powershell
$p = Start-Process '.\IPv6省级统计助手.exe' -PassThru
Start-Sleep -Seconds 5
$connections = Get-NetTCPConnection -OwningProcess $p.Id -ErrorAction SilentlyContinue
if ($connections) { $connections | Format-Table; throw 'Application opened a network connection.' }
Stop-Process -Id $p.Id
```

Expected: `$connections` 为空。

- [ ] **Step 4: 使用本地真实 5 月四表验证功能与性能**

在 Windows 上从受控本地介质复制四张源表，不复制到仓库。通过 UI 选择四表，确认识别为 2026 年 05 月，生成到空目录；用秒表记录从点击生成到成功提示的时间。

Run after generation:

```powershell
$files = Get-ChildItem '.\2026年05月统计结果' -Filter '*.xlsx'
if ($files.Count -ne 31) { throw "Expected 31 files, got $($files.Count)." }
$expected = @('北京-2026年05月.xlsx','天津-2026年05月.xlsx','新疆-2026年05月.xlsx')
foreach ($name in $expected) { if (-not (Test-Path (Join-Path $files[0].DirectoryName $name))) { throw "Missing $name" } }
```

Expected: 31 个文件、命名正确、总耗时不超过 15 秒。记录 UI 显示的预检通过状态、耗时和上述命令输出，但不记录业务数值；单元格映射和公式缓存的精确性由 Task 13 和 Task 18 的自动化回读测试提供证据。

- [ ] **Step 5: 验证严格失败、取消和目录编号**

依次执行：把一个必填值改为 `NULL`；生成中点击取消；在已有 `2026年05月统计结果` 时再次生成。

Expected: `NULL` 和取消均留下 0 个新正式结果；再次成功生成使用 `2026年05月统计结果 (2)`，不覆盖旧目录。

- [ ] **Step 6: 执行 Microsoft Defender 扫描**

```powershell
$since = Get-Date
Start-MpScan -ScanType CustomScan -ScanPath (Resolve-Path '.\IPv6省级统计助手')
$detections = Get-MpThreatDetection | Where-Object InitialDetectionTime -ge $since
if ($detections) { $detections | Format-List; throw 'Microsoft Defender reported a detection.' }
```

Expected: `$detections` 为空。不要关闭 Defender 或添加排除项。

- [ ] **Step 7: 按证书条件处理签名**

If no valid code-signing certificate is provided, keep the ZIP unsigned and retain the SmartScreen warning in `docs/user-guide.md`. If a certificate is provided, sign before ZIP creation and verify:

```powershell
if (-not $env:SIGN_CERT_PATH -or -not $env:SIGN_CERT_PASSWORD -or -not $env:SIGN_TIMESTAMP_URL) {
  throw 'Set SIGN_CERT_PATH, SIGN_CERT_PASSWORD, and SIGN_TIMESTAMP_URL in the release environment.'
}
signtool sign /fd SHA256 /td SHA256 /tr $env:SIGN_TIMESTAMP_URL `
  /f $env:SIGN_CERT_PATH /p $env:SIGN_CERT_PASSWORD `
  artifacts\publish-win-x64\IPv6省级统计助手.exe
signtool verify /pa /v artifacts\publish-win-x64\IPv6省级统计助手.exe
$package = 'artifacts\IPv6省级统计助手-v1.0.0-win-x64.zip'
Remove-Item $package, "$package.sha256" -Force -ErrorAction SilentlyContinue
Compress-Archive -Path 'artifacts\publish-win-x64\*' -DestinationPath $package -CompressionLevel Optimal
$hash = Get-FileHash $package -Algorithm SHA256
"$($hash.Hash)  $([IO.Path]::GetFileName($package))" | Set-Content "$package.sha256" -Encoding ascii
```

Expected: 未提供证书时明确记录“未签名”；提供证书时 `signtool verify` 成功。证书路径、密码和时间戳地址只从发布负责人处取得，不写入仓库或日志。

- [ ] **Step 8: 最终仓库和提交信息审计**

```bash
git status --short --branch
git log --format='%H%n%an <%ae>%n%s%n%b' --reverse
git log --format='%B' | rg -ni 'Cloud Code|Codex|Copilot|AI Agent|coding tool' && exit 1 || true
git ls-files | rg '5月统计数据|\.xlsx$' | rg -v '^assets/templates/Ipv6ReportTemplate\.xlsx$' && exit 1 || true
```

Expected: 工作区干净；所有作者均为 `HTYSky <htysteve@outlook.com>`；提交格式正确且不含禁止信息；Git 中唯一允许的 `.xlsx` 是脱敏内置模板。

## 完成判定

只有 Task 20 的所有必需步骤都有本次运行的成功证据，才能宣称 v1.0 完成。若缺少真实 Windows、Defender、原始样例或代码签名证书，分别如实报告：前三者阻塞最终验收；证书仅影响签名状态，不阻塞功能验收。
