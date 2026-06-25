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
