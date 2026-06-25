param([string]$Version = "1.0.0")

$root = Split-Path -Parent $PSScriptRoot
$publish = Join-Path $root "artifacts/publish-win-x64"
$package = Join-Path $root "artifacts/IPv6省级统计助手-v$Version-win-x64.zip"

Remove-Item $publish -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $package -Force -ErrorAction SilentlyContinue

Write-Host "Restoring..." -ForegroundColor Cyan
dotnet restore (Join-Path $root "Ipv6ProvinceStatistics.sln") --locked-mode
if ($LASTEXITCODE -ne 0) { throw "Restore failed." }

Write-Host "Testing..." -ForegroundColor Cyan
dotnet test (Join-Path $root "tests/Ipv6ProvinceStatistics.UnitTests") -c Release --no-restore
dotnet test (Join-Path $root "tests/Ipv6ProvinceStatistics.IntegrationTests") -c Release --no-restore
dotnet test (Join-Path $root "tests/Ipv6ProvinceStatistics.WindowsTests") -c Release --no-restore

Write-Host "Publishing..." -ForegroundColor Cyan
dotnet publish (Join-Path $root "src/Ipv6ProvinceStatistics.App") -c Release -r win-x64 `
  --self-contained true -p:PublishSingleFile=false -p:PublishTrimmed=false `
  -p:EnableWindowsTargeting=true -o $publish

Write-Host "Zipping..." -ForegroundColor Cyan
Compress-Archive -Path "$publish/*" -DestinationPath $package -CompressionLevel Optimal
$hash = Get-FileHash $package -Algorithm SHA256
"$($hash.Hash)  $([IO.Path]::GetFileName($package))" | Set-Content "$package.sha256" -Encoding ascii

Write-Host "Done! Package: $package" -ForegroundColor Green
Write-Host "SHA256: $($hash.Hash)" -ForegroundColor Green
