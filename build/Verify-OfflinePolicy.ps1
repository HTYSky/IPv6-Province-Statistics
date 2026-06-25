param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$errors = @()
$forbidden = @('HttpClient', 'WebClient', 'WebRequest', 'TcpClient', 'Socket', 'SmtpClient', 'FtpWebRequest')
$files = Get-ChildItem -Path (Join-Path $Root 'src') -Filter '*.cs' -Recurse
foreach ($file in $files) {
    $content = Get-Content $file.FullName -Raw
    foreach ($term in $forbidden) {
        if ($content -match $term) { $errors += "$($file.Name): $term" }
    }
}
if ($errors.Count -gt 0) { $errors | ForEach-Object { Write-Error $_ }; exit 1 }
Write-Host "Offline policy verified: 0 violations." -ForegroundColor Green
