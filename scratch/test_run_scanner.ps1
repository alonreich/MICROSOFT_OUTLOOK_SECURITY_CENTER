$json = @{
    mode = "History"
    scanningSpeed = 50
    processedIds = @()
    releasedFingerprints = @()
    spamKeywords = @("viagra", "lottery")
    rubrics = @{ weights = @{}; toggles = @{}; spamThresholdPercent = 40 }
    whitelist = @{ emails = @(); ips = @(); domains = @(); combos = @() }
    blacklist = @{ emails = @(); ips = @(); domains = @(); combos = @() }
    vtKey = ""
    threatIntelligenceLevel = 1
    onAccessEnabled = $true
    onDemandLimit = 1000
    deepHistoryScanEnabled = $true
} | ConvertTo-Json -Compress

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText("$PSScriptRoot\scanner_input.json", "$json`n", $utf8NoBom)

$proc = Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "outlook-scanner.ps1" -PassThru -RedirectStandardInput "$PSScriptRoot\scanner_input.json" -RedirectStandardOutput "$PSScriptRoot\scanner_output.bin" -RedirectStandardError "$PSScriptRoot\scanner_error.log"

Start-Sleep -Seconds 4
$hasExited = $proc.HasExited
$exitCode = if ($hasExited) { $proc.ExitCode } else { "STILL_RUNNING" }
if (!$hasExited) {
    Stop-Process -Id $proc.Id -Force
}

$err = if (Test-Path "$PSScriptRoot\scanner_error.log") { Get-Content "$PSScriptRoot\scanner_error.log" -Raw } else { "" }
$outLen = if (Test-Path "$PSScriptRoot\scanner_output.bin") { (Get-Item "$PSScriptRoot\scanner_output.bin").Length } else { 0 }

Write-Output "Exited: $hasExited, Code: $exitCode, OutBytes: $outLen"
Write-Output "Errors: $err"
