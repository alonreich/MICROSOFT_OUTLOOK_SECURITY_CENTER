$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$json = Get-Content "$PSScriptRoot\scanner_input.json" -Raw

$proc = Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "outlook-scanner.ps1" -PassThru -RedirectStandardInput "$PSScriptRoot\scanner_input.json" -RedirectStandardOutput "$PSScriptRoot\scanner_output.bin" -RedirectStandardError "$PSScriptRoot\scanner_error.log"

Start-Sleep -Seconds 12
$hasExited = $proc.HasExited
$exitCode = if ($hasExited) { $proc.ExitCode } else { "STILL_RUNNING" }
if (!$hasExited) {
    Stop-Process -Id $proc.Id -Force
}

$err = if (Test-Path "$PSScriptRoot\scanner_error.log") { Get-Content "$PSScriptRoot\scanner_error.log" -Raw } else { "" }
$outLen = if (Test-Path "$PSScriptRoot\scanner_output.bin") { (Get-Item "$PSScriptRoot\scanner_output.bin").Length } else { 0 }

Write-Output "Exited: $hasExited, Code: $exitCode, OutBytes: $outLen"
Write-Output "Errors: $err"
