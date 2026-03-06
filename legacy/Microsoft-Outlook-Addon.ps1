$vt_key = "80a58ac4dbf037bebb6190a350160f451932a4a3cd56085c34e5b6483e058b98"
$temp_dir = "$env:USERPROFILE\Desktop\OutlookScanTemp"

if (!(Test-Path $temp_dir)) { New-Item -ItemType Directory -Path $temp_dir | Out-Null }

try {
    $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch {
    $Outlook = New-Object -ComObject Outlook.Application
}

if ($null -eq $Outlook) { 
    Write-Error "Failed to connect to Outlook. Ensure Outlook is open."
    exit 
}

$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6)
$Items = $Inbox.Items.Restrict("[Unread] = true")

foreach ($Item in $Items) {
    if ($Item.Attachments.Count -gt 0) {
        foreach ($Attachment in $Item.Attachments) {
            if ($Attachment.Size -lt 10240 -or $Attachment.FileName -match "(?i)\.(png|jpg|jpeg|gif)$") { continue }

            $FilePath = Join-Path $temp_dir $Attachment.FileName
            
            try {
                $Attachment.SaveAsFile($FilePath)
                $Hash = (Get-FileHash $FilePath -Algorithm SHA256).Hash
                $Headers = @{"x-apikey" = $vt_key}
                
                $success = $false
                while (-not $success) {
                    try {
                        $Report = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/files/$Hash" -Headers $Headers -ErrorAction Stop
                        $success = $true
                        if ($Report.data.attributes.last_analysis_stats.malicious -gt 0) {
                            Write-Host "!!! DANGER: $($Attachment.FileName) is MALICIOUS." -ForegroundColor Red
                            $Item.Move($Namespace.GetDefaultFolder(3)) | Out-Null
                            break
                        } else { 
                            Write-Host "Clean: $($Attachment.FileName)" -ForegroundColor Green 
                        }
                    } catch {
                        if ($_.Exception.Message -match "429") {
                            Write-Host "Rate limit hit. Cooling down for 60 seconds..." -ForegroundColor Yellow
                            Start-Sleep -Seconds 60
                        } else { 
                            throw $_ 
                        }
                    }
                }
            } catch {
                Write-Warning "Could not process $($Attachment.FileName): $($_.Exception.Message)"
            } finally {
                if (Test-Path $FilePath) { Remove-Item $FilePath -Force }
            }
            
            Start-Sleep -Seconds 16
        }
    }
}