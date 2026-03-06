param(
    [string]$ExchangeFile = "",
    [string]$Mode = "",
    [string]$TargetEntryId = "",
    [string]$OriginalFolder = ""
)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [int]$score = 0)
    $msg = @{
        timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        status    = $status; details = $details; verdict = $verdict; action = $action; entryId = $entryId; tier = $tier; phase = $phase; sender = $sender; ip = $ip; domain = $domain; originalFolder = $originalFolder; fullHeaders = $fullHeaders; score = $score
    }
    Write-Output ($msg | ConvertTo-Json -Compress)
}
function Get-Outlook {
    $obj = $null
    try { 
        $obj = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") 
    } catch {
        try { 
            $obj = New-Object -ComObject Outlook.Application
            $ns = $obj.GetNamespace("MAPI")
            $ns.Logon("Outlook", $null, $false, $true)
            $inbox = $ns.GetDefaultFolder(6)
            $explorer = $obj.Explorers.Add($inbox)
            $explorer.WindowState = 1 
        } catch {
            try { $obj = New-Object -ComObject Outlook.Application } catch {}
        }
    }
    return $obj
}
if ($Mode -eq "Release") {
    try {
        $Outlook = Get-Outlook
        if ($null -eq $Outlook) { throw "Outlook session not available." }
        $Namespace = $Outlook.GetNamespace("MAPI")
        $Item = $null
        for ($r=0; $r -lt 5; $r++) {
            try { $Item = $Namespace.GetItemFromID($TargetEntryId); if ($null -ne $Item) { break } } catch { Start-Sleep -Seconds 1 }
        }
        if ($null -eq $Item) { throw "Security Item not found (ID might have changed or item was deleted)." }
        function Find-Folder {
            param($parent, $name)
            if ($parent.Name -eq $name) { return $parent }
            foreach ($sub in $parent.Folders) {
                $found = Find-Folder -parent $sub -name $name
                if ($null -ne $found) { return $found }
            }
            return $null
        }
        $TargetFolder = $null
        if ($OriginalFolder) {
            foreach ($store in $Namespace.Stores) {
                $TargetFolder = Find-Folder -parent $store.GetRootFolder() -name $OriginalFolder
                if ($null -ne $TargetFolder) { break }
            }
        }
        if ($null -eq $TargetFolder) { $TargetFolder = $Namespace.GetDefaultFolder(6) }
        $moved = $false
        for ($i=0; $i -lt 3; $i++) {
            try { 
                $Item.Move($TargetFolder) | Out-Null
                $moved = $true
                break 
            } catch { 
                Start-Sleep -Seconds 2 
            }
        }
        if (-not $moved) { throw "Outlook COM Error: Item is locked by another process or Outlook is busy." }
        $probeSuccess = $false
        $probeMessage = ""
        for ($p=0; $p -lt 5; $p++) {
            Start-Sleep -Seconds 3
            try {
                $VerifyItem = $Namespace.GetItemFromID($TargetEntryId)
                if ($null -ne $VerifyItem) {
                    if ($VerifyItem.Parent.Name -eq $TargetFolder.Name) {
                        $probeSuccess = $true
                        $probeMessage = "VERIFIED: Item physically restored to '$($TargetFolder.Name)'"
                        break
                    } else {
                        $probeMessage = "SYNC WARNING: Item found in unexpected folder '$($VerifyItem.Parent.Name)'"
                    }
                }
            } catch {
                $probeMessage = "PROBE PENDING: Outlook database synchronizing..."
            }
        }
        if ($probeSuccess) {
            Write-Output (@{status="Success"; message=$probeMessage} | ConvertTo-Json -Compress)
        } else {
            Write-Output (@{status="Success"; message="Item move command accepted by Outlook. (Final sync pending in background)"} | ConvertTo-Json -Compress)
        }
    } catch {
        Write-Output (@{status="Error"; message=$_.Exception.Message} | ConvertTo-Json -Compress)
    }
    exit
}
$RunMode = "OnAccess"; $VTKey = ""; $spamKeywords = @(); $rubrics = @{}; $whitelist = @{}; $processedIds = @()
if (![string]::IsNullOrEmpty($ExchangeFile) -and (Test-Path $ExchangeFile)) {
    try {
        $exchange = Get-Content $ExchangeFile -Raw | ConvertFrom-Json
        $RunMode = $exchange.mode; $VTKey = $exchange.vtApiKey; $spamKeywords = $exchange.spamKeywords
        $rubrics = $exchange.rubrics; $whitelist = $exchange.whitelist; $processedIds = $exchange.processedIds
    } catch { Send-Status -status "Error" -details "Config Read Failed"; exit }
}
$processedSet = New-Object System.Collections.Generic.HashSet[string]
if ($null -ne $processedIds) { foreach ($id in $processedIds) { if ($id) { [void]$processedSet.Add($id) } } }
$Outlook = Get-Outlook
if ($null -eq $Outlook) {
    Send-Status -status "Error" -details "Failed to initialize Outlook COM." -phase "STARTUP"
    exit
}
try {
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Inbox = $Namespace.GetDefaultFolder(6)
    $JunkFolder = $Namespace.GetDefaultFolder(23)
    $DeletedFolder = $Namespace.GetDefaultFolder(3)
    if ($RunMode -eq "FullScan") {
        $Items = $Inbox.Items
        $Items.Sort("[ReceivedTime]", $false)
    } else {
        $Items = $Inbox.Items.Restrict("[Unread] = true")
        $Items.Sort("[ReceivedTime]", $true)
    }
    $totalInFolder = $Items.Count
    Send-Status -status "Active" -details "Inbox stream started. Auditing $totalInFolder items..." -phase "STARTUP"
    function Get-HashReputation {
        param([string]$hash)
        if (-not [string]::IsNullOrEmpty($VTKey)) {
            try {
                $vt_res = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/files/$hash" -Headers @{"x-apikey" = $VTKey} -TimeoutSec 10 -ErrorAction Stop
                if ($vt_res.data.attributes.last_analysis_stats.malicious -ge 3) { return "MALICIOUS (VT)" }
                if ($vt_res.data.attributes.last_analysis_stats.malicious -ge 1) { return "SUSPICIOUS (VT)" }
            } catch {}
        }
        return "UNKNOWN"
    }
    for ($idx = 1; $idx -le $totalInFolder; $idx++) {
        $Item = $null
        try { $Item = $Items.Item($idx) } catch { continue }
        if ($null -eq $Item) { continue }
        $Id = $Item.EntryID
        if ($processedSet.Contains($Id)) { 
            $Item = $null; continue
        }
        $Subject = $Item.Subject; $Sender = $Item.SenderEmailAddress; $Domain = $Sender.Split('@')[-1]
        Send-Status -status "Scanning" -details "$Subject" -entryId $Id -phase "FORENSICS" -sender $Sender -domain $Domain
        $score = 0; $isMalicious = $false; $detectionTier = ""; $IP = "N/A"
        $Headers = ""; try { $Headers = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") } catch {}
        if ([string]::IsNullOrWhiteSpace($Headers)) { try { $Headers = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F") } catch {} }
        if ($Headers -match "spf=fail") { $score += 5; $detectionTier = "Identity Check Failed: SPF REJECTED" }
        elseif ($Headers -match "dkim=fail") { $score += 5; $detectionTier = "Identity Check Failed: DKIM REJECTED" }
        elseif ($Headers -match "dmarc=fail") { $score += 5; $detectionTier = "Identity Check Failed: DMARC REJECTED" }
        $ipRegex = "(?:\d{1,3}\.){3}\d{1,3}|(?:[a-fA-F0-9]{0,4}:){2,7}[a-fA-F0-9]{0,4}"
        if ($Headers -match "X-Originating-IP: \s*[\(\[]?(?<val>$ipRegex)[\)\]]?") {
            $IP = $Matches['val']
        }
        if ($IP -eq "N/A") {
            $hops = [regex]::Matches($Headers, "[\(\[](?<val>$ipRegex)[\)\]]")
            for ($i = $hops.Count - 1; $i -ge 0; $i--) {
                $c = $hops[$i].Groups['val'].Value
                if ($c -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1|255\.255\.255\.255)") {
                    $IP = $c; break
                }
            }
        }
        $isWL = $false
        if ($whitelist.emails -contains $Sender) { $isWL = $true }
        if (-not $isWL -and $whitelist.ips -contains $IP) { $isWL = $true }
        if (-not $isWL -and $whitelist.domains -contains $Domain) { $isWL = $true }
        if (-not $isWL) { foreach ($c in $whitelist.combos) { if ($c.ip -eq $IP -and $c.domain -eq $Domain) { $isWL = $true; break } } }
        if ($isWL) { Send-Status -status "Finished" -details "$Subject" -verdict "Clean" -action "Keep in Inbox" -entryId $Id -sender $Sender -ip $IP -domain $Domain -tier "User Whitelist"; $Item = $null; continue }
        foreach ($kw in $spamKeywords) { if ($Subject -match "(?i)$kw") { $score += 3; $detectionTier = "Subject Keyword Match: '$kw'" } }
        $body = $Item.Body; $allHashes = @()
        try {
            $allHashes += ([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($body)) | ForEach-Object { $_.ToString("x2") }) -join ""
            foreach ($att in $Item.Attachments) {
                $bin = $att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102")
                if ($null -ne $bin) { $allHashes += ([System.Security.Cryptography.SHA256]::Create().ComputeHash($bin) | ForEach-Object { $_.ToString("x2") }) -join "" }
            }
        } catch {}
        foreach ($h in $allHashes) {
            $rep = Get-HashReputation -hash $h
            if ($rep -match "MALICIOUS") { $isMalicious = $true; $detectionTier = "MALICIOUS SIGNATURE MATCH: $rep"; break }
            if ($rep -match "SUSPICIOUS") { $score += 5; $detectionTier = "SUSPICIOUS SIGNATURE MATCH: $rep" }
        }
        $origName = $Item.Parent.Name
        if ($isMalicious) {
            $movedItem = $Item.Move($DeletedFolder)
            $newId = $movedItem.EntryID
            Send-Status -status "THREAT BLOCKED" -details "$Subject" -verdict "Malicious" -action "Moved to Deleted Items" -entryId $newId -sender $Sender -ip $IP -domain $Domain -tier $detectionTier -originalFolder $origName -fullHeaders $Headers
        } elseif ($score -ge $rubrics.threshold) {
            $movedItem = $Item.Move($JunkFolder)
            $newId = $movedItem.EntryID
            $verdict = "Spam"; if ($score -ge ($rubrics.threshold + 5)) { $verdict = "Suspicious" }
            $reason = if ($detectionTier) { $detectionTier } else { "Forensic Score: $score" }
            Send-Status -status "SPAM FILTERED" -details "$Subject" -verdict $verdict -action "Moved to Junk Email" -entryId $newId -sender $Sender -ip $IP -domain $Domain -tier $reason -originalFolder $origName -fullHeaders $Headers -score $score
        } else {
            Send-Status -status "Finished" -details "$Subject" -verdict "Clean" -action "Keep in Inbox" -entryId $Id -tier "Multi-Tier Verification Passed" -originalFolder $origName -sender $Sender -ip $IP -domain $Domain -fullHeaders $Headers
        }
        $Item = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        if ($verdict -ne "Clean") { Start-Sleep -Seconds 10 }
    }
} catch { Send-Status -status "Error" -details "Critical Failure: $($_.Exception.Message)" -phase "CRASH" }
Send-Status -status "Idle" -details "Sync completed." -phase "STANDBY"
