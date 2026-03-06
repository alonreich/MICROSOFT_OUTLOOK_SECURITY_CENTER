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
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "outlook.exe"
            $psi.WindowStyle = "Minimized"
            $p = [System.Diagnostics.Process]::Start($psi)
            for ($i = 0; $i -lt 15; $i++) {
                Start-Sleep -Seconds 2
                try {
                    $obj = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
                    if ($null -ne $obj) { break }
                } catch {}
            }
        } catch {
            try { $obj = New-Object -ComObject Outlook.Application } catch {}
        }
    }
    return $obj
}

# --- RELEASE MODE HANDLER ---
if ($Mode -eq "Release") {
    try {
        $Outlook = Get-Outlook
        if ($null -eq $Outlook) { throw "Outlook session not available." }
        $Namespace = $Outlook.GetNamespace("MAPI")
        $Item = $Namespace.GetItemFromID($TargetEntryId)
        if ($null -eq $Item) { throw "Security Item not found in Outlook store." }

        $TargetFolder = $null
        if ($OriginalFolder) {
            $Inbox = $Namespace.GetDefaultFolder(6)
            if ($Inbox.Name -eq $OriginalFolder) { $TargetFolder = $Inbox }
            else {
                # Check direct subfolders of Inbox
                foreach ($f in $Inbox.Folders) { if ($f.Name -eq $OriginalFolder) { $TargetFolder = $f; break } }
            }
        }

        # Global Fallback to Inbox
        if ($null -eq $TargetFolder) { $TargetFolder = $Namespace.GetDefaultFolder(6) }

        $Item.Move($TargetFolder) | Out-Null

        # --- RESILIENCE PROBE ---
        # Give Outlook a moment to commit the database change
        Start-Sleep -Seconds 2

        try {
            $VerifyItem = $Namespace.GetItemFromID($TargetEntryId)
            if ($VerifyItem.Parent.Name -eq $TargetFolder.Name) {
                Write-Output (@{status="Success"; message="VERIFIED: Item successfully restored to '$($TargetFolder.Name)'"} | ConvertTo-Json -Compress)
            } else {
                throw "SYNC FAILURE: Item reported moved to '$($TargetFolder.Name)' but probe found it in '$($VerifyItem.Parent.Name)'."
            }
        } catch {
            throw "PROBE FAILED: Could not locate item after move. It may have been intercepted by another process or rule. Details: $($_.Exception.Message)"
        }
        } catch {
        Write-Output (@{status="Error"; message=$_.Exception.Message} | ConvertTo-Json -Compress)
        }
    exit
}

# --- STANDARD SCANNER LOGIC ---
$RunMode = "OnAccess"; $VTKey = ""; $spamKeywords = @(); $rubrics = @{}; $whitelist = @{}; $processedIds = @()
if (![string]::IsNullOrEmpty($ExchangeFile) -and (Test-Path $ExchangeFile)) {
    try {
        $exchange = Get-Content $ExchangeFile -Raw | ConvertFrom-Json
        $RunMode = $exchange.mode; $VTKey = $exchange.vtApiKey; $spamKeywords = $exchange.spamKeywords
        $rubrics = $exchange.rubrics; $whitelist = $exchange.whitelist; $processedIds = $exchange.processedIds
    } catch { Send-Status -status "Error" -details "Config Read Failed"; exit }
}

$vt_key = $VTKey
if ([string]::IsNullOrEmpty($vt_key)) { $vt_key = "80a58ac4dbf037bebb6190a350160f451932a4a3cd56085c34e5b6483e058b98" }
$processedSet = New-Object System.Collections.Generic.HashSet[string]
if ($null -ne $processedIds) { foreach ($id in $processedIds) { if ($id) { [void]$processedSet.Add($id) } } }

Send-Status -status "Initializing" -details "Detecting Microsoft Outlook status..." -phase "STARTUP"
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
        if (-not [string]::IsNullOrEmpty($vt_key)) {
            try {
                $vt_res = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/files/$hash" -Headers @{"x-apikey" = $vt_key} -TimeoutSec 10 -ErrorAction Stop
                if ($vt_res.data.attributes.last_analysis_stats.malicious -ge 3) { return "MALICIOUS (VT)" }
                if ($vt_res.data.attributes.last_analysis_stats.malicious -ge 1) { return "SUSPICIOUS (VT)" }
            } catch {}
        }
        try {
            $dns = Resolve-DnsName -Name "$hash.hash.cymru.com" -Type TXT -Timeout 5 -ErrorAction SilentlyContinue
            if ($dns.Strings -match "127.0.0.2") { return "MALICIOUS (Cymru)" }
        } catch {}
        return "UNKNOWN"
    }

    for ($idx = 1; $idx -le $totalInFolder; $idx++) {
        $Item = $null
        try { $Item = $Items.Item($idx) } catch { continue }
        if ($null -eq $Item) { continue }
        $Id = $Item.EntryID
        if ($processedSet.Contains($Id)) { continue }
        $Subject = $Item.Subject; $Sender = $Item.SenderEmailAddress; $Domain = $Sender.Split('@')[-1]
        Send-Status -status "Scanning" -details "$Subject" -entryId $Id -phase "FORENSICS" -sender $Sender -domain $Domain
        $score = 0; $isMalicious = $false; $detectionTier = ""; $IP = "N/A"
        
        # --- ROBUST FORENSIC EXTRACTION ---
        $Headers = ""
        try {
            $Headers = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
            if ([string]::IsNullOrWhiteSpace($Headers)) { $Headers = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F") }
            if ($Headers -match "spf=fail|dkim=fail|dmarc=fail") { $score += 5; $detectionTier = "Identity Verification Failed (Headers)" }
            
            $ipRegex = "(?:\d{1,3}\.){3}\d{1,3}|(?:[a-fA-F0-9]{0,4}:){2,7}[a-fA-F0-9]{0,4}"
            if ($Headers -match "X-Originating-IP: \s*[\(\[]?(?<val>$ipRegex)[\)\]]?") {
                $c = $Matches['val']; if ($c -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1)") { $IP = $c }
            }
            if ($IP -eq "N/A") {
                $hops = [regex]::Matches($Headers, "[\(\[](?<val>$ipRegex)[\)\]]")
                for ($i = $hops.Count - 1; $i -ge 0; $i--) {
                    $c = $hops[$i].Groups['val'].Value
                    if ($c -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1|255\.255\.255\.255)") {
                        if ($c.Contains(".") -or $c.Contains(":")) { $IP = $c; break }
                    }
                }
            }
            if ([string]::IsNullOrWhiteSpace($Headers)) {
                $Headers = "FORENSIC RECONSTRUCTION:`r`nFrom: $($Item.SenderName) <$($Item.SenderEmailAddress)>`r`nTo: $($Item.To)`r`nSubject: $($Item.Subject)`r`nReceived: $($Item.ReceivedTime)"
            }
        } catch {}

        $isWL = $false
        if ($whitelist.emails -contains $Sender) { $isWL = $true }
        if (-not $isWL -and $whitelist.ips -contains $IP) { $isWL = $true }
        if (-not $isWL -and $whitelist.domains -contains $Domain) { $isWL = $true }
        if (-not $isWL) { foreach ($c in $whitelist.combos) { if ($c.ip -eq $IP -and $c.domain -eq $Domain) { $isWL = $true; break } } }
        if ($isWL) { Send-Status -status "Finished" -details "$Subject (Whitelisted)" -verdict "Clean" -action "Keep in Inbox" -entryId $Id -sender $Sender -ip $IP -domain $Domain -tier "User Whitelist"; continue }
        
        $spamMatch = $false
        foreach ($kw in $spamKeywords) { if ($Subject -match "(?i)$kw") { $score += 3; $spamMatch = $true; $detectionTier = "Subject Keyword Match: '$kw'" } }
        
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
            $Item.Move($DeletedFolder) | Out-Null
            Send-Status -status "THREAT BLOCKED" -details "$Subject" -verdict "Malicious" -action "Moved to Deleted Items" -entryId $Id -sender $Sender -ip $IP -domain $Domain -tier $detectionTier -originalFolder $origName -fullHeaders $Headers
        } elseif ($score -ge $rubrics.threshold) {
            $Item.Move($JunkFolder) | Out-Null
            $verdict = "Spam"; if ($score -ge ($rubrics.threshold + 5)) { $verdict = "Suspicious" }
            $reason = if ($detectionTier) { $detectionTier } else { "Forensic Score: $score" }
            Send-Status -status "SPAM FILTERED" -details "$Subject" -verdict $verdict -action "Moved to Junk Email" -entryId $Id -sender $Sender -ip $IP -domain $Domain -tier $reason -originalFolder $origName -fullHeaders $Headers -score $score
        } else {
            Send-Status -status "Finished" -details "$Subject" -verdict "Clean" -action "Keep in Inbox" -entryId $Id -tier "Multi-Tier Verification Passed" -originalFolder $origName -sender $Sender -ip $IP -domain $Domain -fullHeaders $Headers
        }
        if ($verdict -ne "Clean") { Start-Sleep -Seconds 10 }
    }
} catch { Send-Status -status "Error" -details "Critical: $($_.Exception.Message)" -phase "CRASH" }
Send-Status -status "Idle" -details "Sync completed." -phase "STANDBY"
