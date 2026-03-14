param([string]$Mode = "", [int]$ParentPid = 0)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

$Global:StopRequested = $false
$Global:ExcludedFolderIds = New-Object System.Collections.Generic.HashSet[string]
$Global:ReleasedFingerprints = New-Object System.Collections.Generic.HashSet[string]
$ps = New-Object System.Collections.Generic.HashSet[string]

function Send-Heartbeat { Write-Output (@{type="heartbeat"; timestamp=(Get-Date -Format "yyyy-MM-dd HH:mm:ss")} | ConvertTo-Json -Compress) }
function Release-Com { param($O) if ($null -ne $O) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($O) } catch {} } }

function Get-Fingerprint {
    param($item, $ip)
    if (!$item) { return [guid]::NewGuid().ToString() }
    try {
        $sender = try { $item.Sender } catch { $null }
        $se = if ($sender) { Resolve-Email -Recipient $sender } else { try { $item.SenderEmailAddress } catch { "Unknown" } }
        $su = if ($item.Subject) { $item.Subject } else { "No Subject" }
        $rt = try { if ($item.ReceivedTime) { $item.ReceivedTime.ToString("yyyyMMddHHmmss") } else { "00000000000000" } } catch { "00000000000000" }
        $raw = "$se|$su|$rt|$ip"
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
        return [System.BitConverter]::ToString($hash).Replace("-", "").ToLower()
    } catch { 
        $eid = try { $item.EntryID } catch { $null }
        if ($eid) { return $eid }
        return [guid]::NewGuid().ToString()
    }
}

function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$originalEntryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [float]$score = 0, [string]$body = "", [bool]$unread = $false, [string]$scanType = "", [string]$to = "", [string]$cc = "", [string]$fingerprint = "", [string]$timestamp = "")
    $h = ""; if (![string]::IsNullOrEmpty($fullHeaders)) { try { $h = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($fullHeaders)) } catch {} }
    $b = ""; if (![string]::IsNullOrEmpty($body)) { try { $b = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($body)) } catch {} }
    $ts = if (![string]::IsNullOrEmpty($timestamp)) { $timestamp } else { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
    $obj = @{
        timestamp = $ts; status = $status; details = $details; verdict = $verdict; action = $action;
        entryId = $entryId; originalEntryId = $originalEntryId; tier = $tier; phase = $phase; sender = $sender;
        ip = $ip; domain = $domain; originalFolder = $originalFolder; fullHeaders = $h; score = $score;
        body = $b; unread = $unread; scanType = $scanType; to = $to; cc = $cc; fingerprint = $fingerprint
    }
    $json = $obj | ConvertTo-Json -Compress
    Write-Output $json
}

function Invoke-OutlookMethod {
    param($ScriptBlock, $MaxRetries = 10)
    $retryCount = 0
    while ($retryCount -lt $MaxRetries) {
        try { 
            # Yield slightly before call to give Outlook breathing room
            Start-Sleep -Milliseconds 10
            return & $ScriptBlock 
        }
        catch [System.Runtime.InteropServices.COMException] {
            $code = $_.Exception.ErrorCode
            # Handle "Busy" or "Call Rejected" errors
            if ($code -eq -2147418111 -or $code -eq -2147417846 -or $code -eq -2147220948) { 
                $retryCount++; 
                Start-Sleep -Seconds ($retryCount * 0.5) 
            } else { throw $_ }
        } catch { throw $_ }
    }
    throw "Outlook busy timeout after $MaxRetries retries."
}

function Get-Outlook {
    $o = $null
    try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { 
        try { 
            if (!(Get-Process outlook -ErrorAction SilentlyContinue)) {
                $path = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe" -ErrorAction SilentlyContinue)."(default)"
                if ([string]::IsNullOrEmpty($path)) { $path = "outlook.exe" }
                Start-Process $path -WindowStyle Minimized
                for ($i=0; $i -lt 20; $i++) {
                    Start-Sleep -Seconds 1
                    try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application"); if ($o) { break } } catch {}
                }
            }
            if (!$o) { $o = New-Object -ComObject Outlook.Application } 
        } catch { return $null } 
    }
    return $o
}

function Init-Exclusions {
    param($Namespace)
    $folderIds = @(3, 4, 5, 16, 23)
    foreach ($S in $Namespace.Stores) {
        foreach ($id in $folderIds) {
            try { $f = $S.GetDefaultFolder($id); if ($f) { [void]$Global:ExcludedFolderIds.Add($f.EntryID); Release-Com $f } } catch {}
        }
        try {
            $root = $S.GetRootFolder()
            if ($root) {
                $flds = $root.Folders
                foreach ($f in $flds) {
                    if ($f.Name -match "^(Sync Issues|Conflicts|Local Failures|Server Failures)$") { [void]$Global:ExcludedFolderIds.Add($f.EntryID) }
                    Release-Com $f
                }
                Release-Com $flds; Release-Com $root
            }
        } catch {}
        Release-Com $S
    }
}

function Get-Property {
    param($item, $propName)
    if (!$item) { return $null }
    try {
        $pa = $item.PropertyAccessor
        $val = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/$propName")
        Release-Com $pa
        return $val
    } catch { return $null }
}

function Resolve-Email {
    param($Recipient)
    if (!$Recipient) { return "Unknown" }
    $addr = $null
    try {
        $ae = $Recipient.AddressEntry
        if ($ae) {
            if ($ae.Type -eq "SMTP") { $addr = $ae.Address }
            else {
                $user = try { $ae.GetExchangeUser() } catch { $null }
                if ($user) { $addr = $user.PrimarySmtpAddress; Release-Com $user }
                if (!$addr) {
                    $pa = try { $Recipient.PropertyAccessor } catch { $null }
                    if ($pa) {
                        try { $addr = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E") } catch {}
                        Release-Com $pa
                    }
                }
            }
            Release-Com $ae
        }
    } catch {}
    if (!$addr) { $addr = try { $Recipient.Address } catch { $null } }
    if ([string]::IsNullOrWhiteSpace($addr)) { return "Unknown" }
    return $addr
}

function Parse-Forensics {
    param($item)
    $headers = Get-Property $item "0x007D001E"
    $senderIp = "N/A"
    if ($headers) {
        $ipRegex = "\b(?:\d{1,3}\.){3}\d{1,3}\b"
        $orig = Get-Property $item "0x0068001E"
        if ($orig -and $orig -match $ipRegex -and $orig -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.|52\.212\.)") { $senderIp = $Matches[0] }
        if ($senderIp -eq "N/A") {
            $received = $headers -split "`r`n" | Where-Object { $_ -match "^Received:" }
            $foundIps = New-Object System.Collections.Generic.List[string]
            foreach ($line in $received) {
                if ($line -match "from.*?(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") { [void]$foundIps.Add($Matches[1]) }
                elseif ($line -match "\[(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\]") { [void]$foundIps.Add($Matches[1]) }
            }
            for ($i = $foundIps.Count - 1; $i -ge 0; $i--) {
                $cand = $foundIps[$i]
                if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.|52\.212\.|52\.9[45]\.)") {
                    $senderIp = $cand; break
                }
            }
        }
    }
    $Snd = try { $item.Sender } catch { $null }
    $from = if ($Snd) { Resolve-Email -Recipient $Snd; Release-Com $Snd } else { "Unknown" }
    if ($from -eq "Unknown") { $from = try { $item.SenderEmailAddress } catch { "Unknown" } }
    $sName = try { $item.SenderName } catch { $null }
    if ($sName -and $sName -ne $from) { $from = "$($sName) <$from>" }
    $toList = New-Object System.Collections.Generic.List[string]; $ccList = New-Object System.Collections.Generic.List[string]
    try {
        $Recips = $item.Recipients
        if ($Recips) {
            foreach ($recip in $Recips) {
                $addr = Resolve-Email -Recipient $recip
                if ($recip.Type -eq 1) { [void]$toList.Add($addr) } elseif ($recip.Type -eq 2) { [void]$ccList.Add($addr) }
                Release-Com $recip
            }
            Release-Com $Recips
        }
    } catch {}
    return @{ ip=$senderIp; headers=$headers; from=$from; to=($toList -join "; "); cc=($ccList -join "; ") }
}

function Verify-Move {
    param($entryId, $targetFolder)
    if (!$entryId -or !$targetFolder) { return $false }
    # Probing for 10 seconds as requested
    for ($i = 0; $i -lt 10; $i++) {
        try {
            $item = $N.GetItemFromID($entryId)
            if ($item) {
                $parent = $item.Parent
                $match = ($parent.EntryID -eq $targetFolder.EntryID)
                Release-Com $parent; Release-Com $item
                if ($match) { return $true }
            }
        } catch {}
        Start-Sleep -Seconds 1
    }
    return $false
}

function Robust-Move {
    param($item, $targetFolder)
    if (!$item -or !$targetFolder) { return $null }
    try {
        $origUnread = $item.UnRead
        $entryId = $item.EntryID
        $parent = try { $item.Parent } catch { $null }
        $isSame = if ($parent) { $res = ($parent.EntryID -eq $targetFolder.EntryID); Release-Com $parent; $res } else { $false }
        if ($isSame) { $item.UnRead = $origUnread; $item.Save(); return $item }
        
        $m = Invoke-OutlookMethod { $item.Move($targetFolder) }
        if ($null -ne $m) {
            $m.UnRead = $origUnread; $m.Save()
            if (Verify-Move -entryId $m.EntryID -targetFolder $targetFolder) {
                return $m
            } else {
                throw "Move verification failed for item: $($item.Subject)"
            }
        }
    } catch {
        Send-Status -status "ERROR" -details "Move failed: $($_.Exception.Message)"
        return $null
    }
    return $null
}

$O = Get-Outlook; if (!$O) { exit }
$N = Invoke-OutlookMethod { $O.GetNamespace("MAPI") }
Init-Exclusions $N

if ($Mode -eq "Worker") {
    $ps = New-Object System.Collections.Generic.HashSet[string]
    while ($true) {
        if ($ParentPid -gt 0 -and !(Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { exit }
        $line = [Console]::In.ReadLine()
        if ([string]::IsNullOrEmpty($line)) { Start-Sleep -Milliseconds 100; continue }
        $Ex = try { $line | ConvertFrom-Json } catch { $null }
        if (!$Ex) { continue }
        $Action = $Ex.action; $entryIds = $Ex.entryIds
        if ($Action -eq "Release") {
            for ($i = 0; $i -lt $entryIds.Count; $i++) {
                if ($i -gt 0) { Start-Sleep -Milliseconds 200 } # Throttle between items
                $id = $entryIds[$i]; $origF = if ($Ex.originalFolders) { $Ex.originalFolders[$i] } else { $null }
                try {
                    $item = $N.GetItemFromID($id); $origUnread = $item.UnRead
                    $fData = Parse-Forensics -item $item
                    $su = $item.Subject; $ts = try { $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") } catch { "" }
                    $fp = Get-Fingerprint -item $item -ip $fData.ip
                    if (!$Global:ReleasedFingerprints.Contains($fp)) { [void]$Global:ReleasedFingerprints.Add($fp) }
                    
                    $targetFolder = if ($origF) { try { $N.GetFolderFromID($origF) } catch {} } else { $null }
                    if (!$targetFolder) { $targetFolder = $item.Parent.Store.GetDefaultFolder(6) }
                    
                    $moved = Robust-Move -item $item -targetFolder $targetFolder
                    if ($moved) { 
                        Send-Status -status "Finished" -details $su -verdict "Safe" -entryId $moved.EntryID -originalEntryId $id -action "Released" -unread $moved.UnRead -fingerprint $fp -sender $fData.from -ip $fData.ip -timestamp $ts
                        Write-Output (@{type="store-update"; key="releasedFingerprints"; value=$fp} | ConvertTo-Json -Compress)
                        Release-Com $moved 
                    } else {
                        Send-Status -status "ERROR" -details "Release failed: Verification timeout or COM error" -entryId $id
                    }
                } catch { Send-Status -status "Error" -details $_.Exception.Message -entryId $id }
            }
        } elseif ($Action -eq "Quarantine") {
            for ($i = 0; $i -lt $entryIds.Count; $i++) {
                if ($i -gt 0) { Start-Sleep -Milliseconds 200 }
                $id = $entryIds[$i]
                try {
                    $item = $N.GetItemFromID($id); $origUnread = $item.UnRead
                    $fData = Parse-Forensics -item $item
                    $su = $item.Subject; $ts = try { $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") } catch { "" }
                    $fp = Get-Fingerprint -item $item -ip $fData.ip
                    $targetFolder = $item.Parent.Store.GetDefaultFolder(23)
                    $moved = Robust-Move -item $item -targetFolder $targetFolder
                    if ($moved) { 
                        [void]$ps.Add($fp); [void]$ps.Add($moved.EntryID)
                        Send-Status -status "Finished" -details $su -verdict "Spam" -entryId $moved.EntryID -originalEntryId $id -action "Quarantined" -unread $moved.UnRead -fingerprint $fp -sender $fData.from -ip $fData.ip -timestamp $ts
                        Release-Com $moved 
                    } else {
                        Send-Status -status "ERROR" -details "Quarantine failed: Verification timeout or COM error" -entryId $id
                    }
                } catch { Send-Status -status "Error" -details $_.Exception.Message -entryId $id }
            }
        } elseif ($Action -eq "Delete") {
            for ($i = 0; $i -lt $entryIds.Count; $i++) {
                if ($i -gt 0) { Start-Sleep -Milliseconds 200 }
                $id = $entryIds[$i]
                try {
                    $item = $N.GetItemFromID($id)
                    $fData = Parse-Forensics -item $item
                    $su = $item.Subject; $ts = try { $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") } catch { "" }
                    $targetFolder = $item.Parent.Store.GetDefaultFolder(3)
                    $moved = Robust-Move -item $item -targetFolder $targetFolder
                    if ($moved) { 
                        Send-Status -status "Finished" -details $su -verdict "Malicious" -entryId $moved.EntryID -originalEntryId $id -action "Deleted" -sender $fData.from -ip $fData.ip -timestamp $ts
                        Release-Com $moved 
                    } else {
                        Send-Status -status "ERROR" -details "Deletion failed: Verification timeout or COM error" -entryId $id
                    }
                } catch { Send-Status -status "Error" -details $_.Exception.Message -entryId $id }
            }
        } elseif ($Action -eq "Check-Existence") {
            $items = $Ex.data.items; $removed = New-Object System.Collections.Generic.List[object]; $foundElsewhere = New-Object System.Collections.Generic.List[object]; $rid = $Ex.rid
            $idx = 0
            foreach ($i in $items) {
                if ($idx % 5 -eq 0 -and $idx -gt 0) { Start-Sleep -Milliseconds 200 }
                $idx++
                $id = $i.entryId; $fp = $i.fingerprint; $cat = $i.category
                $status = "NOT_FOUND" # Default
                try {
                    $item = $N.GetItemFromID($id)
                    if ($item) {
                        $parent = $item.Parent
                        $targetFolderId = if ($cat -eq "malicious") { 3 } elseif ($cat -eq "spam" -or $cat -eq "suspicious") { 23 } else { 6 }
                        $targetFolder = $item.Parent.Store.GetDefaultFolder($targetFolderId)
                        if ($parent.EntryID -eq $targetFolder.EntryID) { 
                            $status = "OK" 
                        } else { 
                            $status = "MOVED"
                            # Collect metadata if it needs re-scanning
                            $fData = Parse-Forensics -item $item
                            $su = $item.Subject; $ts = try { $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") } catch { "" }
                            [void]$foundElsewhere.Add(@{ entryId=$id; fingerprint=$fp; category=$cat; subject=$su; sender=$fData.from; ip=$fData.ip; timestamp=$ts })
                        }
                        Release-Com $parent; Release-Com $targetFolder; Release-Com $item
                    }
                } catch { $status = "NOT_FOUND" }
                
                if ($status -ne "OK") {
                    [void]$removed.Add(@{ entryId = $id; category = $cat; status = $status })
                }
            }
            Write-Output (@{ type = "store-data"; rid = $rid; value = @{ removed = $removed; elsewhere = $foundElsewhere } } | ConvertTo-Json -Compress)
        }
    }
    exit
}

Send-Heartbeat
$C = [Console]::In.ReadLine(); if (!$C) { exit }
$Ex = $C | ConvertFrom-Json
$sk = $Ex.spamKeywords; $ru = $Ex.rubrics; $wl = $Ex.whitelist; $bl = $Ex.blacklist; $Vk = $Ex.vtKey
$ps = New-Object System.Collections.Generic.HashSet[string]; foreach ($id in $Ex.processedIds) { if ($id) { [void]$ps.Add($id) } }
$Global:ReleasedFingerprints = New-Object System.Collections.Generic.HashSet[string]; 
if ($Ex.releasedFingerprints) { foreach ($fp in $Ex.releasedFingerprints) { if ($fp) { [void]$Global:ReleasedFingerprints.Add($fp) } } }

$RunspacePool = [runspacefactory]::CreateRunspacePool(1, 16); $RunspacePool.Open()
$CurrentBatch = New-Object System.Collections.ArrayList
$Global:ScanQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

$AnalysisScript = {
    param($itemData, $sk, $ru, $wl, $bl, $Vk)
    function Log-Info($m) { Write-Output (@{status="INFO"; details=$m; sender=$itemData.Se; ip=$itemData.IP} | ConvertTo-Json -Compress) }
    $Se = $itemData.Se; $IP = $itemData.IP; $Do = $itemData.Do; $Hs = $itemData.Hs; $by = $itemData.by; $Su = $itemData.Su
    $bare = if ($Se -match "<(.+)>$") { $Matches[1] } else { $Se }
    $combo = "$IP|$Do"
    Log-Info "Starting analysis for email: $Su"
    if ($wl.emails -contains $bare -or $wl.ips -contains $IP -or $wl.domains -contains $Do -or $wl.combos -contains $combo) { 
        Log-Info "Email found in whitelist. Marking as Safe."
        return @{ mv = "CLEAN"; verdict = "Safe"; tier = "Trusted (Whitelist)"; score = 100; action = "None" }
    }
    if ($bl.emails -contains $bare -or $bl.ips -contains $IP -or $bl.domains -contains $Do -or $bl.combos -contains $combo) { 
        Log-Info "Email found in blacklist. Marking as Spam."
        return @{ mv = "SPAM"; verdict = "Spam"; tier = "BLACKLIST_HIT"; score = 0; action = "Quarantined" }
    }
    $sc = 0.0; $hits = New-Object System.Collections.Generic.List[string]; $W = $ru.weights; $T = $ru.toggles
    if ($Vk -and $IP -ne "N/A" -and $IP -notmatch "internal") {
        try {
            Log-Info "Querying VirusTotal for IP: $IP"
            $vt = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/ip_addresses/$IP" -Headers @{"x-apikey"=$Vk} -TimeoutSec 5
            $mal = $vt.data.attributes.last_analysis_stats.malicious
            if ($mal -gt 0) { 
                Log-Info "VirusTotal result: MALICIOUS ($mal hits)"
                $sc += 50; [void]$hits.Add("VIRUSTOTAL_MALICIOUS") 
            } else { Log-Info "VirusTotal result: Clean" }
        } catch { Log-Info "VirusTotal query failed or timed out." }
    }
    if ($T.dmarc -and $Hs -match "dmarc=fail") { $sc += ($W.dmarc / 10.0); [void]$hits.Add("DMARC_FAIL") }
    if ($T.spf -and $Hs -match "spf=fail") { $sc += ($W.spf / 10.0); [void]$hits.Add("SPF_FAIL") }
    if ($T.dkim -and $Hs -match "dkim=fail") { $sc += ($W.dkim / 10.0); [void]$hits.Add("DKIM_FAIL") }
    if ($T.alignment -and ($Hs -match "header.from=.*?;.*?(smtp.mailfrom|smtp.auth).*?domain=.*?;.*?fail" -or $Hs -match "alignment=fail")) { $sc += ($W.alignment / 10.0); [void]$hits.Add("ALIGNMENT_FAIL") }
    if ($T.rdns -and ($Hs -match "Received-SPF:.*?helo.*?fail" -or $Hs -match "Authentication-Results:.*?ptr=fail")) { $sc += ($W.rdns / 10.0); [void]$hits.Add("RDNS_FAIL") }
    if ($T.rbl -and ($Hs -match "X-RBL-Warning" -or $Hs -match "blocked using.*?spamhaus")) { $sc += ($W.rbl / 10.0); [void]$hits.Add("RBL_LISTED") }
    if ($T.body -and ($by -match "(https?://[^\s/$.?#].[^\s]*)")) { $sc += ($W.body / 10.0); [void]$hits.Add("SUSPICIOUS_LINKS") }
    if ($T.heuristics) {
        foreach ($kw in $sk) { 
            if ($Su -match [regex]::Escape($kw) -or $by -match [regex]::Escape($kw)) { 
                Log-Info "Keyword match found: $kw"
                $sc += ($W.heuristics / 10.0); [void]$hits.Add("KEYWORD_MATCH"); break 
            } 
        }
    }
    $score = [Math]::Max(0, (100 - ($sc * 10)))
    $verdict = if ($hits -contains "VIRUSTOTAL_MALICIOUS") { "Malicious" } elseif ($score -le $ru.spamThresholdPercent) { "Spam" } else { "Safe" }
    Log-Info "Analysis complete. Score: $score%, Verdict: $verdict"
    return @{ mv = if ($verdict -eq "Malicious") { "MALICIOUS" } elseif ($verdict -eq "Spam") { "SPAM" } else { "CLEAN" }; verdict = $verdict; tier = ([string]::Join(", ", $hits) -replace "^$", "Clean"); score = $score; action = if ($verdict -eq "Safe") { "None" } else { if ($verdict -eq "Malicious") { "Deleted" } else { "Quarantined" } } }
}

function Process-Batch {
    foreach ($job in $CurrentBatch) {
        Send-Heartbeat
        $complete = $job.Handle.AsyncWaitHandle.WaitOne(15000)
        if (!$complete) { try { $job.PS.Stop() } catch {} }
        
        $R = try { $job.PS.EndInvoke($job.Handle) } catch { $null }
        $itemData = $job.Data
        
        if ($complete -and $R) {
            $t = try { $N.GetItemFromID($itemData.Id) } catch { $null }
            if ($t) {
                if ($R.mv -eq "MALICIOUS") {
                    $def3 = try { $t.Parent.Store.GetDefaultFolder(3) } catch { $null }
                    if ($def3) {
                        $m = Robust-Move $t $def3
                        Release-Com $def3
                        if ($m) { 
                            Send-Status -status "THREAT BLOCKED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -originalFolder $itemData.OrigFolder -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by -fingerprint $itemData.Finger; Release-Com $m 
                        }
                    }
                } elseif ($R.mv -eq "SPAM") {
                    $def23 = try { $t.Parent.Store.GetDefaultFolder(23) } catch { $null }
                    if ($def23) {
                        $m = Robust-Move $t $def23
                        Release-Com $def23
                        if ($m) { 
                            Send-Status -status "SPAM FILTERED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -originalFolder $itemData.OrigFolder -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by -fingerprint $itemData.Finger; Release-Com $m 
                        }
                    }
                } else {
                    Send-Status -status "Finished" -details $itemData.Su -verdict "Safe" -entryId $itemData.Id -originalEntryId $itemData.Id -originalFolder $itemData.OrigFolder -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $t.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by -fingerprint $itemData.Finger
                }
                Release-Com $t
            }
        }
        try { $job.PS.Dispose() } catch {}
    }
    $CurrentBatch.Clear()
}

function Register-Events($folder) {
    if ($Global:ExcludedFolderIds.Contains($folder.EntryID)) { return }
    if ($folder.DefaultItemType -eq 0) {
        $items = $folder.Items
        $sub = Register-ObjectEvent -InputObject $items -EventName "ItemAdd" -Action { $Global:ScanQueue.Enqueue($Event.SourceEventArgs[0].EntryID) }
    }
    $flds = try { $folder.Folders } catch { $null }
    if ($flds) {
        foreach ($f in $flds) { Register-Events $f; Release-Com $f }
        Release-Com $flds
    }
}

foreach ($S in $N.Stores) {
    $R = try { $S.GetRootFolder() } catch { $null }
    if (!$R) { Release-Com $S; continue }
    $stack = New-Object System.Collections.Generic.Stack[object]; $stack.Push($R)
    while ($stack.Count -gt 0) {
        Send-Heartbeat
        $f = $stack.Pop()
        $isExcluded = try { $Global:ExcludedFolderIds.Contains($f.EntryID) } catch { $true }
        if (!$isExcluded -and $f.DefaultItemType -eq 0) {
            $coll = $f.Items
            $items = if ($Ex.mode -eq "OnAccess") { try { $coll.Restrict("[UnRead] = True") } catch { $coll } } else { $coll }
            $count = 0
            foreach ($t in $items) {
                $count++
                if ($count % 5 -eq 0) { Start-Sleep -Milliseconds 100 } # Throttling for UI responsiveness
                if ($count % 10 -eq 0) { Send-Heartbeat }
                $isNote = try { $t.MessageClass -like "IPM.Note*" } catch { $false }
                if ($isNote -and $t.Subject -notmatch "^(Synchronization Log:|Modification Resolution)") {
                    $fData = Parse-Forensics $t
                    $fp = Get-Fingerprint -item $t -ip $fData.ip
                    if ($ps.Contains($fp) -or $Global:ReleasedFingerprints.Contains($fp)) { Release-Com $t; continue }
                    $origUnread = try { $t.UnRead } catch { $true }
                    $parent = try { $t.Parent } catch { $null }
                    $origFolderId = if ($parent) { $eid = $parent.EntryID; Release-Com $parent; $eid } else { "" }; 
                    $itemData = @{ Id=$t.EntryID; Su=$t.Subject; Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$t.Body; To=$fData.to; Cc=$fData.cc; OrigFolder=$origFolderId; Finger=$fp }
                    try { if ($t.UnRead -ne $origUnread) { $t.UnRead = $origUnread; $t.Save() } } catch {}
                    $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk)
                    $psi.RunspacePool = $RunspacePool; [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
                    if ($CurrentBatch.Count -ge 8) { Process-Batch }
                }
                Release-Com $t
            }
            if ($items -ne $coll) { Release-Com $items }
            Release-Com $coll
        }
        $flds = try { $f.Folders } catch { $null }
        if ($flds) {
            foreach ($sub in $flds) { $stack.Push($sub) }
            Release-Com $flds
        }
        Release-Com $f
    }
    if ($Ex.mode -eq "OnAccess") { Register-Events $R }
    Release-Com $R; Release-Com $S
}

if ($CurrentBatch.Count -gt 0) { Process-Batch }
Send-Status -status "MONITORING" -details "Initial scan complete. Monitoring for new incoming emails..."

while ($true) {
    if ($ParentPid -gt 0 -and !(Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { exit }
    Send-Heartbeat; while ($Global:ScanQueue.TryDequeue([ref]$id)) {
        $t = try { $N.GetItemFromID($id) } catch { $null }
        $isNote = try { $t.MessageClass -like "IPM.Note*" } catch { $false }
        if ($t -and $isNote -and $t.Subject -notmatch "^(Synchronization Log:|Modification Resolution)") {
            $fData = Parse-Forensics $t
            $fp = Get-Fingerprint -item $t -ip $fData.ip
            if ($ps.Contains($fp)) { Release-Com $t; continue }
            $origUnread = try { $t.UnRead } catch { $true }
            $parent = try { $t.Parent } catch { $null }
            $origFolderId = if ($parent) { $eid = $parent.EntryID; Release-Com $parent; $eid } else { "" }; 
            $itemData = @{ Id=$t.EntryID; Su=$t.Subject; Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$t.Body; To=$fData.to; Cc=$fData.cc; OrigFolder=$origFolderId; Finger=$fp }
            try { if ($t.UnRead -ne $origUnread) { $t.UnRead = $origUnread; $t.Save() } } catch {}
            $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk)
            $psi.RunspacePool = $RunspacePool; [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
        }
        Release-Com $t
    }
    if ($CurrentBatch.Count -gt 0) { Process-Batch }
    Start-Sleep -Seconds 2
}
