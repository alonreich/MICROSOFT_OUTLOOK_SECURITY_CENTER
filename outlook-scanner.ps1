param([string]$Mode = "", [int]$ParentPid = 0)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Global:StopRequested = $false
$Global:ExcludedFolderIds = New-Object System.Collections.Generic.HashSet[string]

function Get-Fingerprint {
    param($item, $ip)
    try {
        $se = Resolve-Email -Recipient $item.Sender
        $su = $item.Subject
        $rt = $item.ReceivedTime.ToString("yyyyMMddHHmmss")
        $raw = "$se|$su|$rt|$ip"
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
        return [System.BitConverter]::ToString($hash).Replace("-", "").ToLower()
    } catch { return $item.EntryID }
}

function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$originalEntryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [float]$score = 0, [string]$body = "", [bool]$unread = $false, [string]$scanType = "", [string]$to = "", [string]$cc = "", [string]$fingerprint = "")
    $h = ""; if (![string]::IsNullOrEmpty($fullHeaders)) { try { $h = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($fullHeaders)) } catch {} }
    $b = ""; if (![string]::IsNullOrEmpty($body)) { try { $b = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($body)) } catch {} }
    $ts = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $obj = @{
        timestamp = $ts; status = $status; details = $details; verdict = $verdict; action = $action;
        entryId = $entryId; originalEntryId = $originalEntryId; tier = $tier; phase = $phase; sender = $sender;
        ip = $ip; domain = $domain; originalFolder = $originalFolder; fullHeaders = $h; score = $score;
        body = $b; unread = $unread; scanType = $scanType; to = $to; cc = $cc; fingerprint = $fingerprint
    }
    $json = $obj | ConvertTo-Json -Compress
    Write-Output $json
}

function Send-Heartbeat { Write-Output (@{type="heartbeat"; timestamp=(Get-Date -Format "yyyy-MM-dd HH:mm:ss")} | ConvertTo-Json -Compress) }
function Release-Com { param($O) if ($null -ne $O) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($O) } catch {} } }

function Invoke-OutlookMethod {
    param($ScriptBlock, $MaxRetries = 5)
    $retryCount = 0
    while ($retryCount -lt $MaxRetries) {
        try { return & $ScriptBlock }
        catch [System.Runtime.InteropServices.COMException] {
            $code = $_.Exception.ErrorCode
            if ($code -eq -2147418111 -or $code -eq -2147417846) { $retryCount++; Start-Sleep -Seconds 1 } else { throw $_ }
        } catch { throw $_ }
    }
    throw "Outlook busy timeout."
}

function Get-Outlook {
    $o = $null
    $maxAttempts = 25
    try { 
        $o = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") 
    } catch { 
        try { 
            if (!(Get-Process outlook -ErrorAction SilentlyContinue)) {
                Send-Status -status "INFO" -details "Outlook is not running. Locating and launching..."
                $outlookPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe" -ErrorAction SilentlyContinue)."(default)"
                if ([string]::IsNullOrEmpty($outlookPath)) { $outlookPath = "outlook.exe" }
                
                Send-Status -status "INFO" -details "Launching: $outlookPath"
                $proc = Start-Process $outlookPath -WindowStyle Minimized -PassThru
                
                for ($i=0; $i -lt $maxAttempts; $i++) {
                    Start-Sleep -Seconds 1
                    try {
                        $o = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
                        if ($null -ne $o) { 
                            Send-Status -status "INFO" -details "Outlook connection established."
                            break 
                        }
                    } catch {}
                }
            }
            if ($null -eq $o) { 
                Send-Status -status "INFO" -details "Directly creating Outlook COM object..."
                $o = New-Object -ComObject Outlook.Application
                
                $null = $o.Session
                Start-Sleep -Seconds 2
            } 
        } catch { 
            Send-Status -status "ERROR" -details "CRITICAL: Could not start Outlook. Please open it manually. Error: $($_.Exception.Message)"
            return $null 
        } 
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
            foreach ($f in $root.Folders) {
                if ($f.Name -match "^(Sync Issues|Conflicts|Local Failures|Server Failures)$") { [void]$Global:ExcludedFolderIds.Add($f.EntryID) }
                Release-Com $f
            }
            Release-Com $root
        } catch {}
        Release-Com $S
    }
}

function Get-Property {
    param($item, $propName)
    try { return $item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/$propName") } catch { return $null }
}

function Resolve-Email {
    param($Recipient)
    try {
        if ($Recipient.AddressEntry.Type -eq "SMTP") { return $Recipient.AddressEntry.Address }
        $user = $Recipient.AddressEntry.GetExchangeUser()
        if ($user) { return $user.PrimarySmtpAddress }
        $addr = $Recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        if ($addr) { return $addr }
    } catch {}
    return $Recipient.Address
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
    $from = Resolve-Email -Recipient $item.Sender; if ($item.SenderName -and $item.SenderName -ne $from) { $from = "$($item.SenderName) <$from>" }
    $toList = New-Object System.Collections.Generic.List[string]; $ccList = New-Object System.Collections.Generic.List[string]
    foreach ($recip in $item.Recipients) {
        $addr = Resolve-Email -Recipient $recip
        if ($recip.Type -eq 1) { [void]$toList.Add($addr) } elseif ($recip.Type -eq 2) { [void]$ccList.Add($addr) }
    }
    return @{ ip=$senderIp; headers=$headers; from=$from; to=($toList -join "; "); cc=($ccList -join "; ") }
}

function Robust-Move {
    param($item, $targetFolder)
    if (!$item -or !$targetFolder) { return $null }
    try {
        $origUnread = $item.UnRead
        if ($item.Parent -and $item.Parent.EntryID -eq $targetFolder.EntryID) { 
            $item.UnRead = $origUnread; $item.Save(); return $item 
        }
        $m = Invoke-OutlookMethod { $item.Move($targetFolder) }
        if ($null -ne $m) { $m.UnRead = $origUnread; $m.Save(); return $m }
    } catch { return $null }
    return $null
}

$O = Get-Outlook; if (!$O) { exit }
$N = Invoke-OutlookMethod { $O.GetNamespace("MAPI") }
Init-Exclusions $N

if ($Mode -eq "Worker") {
    while ($true) {
        if ($ParentPid -gt 0 -and !(Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { exit }
        $line = [Console]::In.ReadLine()
        if ([string]::IsNullOrEmpty($line)) { Start-Sleep -Milliseconds 100; continue }
        $Ex = try { $line | ConvertFrom-Json } catch { $null }
        if (!$Ex) { continue }
        $Action = $Ex.action; $entryIds = $Ex.entryIds
        if ($Action -eq "Release") {
            for ($i = 0; $i -lt $entryIds.Count; $i++) {
                $id = $entryIds[$i]; $origF = if ($Ex.originalFolders) { $Ex.originalFolders[$i] } else { $null }
                try {
                    $item = $N.GetItemFromID($id); $origUnread = $item.UnRead
                    $fp = Get-Fingerprint $item; [void]$ps.Add($fp)
                    $targetFolder = if ($origF) { try { $N.GetFolderFromID($origF) } catch {} } else { $null }
                    if (!$targetFolder) { $targetFolder = $item.Parent.Store.GetDefaultFolder(6) }
                    $moved = Robust-Move -item $item -targetFolder $targetFolder
                    if ($moved) { [void]$ps.Add($moved.EntryID); Send-Status -status "Finished" -details "Released to Inbox" -entryId $moved.EntryID -originalEntryId $id -action "Released" -unread $moved.UnRead -fingerprint $fp; Release-Com $moved }
                } catch { Send-Status -status "Error" -details $_.Exception.Message -entryId $id }
            }
        } elseif ($Action -eq "Quarantine") {
            foreach ($id in $entryIds) {
                try {
                    $item = $N.GetItemFromID($id); $origUnread = $item.UnRead
                    $fp = Get-Fingerprint $item; [void]$ps.Add($fp)
                    $targetFolder = $item.Parent.Store.GetDefaultFolder(23)
                    $moved = Robust-Move -item $item -targetFolder $targetFolder
                    if ($moved) { [void]$ps.Add($moved.EntryID); Send-Status -status "Finished" -details "Quarantined to Junk" -entryId $moved.EntryID -originalEntryId $id -action "Quarantined" -unread $moved.UnRead -fingerprint $fp; Release-Com $moved }
                } catch { Send-Status -status "Error" -details $_.Exception.Message -entryId $id }
            }
        }
    }
    exit
}

Send-Heartbeat
$C = [Console]::In.ReadLine(); if (!$C) { exit }
$Ex = $C | ConvertFrom-Json
$sk = $Ex.spamKeywords; $ru = $Ex.rubrics; $wl = $Ex.whitelist; $bl = $Ex.blacklist; $Vk = $Ex.vtKey
$ps = New-Object System.Collections.Generic.HashSet[string]; foreach ($id in $Ex.processedIds) { if ($id) { [void]$ps.Add($id) } }

$RunspacePool = [runspacefactory]::CreateRunspacePool(1, 8); $RunspacePool.Open()
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
        if ($job.Handle.AsyncWaitHandle.WaitOne(5000)) {
            $R = $job.PS.EndInvoke($job.Handle); $job.PS.Dispose(); $itemData = $job.Data
            $t = try { $N.GetItemFromID($itemData.Id) } catch { $null }
            if ($t) {
                if ($R.mv -eq "MALICIOUS") {
                    $Tgt = $t.Parent.Store.GetDefaultFolder(3); $m = Robust-Move $t $Tgt
                    if ($m) { 
                        Send-Status -status "THREAT BLOCKED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -originalFolder $itemData.OrigFolder -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by -fingerprint $itemData.Finger; Release-Com $m 
                    }
                } elseif ($R.mv -eq "SPAM") {
                    $Tgt = $t.Parent.Store.GetDefaultFolder(23); $m = Robust-Move $t $Tgt
                    if ($m) { 
                        Send-Status -status "SPAM FILTERED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -originalFolder $itemData.OrigFolder -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by -fingerprint $itemData.Finger; Release-Com $m 
                    }
                } else {
                    Send-Status -status "Finished" -details $itemData.Su -verdict "Safe" -entryId $itemData.Id -originalEntryId $itemData.Id -originalFolder $itemData.OrigFolder -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $t.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by -fingerprint $itemData.Finger
                }
                Release-Com $t
            }
        }
    }
    $CurrentBatch.Clear()
}

function Register-Events($folder) {
    if ($Global:ExcludedFolderIds.Contains($folder.EntryID)) { return }
    if ($folder.DefaultItemType -eq 0) {
        $items = $folder.Items
        $sub = Register-ObjectEvent -InputObject $items -EventName "ItemAdd" -Action { $Global:ScanQueue.Enqueue($Event.SourceEventArgs[0].EntryID) }
    }
    foreach ($f in $folder.Folders) { Register-Events $f; Release-Com $f }
}

foreach ($S in $N.Stores) {
    $R = $S.GetRootFolder(); $stack = New-Object System.Collections.Generic.Stack[object]; $stack.Push($R)
    while ($stack.Count -gt 0) {
        $f = $stack.Pop()
        if (!$Global:ExcludedFolderIds.Contains($f.EntryID) -and $f.DefaultItemType -eq 0) {
            $items = if ($Ex.mode -eq "OnAccess") { $f.Items.Restrict("[UnRead] = True") } else { $f.Items }
            for ($i=1; $i -le $items.Count; $i++) {
                $t = try { $items.Item($i) } catch { $null }
                if ($t -and $t.MessageClass -like "IPM.Note*" -and $t.Subject -notmatch "^(Synchronization Log:|Modification Resolution)") {
                    $fData = Parse-Forensics $t
                    $fp = Get-Fingerprint -item $t -ip $fData.ip
                    if ($ps.Contains($fp)) { Release-Com $t; continue }
                    $origUnread = $t.UnRead
                    $origFolderId = if ($t.Parent -and $t.Parent.EntryID) { $t.Parent.EntryID } else { "" }; 
                    $itemData = @{ Id=$t.EntryID; Su=$t.Subject; Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$t.Body; To=$fData.to; Cc=$fData.cc; OrigFolder=$origFolderId; Finger=$fp }
                    if ($t.UnRead -ne $origUnread) { $t.UnRead = $origUnread; $t.Save() }
                    $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk)
                    $psi.RunspacePool = $RunspacePool; [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
                    if ($CurrentBatch.Count -ge 8) { Process-Batch }
                }
                Release-Com $t
            }
        }
        foreach ($sub in $f.Folders) { $stack.Push($sub) }
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
        if ($t -and $t.MessageClass -like "IPM.Note*" -and $t.Subject -notmatch "^(Synchronization Log:|Modification Resolution)") {
            $fData = Parse-Forensics $t
            $fp = Get-Fingerprint -item $t -ip $fData.ip
            if ($ps.Contains($fp)) { Release-Com $t; continue }
            $origUnread = $t.UnRead
            $origFolderId = if ($t.Parent -and $t.Parent.EntryID) { $t.Parent.EntryID } else { "" }; 
            $itemData = @{ Id=$t.EntryID; Su=$t.Subject; Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$t.Body; To=$fData.to; Cc=$fData.cc; OrigFolder=$origFolderId; Finger=$fp }
            if ($t.UnRead -ne $origUnread) { $t.UnRead = $origUnread; $t.Save() }
            $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk)
            $psi.RunspacePool = $RunspacePool; [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
        }
        Release-Com $t
    }
    if ($CurrentBatch.Count -gt 0) { Process-Batch }
    Start-Sleep -Seconds 2
}

