param([string]$Mode = "", [int]$ParentPid = 0)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Global:StopRequested = $false
$Global:ExcludedFolderIds = New-Object System.Collections.Generic.HashSet[string]

function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$originalEntryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [float]$score = 0, [string]$body = "", [bool]$unread = $false, [string]$scanType = "", [string]$to = "", [string]$cc = "")
    $h = ""; if (![string]::IsNullOrEmpty($fullHeaders)) { try { $h = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($fullHeaders)) } catch {} }
    $b = ""; if (![string]::IsNullOrEmpty($body)) { try { $b = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($body)) } catch {} }
    $ts = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    function esc($s) { if ([string]::IsNullOrEmpty($s)) { return "" }; return $s.Replace('\', '\\').Replace('"', '\"').Replace("`r", "").Replace("`n", "\n") }
    $ur = if ($unread) { "true" } else { "false" }
    $json = "{" +
        "`"timestamp`":`"$(esc $ts)`"," +
        "`"status`":`"$(esc $status)`"," +
        "`"details`":`"$(esc $details)`"," +
        "`"verdict`":`"$(esc $verdict)`"," +
        "`"action`":`"$(esc $action)`"," +
        "`"entryId`":`"$(esc $entryId)`"," +
        "`"originalEntryId`":`"$(esc $originalEntryId)`"," +
        "`"tier`":`"$(esc $tier)`"," +
        "`"phase`":`"$(esc $phase)`"," +
        "`"sender`":`"$(esc $sender)`"," +
        "`"ip`":`"$(esc $ip)`"," +
        "`"domain`":`"$(esc $domain)`"," +
        "`"originalFolder`":`"$(esc $originalFolder)`"," +
        "`"fullHeaders`":`"$h`"," +
        "`"score`":$($score.ToString([System.Globalization.CultureInfo]::InvariantCulture))," +
        "`"body`":`"$b`"," +
        "`"unread`":$ur," +
        "`"scanType`":`"$(esc $scanType)`"," +
        "`"to`":`"$(esc $to)`"," +
        "`"cc`":`"$(esc $cc)`"" +
    "}"
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
    try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { 
        try { $o = New-Object -ComObject Outlook.Application; Start-Sleep -Seconds 2 } catch { return $null } 
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
        $ipMatch = $headers -split "`r`n" | Where-Object { $_ -match "^(X-Sender-IP|X-Originating-IP|X-Remote-IP|X-Client-IP):.*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})" }
        if ($ipMatch) {
            foreach ($m in $ipMatch) {
                if ($m -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                    $cand = $Matches[1]
                    if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") { $senderIp = $cand; break }
                }
            }
        }
        if ($senderIp -eq "N/A") {
            $received = $headers -split "`r`n" | Where-Object { $_ -match "^Received:" }
            $foundIps = New-Object System.Collections.Generic.List[string]
            foreach ($line in $received) {
                $matches = [regex]::Matches($line, "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})")
                foreach ($m in $matches) { [void]$foundIps.Add($m.Value) }
            }
            for ($i = $foundIps.Count - 1; $i -ge 0; $i--) {
                $cand = $foundIps[$i]
                if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") {
                    $senderIp = $cand
                    break
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
    try {
        if ($item.Parent.EntryID -eq $targetFolder.EntryID) { return $item }
        $origUnread = $item.UnRead
        $m = Invoke-OutlookMethod { $item.Move($targetFolder) }
        if ($null -ne $m) { $m.UnRead = $origUnread; $m.Save(); return $m }
    } catch {
        try {
            $origUnread = $item.UnRead; $c = $item.Copy()
            try { $m = Invoke-OutlookMethod { $c.Move($targetFolder) }
                if ($null -ne $m) { $m.UnRead = $origUnread; $m.Save(); try { $item.Delete() } catch {}; return $m } 
            } finally { Release-Com $c }
        } catch { return $null }
    }
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
                    $item = $N.GetItemFromID($id); $targetFolder = if ($origF) { try { $N.GetFolderFromID($origF) } catch {} } else { $null }
                    if (!$targetFolder) { $targetFolder = $item.Parent.Store.GetDefaultFolder(6) }
                    $moved = Robust-Move -item $item -targetFolder $targetFolder
                    if ($moved) { Send-Status -status "Finished" -details "Released to Inbox" -entryId $moved.EntryID -originalEntryId $id -action "Released" -unread $moved.UnRead; Release-Com $moved }
                } catch { Send-Status -status "Error" -details $_.Exception.Message -entryId $id }
            }
        } elseif ($Action -eq "Quarantine") {
            foreach ($id in $entryIds) {
                try {
                    $item = $N.GetItemFromID($id); $targetFolder = $item.Parent.Store.GetDefaultFolder(23)
                    $moved = Robust-Move -item $item -targetFolder $targetFolder
                    if ($moved) { Send-Status -status "Finished" -details "Quarantined to Junk" -entryId $moved.EntryID -originalEntryId $id -action "Quarantined" -unread $moved.UnRead; Release-Com $moved }
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
    $Se = $itemData.Se; $IP = $itemData.IP; $Do = $itemData.Do; $Hs = $itemData.Hs; $by = $itemData.by; $Su = $itemData.Su
    $combo = "$IP|$Do"
    if ($wl.emails -contains $Se -or $wl.ips -contains $IP -or $wl.domains -contains $Do -or $wl.combos -contains $combo) { 
        return @{ mv = "CLEAN"; verdict = "Safe"; tier = "Trusted (Whitelist)"; score = 100; action = "None" }
    }
    if ($bl.emails -contains $Se -or $bl.ips -contains $IP -or $bl.domains -contains $Do -or $bl.combos -contains $combo) { 
        return @{ mv = "SPAM"; verdict = "Spam"; tier = "BLACKLIST_HIT"; score = 0; action = "Quarantined" }
    }
    $sc = 0.0; $hits = New-Object System.Collections.Generic.List[string]; $W = $ru.weights; $T = $ru.toggles
    if ($Vk -and $IP -ne "N/A" -and $IP -notmatch "internal") {
        try {
            $vt = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/ip_addresses/$IP" -Headers @{"x-apikey"=$Vk} -TimeoutSec 5
            if ($vt.data.attributes.last_analysis_stats.malicious -gt 0) { $sc += 50; [void]$hits.Add("VIRUSTOTAL_MALICIOUS") }
        } catch {}
    }
    if ($T.dmarc -and $Hs -match "dmarc=fail") { $sc += ($W.dmarc / 10.0); [void]$hits.Add("DMARC_FAIL") }
    if ($T.spf -and $Hs -match "spf=fail") { $sc += ($W.spf / 10.0); [void]$hits.Add("SPF_FAIL") }
    if ($T.heuristics) {
        foreach ($kw in $sk) { if ($Su -match [regex]::Escape($kw) -or $by -match [regex]::Escape($kw)) { $sc += ($W.heuristics / 10.0); [void]$hits.Add("KEYWORD_MATCH"); break } }
    }
    $score = [Math]::Max(0, (100 - ($sc * 10)))
    $verdict = if ($hits -contains "VIRUSTOTAL_MALICIOUS") { "Malicious" } elseif ($score -le $ru.spamThresholdPercent) { "Spam" } else { "Safe" }
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
                    if ($m) { Send-Status -status "THREAT BLOCKED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by; Release-Com $m }
                } elseif ($R.mv -eq "SPAM") {
                    $Tgt = $t.Parent.Store.GetDefaultFolder(23); $m = Robust-Move $t $Tgt
                    if ($m) { Send-Status -status "SPAM FILTERED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by; Release-Com $m }
                } else {
                    Send-Status -status "Finished" -details $itemData.Su -verdict "Safe" -entryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $t.UnRead -to $itemData.To -cc $itemData.Cc -fullHeaders $itemData.Hs -body $itemData.by
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
                if ($t -and $t.MessageClass -like "IPM.Note*" -and !$ps.Contains($t.EntryID) -and $t.Subject -notmatch "^(Synchronization Log:|Modification Resolution)") {
                    $fData = Parse-Forensics $t
                    $itemData = @{ Id=$t.EntryID; Su=$t.Subject; Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$t.Body; To=$fData.to; Cc=$fData.cc }
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
            $itemData = @{ Id=$t.EntryID; Su=$t.Subject; Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$t.Body; To=$fData.to; Cc=$fData.cc }
            $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk)
            $psi.RunspacePool = $RunspacePool; [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
        }
        Release-Com $t
    }
    if ($CurrentBatch.Count -gt 0) { Process-Batch }
    Start-Sleep -Seconds 2
}
