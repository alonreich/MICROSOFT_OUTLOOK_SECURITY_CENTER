param([string]$Mode = "", [int]$ParentPid = 0)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [float]$score = 0, [string]$body = "", [bool]$unread = $false, [string]$scanType = "", [string]$to = "", [string]$cc = "")
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
function Release-And-Collect { param($O) Release-Com -Object $O; [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers() }

function Get-Outlook {
    $o = $null; if (!(Get-Process -Name "outlook" -ErrorAction SilentlyContinue)) { try { Start-Process "outlook.exe" -WindowStyle Minimized; Start-Sleep -Seconds 8 } catch {} }
    try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { try { $o = New-Object -ComObject Outlook.Application; Start-Sleep -Seconds 5 } catch {} }
    if ($null -ne $o) {
        try { $e = $o.ActiveExplorer(); if ($null -ne $e) { $e.WindowState = 1 } } catch {}
        try { $s = $o.Session; if ($null -eq $s) { $o.GetNamespace("MAPI").Logon("Outlook", $null, $false, $false) } elseif (!$s.CurrentUser) { $s.Logon("Outlook", $null, $false, $false) } } catch {}
    }
    return $o
}

function Get-Smtp {
    param($item, $type)
    try {
        if ($type -eq "Sender") {
            if ($item.SenderEmailType -eq "SMTP") { return $item.SenderEmailAddress }
            $pa = $item.PropertyAccessor
            $smtp = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
            if ([string]::IsNullOrWhiteSpace($smtp)) { $smtp = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F") }
            return $smtp
        }
    } catch {}
    return ""
}

function Resolve-Recipients {
    param($recips)
    if ($null -eq $recips) { return "" }
    $resolved = New-Object System.Collections.Generic.List[string]
    foreach ($r in $recips) {
        try {
            if ($r.AddressEntry.Type -eq "SMTP") { [void]$resolved.Add($r.Address) }
            else {
                $pa = $r.PropertyAccessor
                $smtp = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                if ([string]::IsNullOrWhiteSpace($smtp)) { $smtp = $r.Address }
                [void]$resolved.Add($smtp)
            }
        } catch { [void]$resolved.Add($r.Address) }
    }
    return [string]::Join("; ", $resolved)
}

function Robust-Move {
    param($item, $targetFolder)
    try {
        $m = $item.Move($targetFolder)
        if ($null -ne $m) { return $m }
    } catch {
        try {
            $c = $item.Copy()
            $m = $c.Move($targetFolder)
            if ($null -ne $m) { $item.Delete(); return $m }
        } catch { return $null }
    }
    return $null
}

if ($Mode -eq "Release") {
    $H = [Console]::In.ReadLine() | ConvertFrom-Json; if ($null -eq $H -or [string]::IsNullOrWhiteSpace($H.authToken)) { exit }
    $targetIds = $H.targetEntryIds; if ($null -eq $targetIds) { $targetIds = @($H.targetEntryId) }
    $origFolderIds = $H.originalFolders; if ($null -eq $origFolderIds) { $origFolderIds = @($H.originalFolder) }
    $origUnreads = $H.unreads; if ($null -eq $origUnreads) { $origUnreads = @($H.unread) }
    
    function Send-Release-Log {
        param([string]$id, [string]$status, [string]$details, [bool]$success = $false, [string]$newId = "")
        Write-Output (@{type="release-progress"; entryId=$id; status=$status; message=$details; ok=$success; newEntryId=$newId} | ConvertTo-Json -Compress)
    }
    
    $O = Get-Outlook; if ($null -eq $O) { foreach ($id in $targetIds) { Send-Release-Log -id $id -status "Error" -details "Could not connect to Outlook" }; exit }
    $N = $O.GetNamespace("MAPI")
    
    for ($i = 0; $i -lt $targetIds.Count; $i++) {
        $targetId = $targetIds[$i]
        $origFolderId = if ($i -lt $origFolderIds.Count) { $origFolderIds[$i] } else { "" }
        $origUnread = if ($i -lt $origUnreads.Count) { $origUnreads[$i] } else { $false }
        
        Send-Release-Log -id $targetId -status "Starting" -details "Initializing release for item $targetId"
        try {
            $FoundItem = $null; $Attempts = 0; $MaxAttempts = 3
            while ($Attempts -lt $MaxAttempts -and $null -eq $FoundItem) {
                $Attempts++
                try { $FoundItem = $N.GetItemFromID($targetId) } catch {}
                if ($null -eq $FoundItem) {
                    foreach ($st in $N.Stores) { try { $FoundItem = $N.GetItemFromID($targetId, $st.StoreID) } catch {}; if ($null -ne $FoundItem) { break } }
                }
                if ($null -ne $FoundItem) {
                    $CurrentStore = $FoundItem.Parent.Store
                    $TargetFolder = $null
                    if (![string]::IsNullOrWhiteSpace($origFolderId)) { try { $TargetFolder = $N.GetFolderFromID($origFolderId, $CurrentStore.StoreID) } catch {} }
                    if ($null -eq $TargetFolder) { $TargetFolder = $CurrentStore.GetDefaultFolder(6) }
                    if ($FoundItem.Parent.EntryID -eq $TargetFolder.EntryID) {
                        $FoundItem.UnRead = $origUnread
                        Send-Release-Log -id $targetId -status "Finished" -details "Item already in target folder." -success $true -newId $FoundItem.EntryID
                    } else {
                        $MovedItem = Robust-Move -item $FoundItem -targetFolder $TargetFolder
                        if ($null -ne $MovedItem) {
                            $MovedItem.UnRead = $origUnread; $newEntryId = $MovedItem.EntryID
                            Send-Release-Log -id $targetId -status "Finished" -details "Successfully moved." -success $true -newId $newEntryId
                            Release-Com -Object $MovedItem
                        } else { Send-Release-Log -id $targetId -status "Error" -details "Move failed." }
                    }
                } elseif ($Attempts -lt $MaxAttempts) { Start-Sleep -Seconds 2 }
            }
            if ($null -eq $FoundItem) { Send-Release-Log -id $targetId -status "Error" -details "Could not find item." }
        } catch { Send-Release-Log -id $targetId -status "Error" -details $_.Exception.Message } finally { Release-Com -Object $FoundItem; Release-Com -Object $TargetFolder; Release-Com -Object $CurrentStore }
    }
    Release-Com -Object $N; Release-And-Collect -Object $O; exit
}

if ($Mode -eq "Quarantine") {
    $H = [Console]::In.ReadLine() | ConvertFrom-Json; if ($null -eq $H -or [string]::IsNullOrWhiteSpace($H.authToken)) { exit }
    $targetIds = $H.targetEntryIds; if ($null -eq $targetIds) { $targetIds = @($H.targetEntryId) }
    
    function Send-Quar-Log {
        param([string]$id, [string]$status, [string]$details, [bool]$success = $false, [string]$newId = "")
        Write-Output (@{type="quarantine-progress"; entryId=$id; status=$status; message=$details; ok=$success; newEntryId=$newId} | ConvertTo-Json -Compress)
    }
    
    $O = Get-Outlook; if ($null -eq $O) { exit }
    $N = $O.GetNamespace("MAPI")
    foreach ($targetId in $targetIds) {
        $I = $null; try { $I = $N.GetItemFromID($targetId) } catch {}
        if ($null -ne $I) {
            $St = $I.Parent.Store; $Jn = $St.GetDefaultFolder(23)
            $m = Robust-Move -item $I -targetFolder $Jn
            if ($null -ne $m) { Send-Quar-Log -id $targetId -status "Finished" -details "Moved to Junk" -success $true -newId $m.EntryID; Release-Com -Object $m }
            Release-Com -Object $Jn; Release-Com -Object $St; Release-Com -Object $I
        } else { Send-Quar-Log -id $targetId -status "Error" -details "Item not found" }
    }
    Release-Com -Object $N; Release-And-Collect -Object $O; exit
}

$C = [Console]::In.ReadLine(); if ([string]::IsNullOrEmpty($C)) { exit }
$Ex = $null; try { $Ex = $C | ConvertFrom-Json } catch { exit }
if ($null -eq $Ex -or [string]::IsNullOrWhiteSpace($Ex.authToken)) { exit }

$Rm = $Ex.mode; $sk = $Ex.spamKeywords; $ru = $Ex.rubrics; $wl = $Ex.whitelist; $bl = $Ex.blacklist; $pi = $Ex.processedIds; $Vk = $Ex.vtKey; $Pm = $Ex.privacyMode
$ps = New-Object System.Collections.Generic.HashSet[string]; if ($null -ne $pi) { foreach ($id in $pi) { if ($id) { [void]$ps.Add($id) } } }

Send-Heartbeat

$O = Get-Outlook; if ($null -eq $O) { Send-Status -status "Error" -details "Could not connect to Outlook COM object."; exit }
$shaProvider = [System.Security.Cryptography.SHA256]::Create(); $md5Provider = [System.Security.Cryptography.MD5]::Create()
$stType = if ($Rm -eq "OnAccess") { "ON-ACCESS" } else { "ON-DEMAND" }

try {
    $N = $O.GetNamespace("MAPI")
    foreach ($S in $N.Stores) {
        try {
            $R = $S.GetRootFolder(); $stack = New-Object System.Collections.Generic.Stack[object]; $stack.Push($R)
            while ($stack.Count -gt 0) {
                $f = $stack.Pop(); $n = $f.Name
                if ($n -eq "Junk Email" -or $n -eq "Deleted Items" -or $n -eq "Sync Issues" -or $n -eq "Conflicts" -or $n -eq "Local Failures" -or $n -eq "Server Failures") { Release-Com -Object $f; continue }
                
                try {
                    if ($f.DefaultItemType -eq 0) {
                        $Ai = $f.Items; if ($Rm -eq "OnAccess") { try { $Is = $Ai.Restrict("[Unread] = true") } catch { $Is = $Ai } } else { $Is = $Ai }
                        $cnt = 0; try { $cnt = $Is.Count } catch {}
                        if ($cnt -gt 0) {
                            for ($i = 1; $i -le $cnt; $i++) {
                                if ($ParentPid -gt 0 -and -not (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { exit }
                                $t = $null; try { $t = $Is.Item($i) } catch {}
                                if ($null -eq $t) { continue }
                                try {
                                    if ($t.MessageClass -notlike "IPM.Note*") { continue }
                                    if ($t.MessageClass -match "IPM\.Note\.(Storage|Conflict|Schema)") { continue }

                                    $Su = ""; try { $Su = $t.Subject } catch {}
                                    if ($Su -match "^(Synchronization Log:|Modification Resolution)") { continue }

                                    $Id = $t.EntryID; if ($ps.Contains($Id)) { continue }
                                    
                                    $Se = ""; $Do = ""; $IP = "N/A"; $Hs = ""; $by = ""; $curUnread = $false; $pa = $t.PropertyAccessor
                                    try { $curUnread = $t.UnRead } catch {}
                                    
                                    $Se = Get-Smtp -item $t -type "Sender"
                                    if ($Se -match "@") { $Do = $Se.Split("@")[-1] }

                                    try { $Hs = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F"); if ([string]::IsNullOrWhiteSpace($Hs)) { $Hs = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") } } catch {}
                                    if ([string]::IsNullOrWhiteSpace($Hs)) { try { $Hs = "From: $Se`r`nTo: $($t.To)`r`nSubject: $Su`r`nDate: $($t.ReceivedTime)" } catch {} }
                                    
                                    # ROBUST IP EXTRACTION (Issue 5)
                                    $ir = "(?:\d{1,3}\.){3}\d{1,3}"
                                    $lines = $Hs -split "`r`n"; $current = ""; $cleanHs = ""
                                    foreach($line in $lines) { if($line -match "^[A-Za-z0-9-]+:") { if($current) { $cleanHs += $current + "`r`n" }; $current = $line } else { $current += " " + $line.Trim() } }; if($current){ $cleanHs += $current }
                                    
                                    $foundIps = New-Object System.Collections.Generic.List[string]
                                    $parsedLines = $cleanHs -split "`r`n"
                                    # Look from bottom up to find the true origin
                                    for ($lnIdx = $parsedLines.Count - 1; $lnIdx -ge 0; $lnIdx--) { 
                                        $line = $parsedLines[$lnIdx]
                                        if ($line -match "^Received: ") { 
                                            $mips = [regex]::Matches($line, $ir)
                                            foreach ($mip in $mips) { 
                                                $v = $mip.Value
                                                if ($v -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1|169\.254\.)" -and $v -ne "0.0.0.0" -and $v -ne "255.255.255.255") { 
                                                    if (-not $foundIps.Contains($v)) { [void]$foundIps.Add($v) }
                                                } 
                                            } 
                                        } 
                                        if ($line -match "^X-Originating-IP: \s*[\(\[]?(?<v>$ir)[\)\]]?") { 
                                            $v = $Matches['v']
                                            if ($v -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1|169\.254\.)") {
                                                if (-not $foundIps.Contains($v)) { [void]$foundIps.Insert(0, $v) }
                                            }
                                        }
                                    }
                                    
                                    if ($foundIps.Count -gt 0) { 
                                        $IP = $foundIps[0]
                                        foreach ($cand in $foundIps) { if ($cand -notmatch "^(52\.|40\.|104\.|13\.|137\.|191\.|198\.|199\.|20\.|204\.|207\.|208\.|209\.|23\.|40\.|64\.|65\.|94\.|95\.)") { $IP = $cand; break } }
                                    }
                                    
                                    $f_to = Resolve-Recipients -recips $t.Recipients | Where-Object { $_ -match "@" }; if ([string]::IsNullOrWhiteSpace($f_to)) { $f_to = $t.To }
                                    $f_cc = ""; try { $ccList = New-Object System.Collections.Generic.List[string]; foreach ($r in $t.Recipients) { if ($r.Type -eq 2) { $smtp = Resolve-Recipients -recips @($r); if ($smtp) { [void]$ccList.Add($smtp) } } }; $f_cc = [string]::Join("; ", $ccList) } catch {}
                                    $by = $t.Body; if ([string]::IsNullOrWhiteSpace($by)) { try { $html = $t.HTMLBody; if ($html) { $by = $html -replace "<[^>]+>"," " } } catch {} }

                                    if ($null -ne $by -and $by.Length -gt 10000) { $by = $by.Substring(0, 10000) }
                                    if ($null -ne $Hs -and $Hs.Length -gt 5000) { $Hs = $Hs.Substring(0, 5000) }

                                    $bBytes = [System.Text.Encoding]::UTF8.GetBytes($by); $bSha = [BitConverter]::ToString($shaProvider.ComputeHash($bBytes)).Replace("-","").ToLower(); $bMd5 = [BitConverter]::ToString($md5Provider.ComputeHash($bBytes)).Replace("-","").ToLower()
                                    $combo = "$IP|$Do"
                                    
                                    # --- ANALYSIS (SEQUENTIAL, NO RUNSPACE POOL) ---
                                    if ($wl.emails -contains $Se -or $wl.ips -contains $IP -or $wl.domains -contains $Do -or $wl.combos -contains $combo) { 
                                        Send-Status -status "Finished" -details "$Su" -verdict "Safe" -entryId $Id -sender $Se -ip $IP -domain $Do -fullHeaders $Hs -body $by -tier "Trusted (Whitelist)" -unread $curUnread -scanType $stType -to $f_to -cc $f_cc; continue 
                                    }
                                    
                                    $sc = 0.0; $hits = New-Object System.Collections.Generic.List[string]
                                    $isBlacklisted = ($bl.emails -contains $Se -or $bl.ips -contains $IP -or $bl.domains -contains $Do -or $bl.combos -contains $combo)
                                    $mv = "CLEAN"
                                    
                                    if ($isBlacklisted) {
                                        $sc = 100; $mv = "SPAM (Blacklisted)"; [void]$hits.Add("BLACKLIST_HIT")
                                    } else {
                                        $W = $ru.weights; $T = $ru.toggles
                                        if ($T.rbl -and $IP -ne "N/A" -and $IP -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1)") {
                                            $isRBL = $false
                                            foreach ($rbl in @("zen.spamhaus.org", "bl.spamcop.net", "b.barracudacentral.org")) {
                                                try { $rev = ($IP -split "\.")[3..0] -join "."; $d = Resolve-DnsName -Name "$rev.$rbl" -Type A -ErrorAction SilentlyContinue; if ($d) { $isRBL = $true; break } } catch {}
                                            }
                                            if (!$isRBL -and $Do -and $Do -ne "unknown") { try { $d = Resolve-DnsName -Name "$Do.dbl.spamhaus.org" -Type A -ErrorAction SilentlyContinue; if ($d) { $isRBL = $true } } catch {} }
                                            if ($isRBL) { $sc += ($W.rbl / 10.0); [void]$hits.Add("GLOBAL_RBL_HIT") }
                                        }

                                        if ($T.dmarc -and ($Hs -match "dmarc=(fail|none)" -or $Hs -match "Authentication-Results:.*?dmarc=fail")) { $sc += ($W.dmarc / 10.0); [void]$hits.Add("DMARC_AUTH_FAIL") }
                                        if ($T.dkim -and ($Hs -match "dkim=(fail|none)" -or $Hs -match "Authentication-Results:.*?dkim=fail")) { $sc += ($W.dkim / 10.0); [void]$hits.Add("DKIM_SIG_FAIL") }
                                        if ($T.spf -and ($Hs -match "spf=(fail|softfail|none)" -or $Hs -match "Authentication-Results:.*?spf=fail")) { $sc += ($W.spf / 10.0); [void]$hits.Add("SPF_AUTH_FAIL") }

                                        if ($T.rdns -and $IP -ne "N/A" -and $IP -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1)") { 
                                            try { $ptr = [System.Net.Dns]::GetHostEntry($IP).HostName; if (!$ptr -or ($Do -and $ptr -notmatch [regex]::Escape($Do))) { $sc += ($W.rdns / 10.0); [void]$hits.Add("RDNS_NAME_MISMATCH") } } catch { $sc += ($W.rdns / 10.0); [void]$hits.Add("RDNS_RECORDS_MISSING") } 
                                        }
                                        if ($T.alignment -and $Hs -match "Return-Path:.*?<(?<v>.*?)>") { $rp = $Matches['v']; if ($Se -and $rp -and $Se.ToLower() -ne $rp.ToLower()) { $sc += ($W.alignment / 10.0); [void]$hits.Add("SENDER_MISALIGNMENT") } }
                                        
                                        if ($T.heuristics) {
                                            if (![string]::IsNullOrEmpty($Su)) { foreach ($kw in $sk) { if ($Su.IndexOf($kw, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $sc += ($W.heuristics / 10.0); [void]$hits.Add("SUBJECT_SCAM_KEYWORD"); break } } }
                                            if (![string]::IsNullOrEmpty($by)) { foreach ($kw in $sk) { if ($by.IndexOf($kw, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $sc += ($W.heuristics / 10.0); [void]$hits.Add("BODY_SCAM_KEYWORD"); break } } }
                                        }

                                        if ($T.body -and ($by -match "<script" -or $by -match "display:\s*none" -or $by -match "visibility:\s*hidden" -or $by -match "font-size:\s*0")) { $sc += ($W.body / 10.0); [void]$hits.Add("HIDDEN_BODY_ENTROPY") }
                                        
                                        $hitStr = [string]::Join(", ", $hits)
                                        $displayScore = if ($isBlacklisted) { 0 } else { [Math]::Max(0, (100 - ($sc * 10))) }

                                        $verdictThreshold = if ($null -ne $ru.spamThresholdPercent) { $ru.spamThresholdPercent } else { 50 }
                                        if ($displayScore -le $verdictThreshold) { $mv = "SPAM (Heuristics)" }
                                    }
                                    
                                    if ($mv -ne "CLEAN") {                                        $Tgt = $f.Store.GetDefaultFolder(23); $Quar = Robust-Move -item $t -targetFolder $Tgt
                                        if ($null -ne $Quar) { 
                                            Send-Status -status "SPAM FILTERED" -details $Su -verdict "Spam" -action "Quarantined" -entryId $Quar.EntryID -sender $Se -ip $IP -domain $Do -originalFolder $f.EntryID -fullHeaders $Hs -body $by -tier $hitStr -unread $curUnread -scanType $stType -to $f_to -cc $f_cc -score $displayScore
                                            Release-Com -Object $Quar 
                                        }
                                        Release-Com -Object $Tgt
                                    } else {
                                        Send-Status -status "Finished" -details $Su -verdict "Safe" -entryId $Id -sender $Se -ip $IP -domain $Do -originalFolder $f.EntryID -fullHeaders $Hs -body $by -tier "Clean (Passed Security Checks)" -unread $curUnread -scanType $stType -to $f_to -cc $f_cc -score $displayScore
                                    }
                                    
                                } catch {} finally { Release-Com -Object $t; Release-Com -Object $pa }
                            }
                        }
                    }
                } catch {} finally { Release-Com -Object $Is; Release-Com -Object $Ai }

                try { $fs = $f.Folders; foreach ($sub in $fs) { $stack.Push($sub) }; Release-Com -Object $fs } catch {}
                Release-Com -Object $f
            }; Release-Com -Object $R
        } catch {} finally { Release-Com -Object $S }
    }
} catch {
    Send-Status -status "Error" -details "Critical error during scan: $($_.Exception.Message)"
} finally { 
    if ($null -ne $shaProvider) { $shaProvider.Dispose() }
    if ($null -ne $md5Provider) { $md5Provider.Dispose() }
    Release-Com -Object $N; Release-And-Collect -Object $O 
}

Send-Status -status "Idle" -details "Sync completed." -phase "STANDBY"
