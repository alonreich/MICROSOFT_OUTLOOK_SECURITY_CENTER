param(
    [string]$Mode = "",
    [int]$ParentPid = 0
)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [float]$score = 0, [string]$body = "")
    $h = ""; if (![string]::IsNullOrEmpty($fullHeaders)) { try { $h = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($fullHeaders)) } catch {} }
    $b = ""; if (![string]::IsNullOrEmpty($body)) { try { $b = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($body)) } catch {} }
    $ts = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    
    function esc($s) { if ([string]::IsNullOrEmpty($s)) { return "" }; return $s.Replace('\', '\\').Replace('"', '\"').Replace("`r", "").Replace("`n", "\n") }

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
        "`"body`":`"$b`"" +
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

if ($Mode -eq "Release") {
    $H = [Console]::In.ReadLine() | ConvertFrom-Json; if ($null -eq $H -or [string]::IsNullOrWhiteSpace($H.authToken)) { exit }
    $targetId = $H.targetEntryId; $origFolderId = $H.originalFolder;
    
    function Send-Release-Log {
        param([string]$status, [string]$details, [bool]$success = $false, [string]$newId = "")
        Write-Output (@{type="release-progress"; entryId=$targetId; status=$status; message=$details; ok=$success; newEntryId=$newId} | ConvertTo-Json -Compress)
    }

    Send-Release-Log -status "Starting" -details "Initializing release for item $targetId"
    
    try {
        $O = Get-Outlook; if ($null -eq $O) { Send-Release-Log -status "Error" -details "Could not connect to Outlook"; exit }
        $N = $O.GetNamespace("MAPI")
        
        $TargetFolder = $null
        if (![string]::IsNullOrWhiteSpace($origFolderId)) {
            Send-Release-Log -status "Probing" -details "Attempting to locate original folder ($origFolderId)..."
            try { $TargetFolder = $N.GetFolderFromID($origFolderId) } catch { Send-Release-Log -status "Probing" -details "Original folder not found by ID, falling back to Inbox." }
        }
        if ($null -eq $TargetFolder) { $TargetFolder = $N.GetDefaultFolder(6) }
        $TargetFolderId = $TargetFolder.EntryID
        Send-Release-Log -status "Targeting" -details "Destination set to: $($TargetFolder.FullFolderPath)"

        $FoundItem = $null
        $Attempts = 0
        $MaxAttempts = 4
        
        while ($Attempts -lt $MaxAttempts -and $null -eq $FoundItem) {
            $Attempts++
            Send-Release-Log -status "Searching" -details "Searching for item (Attempt $Attempts/$MaxAttempts)..."
            
            try { $FoundItem = $N.GetItemFromID($targetId) } catch {}
            
            if ($null -eq $FoundItem) {
                foreach ($folderType in @(23, 3)) {
                    $Folder = $N.GetDefaultFolder($folderType)
                    Send-Release-Log -status "Searching" -details "Probing $($Folder.Name) folder..."
                    $Items = $Folder.Items; $count = 0; try { $count = $Items.Count } catch {}
                    for ($k = $count; $k -ge 1; $k--) {
                        $j = $null; try { $j = $Items.Item($k); if ($j.EntryID -eq $targetId) { $FoundItem = $j; break } } catch {}
                        if ($null -ne $j -and $null -eq $FoundItem) { Release-Com -Object $j }
                    }
                    Release-Com -Object $Items; Release-Com -Object $Folder
                    if ($null -ne $FoundItem) { break }
                }
            }
            
            if ($null -ne $FoundItem) {
                Send-Release-Log -status "Found" -details "Item located in parent: $($FoundItem.Parent.Name)"
                
                $Parent = $null; try { $Parent = $FoundItem.Parent; if ($Parent.EntryID -eq $TargetFolderId) {
                    Send-Release-Log -status "Finished" -details "Item is already in the target folder." -success $true -newId $FoundItem.EntryID
                    exit
                } } finally { Release-Com -Object $Parent }

                Send-Release-Log -status "Moving" -details "Moving item to $($TargetFolder.Name)..."
                $unRead = $FoundItem.UnRead
                try {
                    $MovedItem = $FoundItem.Move($TargetFolder)
                    if ($null -ne $MovedItem) {
                        $MovedItem.UnRead = $unRead
                        $newEntryId = $MovedItem.EntryID
                        Send-Release-Log -status "Finished" -details "Successfully moved to target folder." -success $true -newId $newEntryId
                        Release-Com -Object $MovedItem
                        exit
                    }
                } catch {
                    Send-Release-Log -status "Moving" -details "Direct move failed ($($_.Exception.Message)). Trying Copy+Delete..."
                    try {
                        $Copy = $FoundItem.Copy()
                        $MovedItem = $Copy.Move($TargetFolder)
                        if ($null -ne $MovedItem) {
                            $MovedItem.UnRead = $unRead
                            $newEntryId = $MovedItem.EntryID
                            $FoundItem.Delete()
                            Send-Release-Log -status "Finished" -details "Successfully copied and restored." -success $true -newId $newEntryId
                            Release-Com -Object $MovedItem
                            exit
                        }
                    } catch {
                        Send-Release-Log -status "Error" -details "All move strategies failed: $($_.Exception.Message)"
                    }
                }
            }
            
            if ($Attempts -lt $MaxAttempts) { Start-Sleep -Seconds 5 }
        }
        
        Send-Release-Log -status "Error" -details "Could not find item $targetId after $MaxAttempts attempts."
    } catch {
        Send-Release-Log -status "Error" -details "Critical exception: $($_.Exception.Message)"
    } finally {
        Release-Com -Object $FoundItem; Release-Com -Object $N; Release-And-Collect -Object $O; Release-Com -Object $TargetFolder
    }
    exit
}

$C = [Console]::In.ReadLine(); if ([string]::IsNullOrEmpty($C)) { exit }
$Ex = $null; try { $Ex = $C | ConvertFrom-Json } catch { exit }
if ($null -eq $Ex -or [string]::IsNullOrWhiteSpace($Ex.authToken)) { exit }
$Rm = $Ex.mode; $sk = $Ex.spamKeywords; $ru = $Ex.rubrics; $wl = $Ex.whitelist; $pi = $Ex.processedIds; $Vk = $Ex.vtKey; $Pm = $Ex.privacyMode
$ps = New-Object System.Collections.Generic.HashSet[string]; if ($null -ne $pi) { foreach ($id in $pi) { if ($id) { [void]$ps.Add($id) } } }
$O = Get-Outlook; if ($null -eq $O) { exit }
$global:sha = [System.Security.Cryptography.SHA256]::Create(); $global:md5 = [System.Security.Cryptography.MD5]::Create()

try {
    $N = $O.GetNamespace("MAPI"); $In = $N.GetDefaultFolder(6); $Jn = $N.GetDefaultFolder(23); $De = $N.GetDefaultFolder(3)
    $tf = New-Object System.Collections.Generic.List[object]
    if ($Rm -eq "FullScan") {
        try {
            $St = $N.Stores; if ($null -eq $St -or $St.Count -eq 0) { foreach ($f in $N.Folders) { [void]$tf.Add($f) } }
            else { foreach ($S in $St) { try { $R = $S.GetRootFolder(); $sk = New-Object System.Collections.Generic.Stack[object]; $sk.Push($R); while ($sk.Count -gt 0) { $p = $sk.Pop(); [void]$tf.Add($p); try { $fs = $p.Folders; foreach ($f in $fs) { $sk.Push($f) }; Release-Com -Object $fs } catch {} }; Release-Com -Object $R } catch {} finally { Release-Com -Object $S } } }
        } catch {}
    } else { [void]$tf.Add($In) }

    $global:VTCB = 0
    function Get-Rep {
        param($s2, $m5)
        if (![string]::IsNullOrEmpty($Vk)) {
            $nw = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
            if ($nw -gt $global:VTCB) {
                try { $u = "https://www.virustotal.com/api/v3/files/$s2"; $h = @{ "x-apikey" = $Vk }; $r = Invoke-RestMethod -Uri $u -Headers $h -Method Get -TimeoutSec 5; if ($r.data.attributes.last_analysis_stats.malicious -gt 0) { return "MALWARE (VT)" } }
                catch { if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -eq 429) { $global:VTCB = $nw + 60 } }
            }
        }
        if (-not $Pm) { try { $d = Resolve-DnsName -Name "$m5.malware.hash.cymru.com" -Type TXT -TimeoutMs 2000 -ErrorAction SilentlyContinue; if ($d.Strings -match "127\.0\.0\.2") { return "MALWARE (Hash DB)" } } catch {} }
        return "CLEAN"
    }

    for ($idx = 0; $idx -lt $tf.Count; $idx++) {
        $f = $tf[$idx]; $Ai = $null; $Is = $null; Send-Status -status "Scanning" -details "Analyzing folder: $($f.Name)"
        try {
            $Ai = $f.Items; if ($Rm -eq "OnAccess") { $Is = $Ai.Restrict("[Unread] = true") } else { $Is = $Ai }
            $cnt = 0; try { $cnt = $Is.Count } catch {}
            if ($cnt -eq 0) { continue }
            $eIds = New-Object System.Collections.Generic.List[string]
            for ($i = 1; $i -le $cnt; $i++) { $t = $null; try { $t = $Is.Item($i); if ($null -ne $t) { $mc = $t.MessageClass; if ($mc -like "IPM.Note*" -or $mc -like "IPM.Post*" -or $mc -like "IPM.Schedule.Meeting*") { [void]$eIds.Add($t.EntryID) } } } catch {} finally { Release-Com -Object $t } }
            foreach ($Id in $eIds) {
                if ($ParentPid -gt 0 -and -not (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { exit }
                Send-Heartbeat; $I = $null
                try {
                    try { $I = $N.GetItemFromID($Id) } catch { continue }; if ($null -eq $I -or ($Rm -eq "OnAccess" -and $ps.Contains($Id))) { continue }
                    $Su = ""; try { $Su = $I.Subject } catch {}; $Se = ""; try { $Se = $I.SenderEmailAddress } catch {}; $Do = "unknown"; if ($Se -match "@") { $Do = $Se.Split('@')[-1] }
                    Send-Status -status "Scanning" -details "$Su" -entryId $Id -sender $Se -domain $Do
                    $sc = 0.0; $mv = "CLEAN"; $IP = "N/A"; $cfId = ""; $pf = $null; try { $pf = $I.Parent; $cfId = $pf.EntryID } catch {} finally { Release-Com -Object $pf }
                    $hits = New-Object System.Collections.Generic.List[string]
                    $pa = $null; $Hs = ""; 
                    try { 
                        $pa = $I.PropertyAccessor; 
                        $Hs = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F"); 
                        if ([string]::IsNullOrWhiteSpace($Hs)) { $Hs = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") } 
                        if ([string]::IsNullOrWhiteSpace($Hs)) { $Hs = $pa.GetProperty("PR_TRANSPORT_MESSAGE_HEADERS") }
                    } catch {}
                    
                    if ([string]::IsNullOrWhiteSpace($Hs)) {
                        # Empirical fallback for internal Exchange/BULK emails where transport headers are stripped
                        try {
                            $f_to = $I.To; $f_from = $I.SenderEmailAddress; $f_date = $I.ReceivedTime;
                            $Hs = "Received: from internal (Exchange/MAPI)`r`nFrom: $f_from`r`nTo: $f_to`r`nDate: $f_date`r`nSubject: $Su`r`n[Note: Standard SMTP transport headers were stripped or not generated by the mail server for this item.]"
                            
                            # Attempt to get Sender SMTP address if primary is X500
                            if ($f_from -match "^/O=") {
                                try { $smtp = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F"); if ($smtp) { $Se = $smtp; $Hs += "`r`nReal-Sender: $smtp" } } catch {}
                            }
                        } catch {}
                    }
                    finally { Release-Com -Object $pa }
                    
                    if (![string]::IsNullOrEmpty($Hs)) {
                        $W = $ru.weights; $T = $ru.toggles; 
                        if ($T.dmarc -and $Hs.Contains("dmarc=fail")) { $sc += ($W.dmarc / 10.0); [void]$hits.Add("DMARC") }; 
                        if ($T.dkim -and $Hs.Contains("dkim=fail")) { $sc += ($W.dkim / 10.0); [void]$hits.Add("DKIM") }; 
                        if ($T.spf -and $Hs.Contains("spf=fail")) { $sc += ($W.spf / 10.0); [void]$hits.Add("SPF") }
                        if ($Hs.Length -lt 1048576) { $ir = "(?:\d{1,3}\.){3}\d{1,3}"; if ($Hs -match "X-Originating-IP: \s*[\(\[]?(?<v>$ir)[\)\]]?") { $IP = $Matches['v'] }; if ($IP -eq "N/A") { $ls = $Hs -split "`r`n"; foreach ($l in $ls) { if ($l -match "from.*?[\(\[](?<v>$ir)[\)\]]") { $v = $Matches['v']; if ($v -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1)") { $IP = $v; break } } } } }
                    }

                    if ($T.rdns -and $IP -ne "N/A") { try { $ptr = [System.Net.Dns]::GetHostEntry($IP).HostName; if ($ptr -and $Do -and $ptr -notmatch [regex]::Escape($Do)) { $sc += ($W.rdns / 10.0); [void]$hits.Add("RDNS") } } catch {} }
                    if ($T.alignment -and $Hs -match "Return-Path: <(?<v>.*?)>") { if ($Se -ne $Matches['v']) { $sc += ($W.alignment / 10.0); [void]$hits.Add("ALIGNMENT") } }
                    
                    $by = ""; 
                    try { 
                        $by = $I.Body;
                        if ([string]::IsNullOrWhiteSpace($by)) {
                            # Empirical fallback: if plain text body is empty, extract from HTML
                            $html = $I.HTMLBody;
                            if (![string]::IsNullOrWhiteSpace($html)) {
                                $by = $html -replace "<[^>]+>"," " -replace "&nbsp;"," " -replace "\s+"," "
                            }
                        }
                    } catch {}; 

                    if ($wl.emails -contains $Se -or $wl.ips -contains $IP -or $wl.domains -contains $Do) { 
                        Send-Status -status "Finished" -details "$Su" -verdict "Safe" -entryId $Id -sender $Se -ip $IP -domain $Do -fullHeaders $Hs -body $by -tier ""; continue 
                    }
                    
                    if ($T.heuristics) { foreach ($kw in $sk) { if ($Su.IndexOf($kw, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $sc += ($W.heuristics / 10.0); [void]$hits.Add("HEURISTICS"); break } } }
                    if ($T.body -and ($by.Contains("<script") -or $by.Contains("display:none"))) { $sc += ($W.body / 10.0); [void]$hits.Add("BODY_ENTROPY") }
                    
                    $tierHits = [string]::Join(", ", $hits)
                    $ah = @(); if (![string]::IsNullOrEmpty($by)) { $b = [System.Text.Encoding]::UTF8.GetBytes($by); $ah += @{ s2 = [BitConverter]::ToString($global:sha.ComputeHash($b)).Replace("-","").ToLower(); m5 = [BitConverter]::ToString($global:md5.ComputeHash($b)).Replace("-","").ToLower() } }
                    $As = $null; try { $As = $I.Attachments; $ac = 0; try { $ac = $As.Count } catch {}; for ($aIdx = 1; $aIdx -le $ac; $aIdx++) { $at = $null; $p = $null; try { $at = $As.Item($aIdx); $p = $at.PropertyAccessor; $sz = 0; try { $sz = $p.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E200003") } catch {}; if ($sz -gt 0 -and $sz -lt 10485760) { $d = $p.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102"); if ($null -ne $d) { $ah += @{ s2 = [BitConverter]::ToString($global:sha.ComputeHash($d)).Replace("-","").ToLower(); m5 = [BitConverter]::ToString($global:md5.ComputeHash($d)).Replace("-","").ToLower() }; $d = $null; [System.GC]::Collect() } } } catch {} finally { Release-Com -Object $p; Release-Com -Object $at } } } catch {} finally { Release-Com -Object $As }
                    foreach ($h in $ah) { $r = Get-Rep -s2 $h.s2 -m5 $h.m5; if ($r -ne "CLEAN") { $mv = $r; [void]$hits.Add("REPUTATION_ENGINE"); break } }
                    
                    $tierHits = [string]::Join(", ", $hits)
                    if ($mv -ne "CLEAN") { $m = $I.Move($De); Send-Status -status "THREAT BLOCKED" -details "$Su" -verdict $mv -action "Deleted" -entryId $m.EntryID -sender $Se -ip $IP -domain $Do -originalFolder $cfId -fullHeaders $Hs -body $by -tier $tierHits; Release-Com -Object $m }
                    elseif ($sc -ge $ru.threshold) { $m = $I.Move($Jn); Send-Status -status "SPAM FILTERED" -details "$Su" -verdict "Spam" -action "Quarantined" -entryId $m.EntryID -sender $Se -ip $IP -domain $Do -originalFolder $cfId -fullHeaders $Hs -score $sc -body $by -tier $tierHits; Release-Com -Object $m }
                    else { Send-Status -status "Finished" -details "$Su" -verdict "Safe" -entryId $Id -originalFolder $cfId -sender $Se -ip $IP -domain $Do -fullHeaders $Hs -score $sc -body $by -tier "" }
                } catch {} finally { Release-Com -Object $I }
                [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
            }
        } finally { Release-Com -Object $Is; Release-Com -Object $Ai; Release-Com -Object $f }
    }
} finally { if ($null -ne $global:sha) { $global:sha.Dispose() }; if ($null -ne $global:md5) { $global:md5.Dispose() }; Release-Com -Object $In; Release-Com -Object $Jn; Release-Com -Object $De; Release-Com -Object $N; Release-And-Collect -Object $O }
Send-Status -status "Idle" -details "Sync completed." -phase "STANDBY"
