param([string]$Mode = "", [int]$ParentPid = 0)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- GLOBAL STATE PERSISTENCE ---
if ($null -eq $Global:DupStack) { $Global:DupStack = $null }
if ($null -eq $Global:DupHashes) { $Global:DupHashes = @{} }
if ($null -eq $Global:DupResults) { $Global:DupResults = New-Object System.Collections.Generic.List[object] }
if ($null -eq $Global:DupScannedCount) { $Global:DupScannedCount = 0 }

# --- ENTERPRISE UTILITIES ---

function Send-Heartbeat { 
    if ($ParentPid -gt 0 -and !(Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { 
        [void][System.Environment]::Exit(0) 
    }
    Write-Output (@{type="heartbeat"; timestamp=(Get-Date -Format "yyyy-MM-dd HH:mm:ss")} | ConvertTo-Json -Compress) 
}

function Release-Com { param($O) if ($null -ne $O) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($O) } catch {} } }

function Log-Progress($m) { Write-Output (@{status="INFO"; details=$m} | ConvertTo-Json -Compress) }

function Get-SHA256 {
    param($string)
    if ([string]::IsNullOrEmpty($string)) { return "N/A" }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($string)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return [System.BitConverter]::ToString($hash).Replace("-", "").ToLower()
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
                $user = $null; try { $user = $ae.GetExchangeUser() } catch {}
                if ($user) { $addr = $user.PrimarySmtpAddress; Release-Com $user }
                if (!$addr) {
                    $pa = $null; try { $pa = $Recipient.PropertyAccessor } catch {}
                    if ($pa) {
                        try { $addr = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E") } catch {}
                        Release-Com $pa
                    }
                }
            }
            Release-Com $ae
        }
    } catch {}
    if (!$addr) { try { $addr = $Recipient.Address } catch {} }
    return if ([string]::IsNullOrWhiteSpace($addr)) { "Unknown" } else { $addr }
}

function Get-Fingerprint {
    param($item, $ip)
    if (!$item) { return [guid]::NewGuid().ToString() }
    try {
        $sender = $null; try { $sender = $item.Sender } catch {}
        $se = "Unknown"; if ($sender) { $se = Resolve-Email -Recipient $sender } else { try { $se = $item.SenderEmailAddress } catch {} }
        $su = "No Subject"; if ($item.Subject) { $su = $item.Subject }
        $rt = "00000000000000"; try { if ($item.ReceivedTime) { $rt = $item.ReceivedTime.ToString("yyyyMMddHHmmss") } } catch {}
        return Get-SHA256 "$se|$su|$rt"
    } catch { 
        $eid = $null; try { $eid = $item.EntryID } catch {}
        return if ($eid) { $eid } else { [guid]::NewGuid().ToString() }
    }
}

function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$originalEntryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [float]$score = 0, [string]$body = "", [bool]$unread = $false, [string]$scanType = "", [string]$to = "", [string]$cc = "", [string]$fingerprint = "", [string]$timestamp = "", [int]$count = 0, [int]$total = 0, [string]$currentFolder = "")
    $h = ""; if (![string]::IsNullOrEmpty($fullHeaders)) { try { $h = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($fullHeaders)) } catch {} }
    $b = ""; if (![string]::IsNullOrEmpty($body)) { try { $b = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($body)) } catch {} }
    $ts = if ([string]::IsNullOrEmpty($timestamp)) { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") } else { $timestamp }
    Write-Output (@{
        timestamp=$ts; status=$status; details=$details; verdict=$verdict; action=$action;
        entryId=$entryId; originalEntryId=$originalEntryId; tier=$tier; phase=$phase; sender=$sender;
        ip=$ip; domain=$domain; originalFolder=$originalFolder; fullHeaders=$h; score=$score;
        body=$b; unread=$unread; scanType=$scanType; to=$to; cc=$cc; fingerprint=$fingerprint;
        count=$count; total=$total; currentFolder=$currentFolder
    } | ConvertTo-Json -Compress)
}

function Invoke-OutlookMethod {
    param($ScriptBlock, $MaxRetries = 10)
    $retryCount = 0
    while ($retryCount -lt $MaxRetries) {
        try { Start-Sleep -Milliseconds 10; return & $ScriptBlock }
        catch [System.Runtime.InteropServices.COMException] {
            $code = $_.Exception.ErrorCode
            if ($code -eq -2147418111 -or $code -eq -2147417846 -or $code -eq -2147220948) { $retryCount++; Start-Sleep -Seconds ($retryCount * 0.5) } else { throw $_ }
        } catch { throw $_ }
    }
    throw "Outlook busy timeout after $MaxRetries retries."
}

function Get-Outlook {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { 
        try { 
            if (!(Get-Process outlook -ErrorAction SilentlyContinue)) { Start-Process "outlook.exe" -WindowStyle Minimized; Start-Sleep -Seconds 5 }
            return New-Object -ComObject Outlook.Application 
        } catch { return $null } 
    }
}

function Init-Exclusions {
    param($Namespace)
    $folderIds = @(3, 4, 5, 16, 23)
    foreach ($S in $Namespace.Stores) {
        foreach ($id in $folderIds) {
            try { $f = $S.GetDefaultFolder($id); if ($f) { [void]$Global:ExcludedFolderIds.Add($f.EntryID); Release-Com $f } } catch {}
        }
        Release-Com $S
    }
}

function Parse-Forensics {
    param($item)
    $headers = try { $item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") } catch { $null }
    $senderIp = "N/A"
    if ($headers) {
        $ipRegex = "\b(?:\d{1,3}\.){3}\d{1,3}\b"
        $received = $headers -split "`r`n" | Where-Object { $_ -match "^Received:" }
        foreach ($line in $received) {
            if ($line -match "\[(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\]") { 
                $cand = $Matches[1]
                if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") { $senderIp = $cand; break }
            }
        }
    }
    $from = "Unknown"; try { $Snd = $item.Sender; if ($Snd) { $from = Resolve-Email -Recipient $Snd; Release-Com $Snd } else { $from = $item.SenderEmailAddress } } catch {}
    $body = ""; try { $body = $item.Body } catch {}
    $atts = New-Object System.Collections.Generic.List[object]
    try {
        $AttsObj = $item.Attachments
        if ($AttsObj) {
            foreach ($at in $AttsObj) {
                [void]$atts.Add(@{ name=$at.FileName; hash="N/A" })
                Release-Com $at
            }
            Release-Com $AttsObj
        }
    } catch {}
    return @{ ip=$senderIp; headers=$headers; from=$from; body=$body; attachments=$atts }
}

function Robust-Move {
    param($item, $targetFolder)
    if (!$item -or !$targetFolder) { return $null }
    try {
        $origUnread = $item.UnRead; $subj = $item.Subject
        $m = Invoke-OutlookMethod { $item.Move($targetFolder) }
        if ($null -ne $m) { $m.UnRead = $origUnread; $m.Save(); return $m }
    } catch { Log-Progress "Robust-Move Error: $($_.Exception.Message)" }
    return $null
}

# --- WORKER MODE ---

if ($Mode -eq "Worker") {
    $O = Get-Outlook; if (!$O) { exit }
    $N = $O.GetNamespace("MAPI")
    Log-Progress "Security Engine Worker: ACTIVE. Monitoring Parent PID: $ParentPid"
    
    while ($true) {
        Send-Heartbeat
        $line = [Console]::In.ReadLine()
        if ([string]::IsNullOrEmpty($line)) { Start-Sleep -Milliseconds 200; continue }
        $Ex = try { $line | ConvertFrom-Json } catch { $null }
        if (!$Ex) { continue }
        
        $Action = $Ex.action
        if ($Action -eq "DuplicateScan") {
            $pauseFile = Join-Path $PSScriptRoot ".dup_pause"
            Log-Progress "Worker: Starting Duplicate Email Discovery phase..."
            
            # Initialize or Resume
            if ($null -eq $Global:DupStack) {
                Log-Progress "Worker: Building master folder stack from Outlook stores..."
                $Global:DupStack = New-Object System.Collections.Generic.Stack[object]
                $Global:DupHashes = @{}
                $Global:DupResults = New-Object System.Collections.Generic.List[object]
                $Global:DupScannedCount = 0
                
                try {
                    foreach ($S in $N.Stores) {
                        Log-Progress "Worker: Indexing Store: $($S.DisplayName)"
                        @(6, 5) | ForEach-Object { 
                            try { 
                                $f = $S.GetDefaultFolder($_); 
                                if ($f) { 
                                    $Global:DupStack.Push(@{Folder=$f; Store=$S.DisplayName}) 
                                    Log-Progress "Worker: Added folder to discovery stack: $($f.Name)"
                                } 
                            } catch { Log-Progress "Worker: Store check failed for ID $_" } 
                        }
                    }
                } catch { Log-Progress "Worker: CRITICAL ERROR during store indexing: $($_.Exception.Message)" }
            }

            if ($Global:DupStack.Count -eq 0) {
                Log-Progress "Worker: FATAL - Discovery stack is empty. Verify Outlook connectivity."
            }

            while ($Global:DupStack.Count -gt 0) {
                # Check for Pause signal
                if (Test-Path $pauseFile) {
                    Write-Output (@{type="duplicate-update"; status="Paused"} | ConvertTo-Json -Compress)
                    break
                }

                $entry = $Global:DupStack.Pop()
                $f = $entry.Folder; $storeName = $entry.Store
                Write-Output (@{type="duplicate-update"; status="Progress"; details="Scanning: $storeName \ $($f.Name)"} | ConvertTo-Json -Compress)
                
                if ($f.DefaultItemType -eq 0) {
                    $items = $f.Items
                    foreach ($t in $items) {
                        try {
                            $Global:DupScannedCount++
                            $sender = "Unknown"; try { $sender = Resolve-Email -Recipient $t.Sender } catch { try { $sender = $t.SenderEmailAddress } catch {} }
                            $subject = ($t.Subject -replace "^(Re:|Fwd:|FW:|RE:)\s*", "").Trim()
                            $received = $t.ReceivedTime.ToString("yyyyMMddHHmmss")
                            $bodySample = if ($t.Body) { $t.Body.Substring(0, [Math]::Min($t.Body.Length, 500)).Trim() } else { "" }
                            $dna = Get-SHA256 "$sender|$subject|$received|$bodySample"
                            
                            $itemObj = @{ entryId=$t.EntryID; subject=$t.Subject; sender=$sender; timestamp=$t.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"); size=$t.Size; folder=$f.Name; store=$storeName; dna=$dna }
                            if (!$Global:DupHashes.ContainsKey($dna)) { $Global:DupHashes[$dna] = $itemObj }
                            else {
                                $survivor = $Global:DupHashes[$dna]
                                if ($itemObj.size -gt $survivor.size) { 
                                    [void]$Global:DupResults.Add($survivor); 
                                    $Global:DupHashes[$dna] = $itemObj 
                                } else { 
                                    [void]$Global:DupResults.Add($itemObj) 
                                }
                                Write-Output (@{type="duplicate-update"; status="Found"; count=$Global:DupResults.Count; current=$itemObj.subject} | ConvertTo-Json -Compress)
                            }
                        } catch {}
                        Release-Com $t
                    }
                    Release-Com $items
                }
                try { $flds = $f.Folders; if ($flds) { foreach ($sub in $flds) { $Global:DupStack.Push(@{Folder=$sub; Store=$storeName}) }; Release-Com $flds } } catch {}
                Release-Com $f
            }

            if ($Global:DupStack.Count -eq 0) {
                Write-Output (@{type="duplicate-update"; status="Finished"; items=$Global:DupResults} | ConvertTo-Json -Compress)
                # Clear state for next fresh scan
                $Global:DupStack = $null
            }
        }
        elseif ($Action -eq "Delete") {
            $count = 0
            foreach ($id in $Ex.entryIds) {
                try {
                    $item = $N.GetItemFromID($id)
                    if ($item) {
                        $subj = $item.Subject
                        $item.Delete()
                        $count++
                        Log-Progress "Worker: Successfully deleted redundant copy: $subj"
                        Release-Com $item
                    }
                } catch {}
            }
            Log-Progress "Worker: Batch cleanup complete. $count items removed."
            Write-Output (@{type="delete-summary"; count=$count} | ConvertTo-Json -Compress)
        }
    }
}

# --- SCANNER MODE ---

$O = Get-Outlook; if (!$O) { exit }
$N = $O.GetNamespace("MAPI")
$Global:ExcludedFolderIds = New-Object System.Collections.Generic.HashSet[string]
Init-Exclusions $N
$Global:ReleasedFingerprints = New-Object System.Collections.Generic.HashSet[string]

function Process-Batch {
    $remaining = [System.Collections.Generic.List[object]]::new($CurrentBatch)
    $CurrentBatch.Clear()
    while ($remaining.Count -gt 0) {
        Send-Heartbeat; $toRemove = New-Object System.Collections.Generic.List[object]
        foreach ($job in $remaining) {
            if ($job.Handle.IsCompleted) {
                [void]$toRemove.Add($job)
                $output = try { $job.PS.EndInvoke($job.Handle) } catch { $null }
                $itemData = $job.Data; $R = $null
                if ($output) { foreach ($obj in $output) { if ($obj.mv) { $R = $obj } } }
                if ($R) {
                    $t = try { $N.GetItemFromID($itemData.Id) } catch { $null }
                    if ($t) {
                        if ($Global:ReleasedFingerprints.Contains($itemData.Finger)) { $R.mv = "CLEAN"; $R.verdict = "Safe" }
                        if ($R.mv -eq "MALICIOUS") {
                            $def3 = try { $t.Parent.Store.GetDefaultFolder(3) } catch { $null }
                            if ($def3) { $m = Robust-Move $t $def3; if ($m) { [void]$ps.Add($itemData.Finger); Send-Status -status "THREAT BLOCKED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -fingerprint $itemData.Finger; Release-Com $m } }
                        } elseif ($R.mv -eq "SPAM") {
                            $def23 = try { $t.Parent.Store.GetDefaultFolder(23) } catch { $null }
                            if ($def23) { $m = Robust-Move $t $def23; if ($m) { [void]$ps.Add($itemData.Finger); Send-Status -status "SPAM FILTERED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -fingerprint $itemData.Finger; Release-Com $m } }
                        } else { 
                            [void]$ps.Add($itemData.Finger)
                            Send-Status -status "Finished" -details $itemData.Su -verdict "Safe" -entryId $itemData.Id -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $t.UnRead -fingerprint $itemData.Finger 
                        }
                        Release-Com $t
                    }
                }
                try { $job.PS.Dispose() } catch {}
            }
        }
        foreach ($r in $toRemove) { [void]$remaining.Remove($r) }
        if ($remaining.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }
}

$AnalysisScript = {
    param($itemData, $sk, $ru, $wl, $bl, $Vk)
    $sc = 0.0; $hits = New-Object System.Collections.Generic.List[string]; $W = $ru.weights; $T = $ru.toggles
    $Se = $itemData.Se; $IP = $itemData.IP; $Do = $itemData.Do; $bare = if ($Se -match "<(.+)>$") { $Matches[1] } else { $Se }; $combo = "$IP|$Do"
    if ($wl.emails -contains $bare -or $wl.ips -contains $IP -or $wl.domains -contains $Do -or $wl.combos -contains $combo) { return @{ mv = "CLEAN"; verdict = "Safe"; score = 100; tier = "Whitelisted" } }
    if ($bl.emails -contains $bare -or $bl.ips -contains $IP -or $bl.domains -contains $Do -or $bl.combos -contains $combo) { return @{ mv = "SPAM"; verdict = "Spam"; score = 0; tier = "Blacklisted" } }
    if ($T.dmarc) { if ($itemData.Hs -match "dmarc=fail") { $sc += ($W.dmarc / 10.0); [void]$hits.Add("DMARC:FAIL") } }
    if ($T.spf) { if ($itemData.Hs -match "spf=fail") { $sc += ($W.spf / 10.0); [void]$hits.Add("SPF:FAIL") } }
    if ($T.dkim) { if ($itemData.Hs -match "dkim=fail") { $sc += ($W.dkim / 10.0); [void]$hits.Add("DKIM:FAIL") } }
    if ($T.heuristics) { foreach ($kw in $sk) { if ($itemData.Su -match [regex]::Escape($kw) -or $itemData.by -match [regex]::Escape($kw)) { $sc += ($W.heuristics / 10.0); [void]$hits.Add("HEURISTICS:MATCH($kw)"); break } } }
    $score = [Math]::Max(0, (100 - ($sc * 10)))
    $verdict = if ($score -le 20) { "Malicious" } elseif ($score -le $ru.spamThresholdPercent) { "Spam" } else { "Safe" }
    return @{ mv = if ($verdict -eq "Malicious") { "MALICIOUS" } elseif ($verdict -eq "Spam") { "SPAM" } else { "CLEAN" }; verdict = $verdict; score = $score; tier = ([string]::Join(" | ", $hits) -replace "^$", "Analysis Complete") }
}

Send-Heartbeat; $C = [Console]::In.ReadLine(); if (!$C) { exit }
$Ex = $C | ConvertFrom-Json; $sk = $Ex.spamKeywords; $ru = $Ex.rubrics; $wl = $Ex.whitelist; $bl = $Ex.blacklist; $Vk = $Ex.vtKey
if ($Ex.releasedFingerprints) { foreach ($fp in $Ex.releasedFingerprints) { [void]$Global:ReleasedFingerprints.Add($fp) } }
$ps = New-Object System.Collections.Generic.HashSet[string]; foreach ($id in $Ex.processedIds) { [void]$ps.Add($id) }

$RunspacePool = [runspacefactory]::CreateRunspacePool(1, 16); $RunspacePool.Open(); $CurrentBatch = New-Object System.Collections.Generic.List[object]; $Global:ScanQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
Log-Progress "Forensic Discovery: Starting master mailbox crawl..."
$stack = New-Object System.Collections.Generic.Stack[object]
foreach ($S in $N.Stores) { @(6, 5) | ForEach-Object { try { $f = $S.GetDefaultFolder($_); if ($f) { $stack.Push($f) } } catch {} } }

while ($stack.Count -gt 0) {
    $f = $stack.Pop(); if ($f.DefaultItemType -eq 0) {
        $items = $f.Items
        foreach ($t in $items) {
            $fData = Parse-Forensics $t; $fp = Get-Fingerprint -item $t -ip $fData.ip
            if ($ps.Contains($fp) -or $Global:ReleasedFingerprints.Contains($fp)) { Release-Com $t; continue }
            $itemData = @{ Id=$t.EntryID; Su=$t.Subject; Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$fData.body; Finger=$fp }
            $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk)
            $psi.RunspacePool = $RunspacePool; [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
            if ($CurrentBatch.Count -ge 8) { Process-Batch }
            Release-Com $t
        }
        Release-Com $items
    }
    try { $flds = $f.Folders; if ($flds) { foreach ($sub in $flds) { $stack.Push($sub) }; Release-Com $flds } } catch {}
    Release-Com $f
}
if ($CurrentBatch.Count -gt 0) { Process-Batch }
Send-Status -status "MONITORING" -details "Live protection active."

while ($true) {
    Send-Heartbeat
    $id = $null
    while ($Global:ScanQueue.TryDequeue([ref]$id)) {
        $t = try { $N.GetItemFromID($id) } catch { $null }
        if ($t) {
            $fData = Parse-Forensics $t; $fp = Get-Fingerprint -item $t -ip $fData.ip
            if ($ps.Contains($fp) -or $Global:ReleasedFingerprints.Contains($fp)) { Release-Com $t; continue }
            $itemData = @{ Id=$id; Su=$t.Subject; Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$fData.body; Finger=$fp }
            $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk)
            $psi.RunspacePool = $RunspacePool; [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
        }
        Release-Com $t
    }
    if ($CurrentBatch.Count -gt 0) { Process-Batch }
    Start-Sleep -Seconds 1
}
