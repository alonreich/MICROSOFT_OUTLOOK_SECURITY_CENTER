param([string]$Mode = "", [int]$ParentPid = 0)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- GLOBAL STATE PERSISTENCE & SYNCHRONIZATION ---
if ($null -eq $Global:StdoutLock) { $Global:StdoutLock = [System.Object]::new() }
if ($null -eq $Global:DupStack) { $Global:DupStack = $null }
if ($null -eq $Global:DupHashes) { $Global:DupHashes = @{} }
if ($null -eq $Global:DupResults) { $Global:DupResults = New-Object System.Collections.Generic.List[object] }
if ($null -eq $Global:DupScannedCount) { $Global:DupScannedCount = 0 }

# --- ENTERPRISE UTILITIES ---

function Send-Structured-Message {
    param([object]$obj)
    $lockObj = if ($null -ne $Global:StdoutLock) { $Global:StdoutLock } elseif ($null -ne $StdoutLock) { $StdoutLock } else { $null }
    if ($null -eq $lockObj) {
        $lockObj = [System.Object]::new()
        $Global:StdoutLock = $lockObj
    }
    [System.Threading.Monitor]::Enter($lockObj)
    try {
        $json = $obj | ConvertTo-Json -Compress -Depth 10
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        $len = $bytes.Length
        $lenHeader = [System.BitConverter]::GetBytes([int]$len)
        if (![System.BitConverter]::IsLittleEndian) { [Array]::Reverse($lenHeader) }
        $packet = New-Object byte[] (4 + $len)
        [System.Buffer]::BlockCopy($lenHeader, 0, $packet, 0, 4)
        [System.Buffer]::BlockCopy($bytes, 0, $packet, 4, $len)
        $stream = [Console]::OpenStandardOutput()
        $stream.Write($packet, 0, $packet.Length)
        $stream.Flush()
    } finally {
        [System.Threading.Monitor]::Exit($lockObj)
    }
}

function Send-Heartbeat { 
    if ($ParentPid -gt 0 -and !(Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { 
        [void][System.Environment]::Exit(0) 
    }
    Send-Structured-Message @{type="heartbeat"; timestamp=(Get-Date -Format "yyyy-MM-dd HH:mm:ss")}
}

function Release-Com { param($O) if ($null -ne $O) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($O) } catch {} } }

function Log-Progress($m) { Send-Structured-Message @{status="INFO"; details=$m} }

function Get-SHA256 {
    param($string)
    if ([string]::IsNullOrEmpty($string)) { return "N/A" }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($string)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return [System.BitConverter]::ToString($hash).Replace("-", "").ToLower()
}

function Get-Property-Safe {
    param($item, $tag, $fallback = "N/A")
    if (!$item) { return $fallback }
    $pa = $null
    try {
        $pa = $item.PropertyAccessor
        if (!$pa) { return $fallback }
        $val = Invoke-OutlookMethod { $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/$tag") }
        if ([string]::IsNullOrWhiteSpace($val)) { return $fallback }
        return $val
    } catch { return $fallback }
    finally { Release-Com $pa }
}

function Resolve-Email {
    param($Recipient)
    if (!$Recipient) { return "Unknown" }
    
    # TRY DIRECT SMTP TAG FIRST (PR_SMTP_ADDRESS)
    $addr = Get-Property-Safe $Recipient "0x39FE001E" ""
    if ([string]::IsNullOrWhiteSpace($addr) -or $addr -match "/o=") {
        # TRY EXCHANGE USER RESOLUTION
        $ae = $null; $user = $null
        try {
            $ae = $Recipient.AddressEntry
            if ($ae) {
                try {
                    $user = $ae.GetExchangeUser()
                    if ($user) { $addr = $user.PrimarySmtpAddress }
                } catch {}
                if ([string]::IsNullOrWhiteSpace($addr)) { $addr = $ae.Name }
            }
        } catch {}
        finally {
            Release-Com $user
            Release-Com $ae
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($addr)) { try { $addr = $Recipient.Address } catch {} }
    if ($addr -match "/o=") { 
        # FINAL CLEANUP: EXTRACT NAME FROM X500
        if ($addr -match "cn=([^/]+)$") { $addr = $Matches[1] }
        else { $addr = ($addr -split "=")[-1] }
    }
    return if ([string]::IsNullOrWhiteSpace($addr)) { "Internal User" } else { $addr }
}

function Get-Fingerprint {
    param($item, $ip)
    if (!$item) { return [guid]::NewGuid().ToString() }
    
    $se = "Unknown"; $Snd = $null
    try { 
        $Snd = $item.Sender
        if ($Snd) { $se = Resolve-Email -Recipient $Snd } 
        else { $se = Get-Property-Safe $item "0x0065001E" "Unknown" } # PR_SENDER_EMAIL_ADDRESS
    } catch { 
        $se = Get-Property-Safe $item "0x0065001E" "Unknown" 
    } finally {
        Release-Com $Snd
    }
    
    $su = Get-Property-Safe $item "0x0037001E" "No Subject" # PR_SUBJECT
    $rt = "00000000000000"; try { if ($item.ReceivedTime) { $rt = $item.ReceivedTime.ToString("yyyyMMddHHmmss") } } catch {}
    
    return Get-SHA256 "$se|$su|$rt"
}

function Get-ItemDna {
    param($mailItem)
    if (!$mailItem) { return "" }
    try {
        $subject = Get-Property-Safe $mailItem "0x0037001E" "No Subject"
        $cleanSubject = ($subject -replace "^(Re:|Fwd:|FW:|RE:)\s*", "").Trim()
        $sender = "Unknown"
        $Snd = $null
        try { 
            $Snd = $mailItem.Sender
            if ($Snd) { $sender = Resolve-Email -Recipient $Snd } 
            else { $sender = Get-Property-Safe $mailItem "0x0065001E" "Unknown" }
        } catch { 
            try { $sender = $mailItem.SenderEmailAddress } catch {} 
        } finally { 
            Release-Com $Snd 
        }
        $received = "00000000000000"
        try { if ($mailItem.ReceivedTime) { $received = $mailItem.ReceivedTime.ToString("yyyyMMddHHmmss") } } catch {}
        $bodySample = Get-Property-Safe $mailItem "0x1000001E" ""
        if ($bodySample.Length -gt 300) { $bodySample = $bodySample.Substring(0, 300) }
        return Get-SHA256 "$sender|$cleanSubject|$received|$bodySample"
    } catch {
        return ""
    }
}

function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$originalEntryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [float]$score = 0, [string]$body = "", [bool]$unread = $false, [string]$scanType = "", [string]$to = "", [string]$cc = "", [string]$fingerprint = "", [string]$timestamp = "", [int]$count = 0, [int]$total = 0, [string]$currentFolder = "")
    $h = ""; if (![string]::IsNullOrEmpty($fullHeaders)) { try { $h = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($fullHeaders)) } catch {} }
    $b = ""; if (![string]::IsNullOrEmpty($body)) { try { $b = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($body)) } catch {} }
    $ts = if ([string]::IsNullOrEmpty($timestamp)) { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") } else { $timestamp }
    Send-Structured-Message @{
        timestamp=$ts; status=$status; details=$details; verdict=$verdict; action=$action;
        entryId=$entryId; originalEntryId=$originalEntryId; tier=$tier; phase=$phase; sender=$sender;
        ip=$ip; domain=$domain; originalFolder=$originalFolder; fullHeaders=$h; score=$score;
        body=$b; unread=$unread; scanType=$scanType; to=$to; cc=$cc; fingerprint=$fingerprint;
        count=$count; total=$total; currentFolder=$currentFolder
    }
}

$Global:OutlookComMutex = $null
function Get-OutlookComMutex {
    if ($null -ne $Global:OutlookComMutex) { return $Global:OutlookComMutex }
    $mutexName = "Global\MOS_Outlook_COM_Lock"
    try {
        $createdNew = $false
        $m = New-Object System.Threading.Mutex($false, $mutexName, [ref]$createdNew)
        $Global:OutlookComMutex = $m
        return $m
    } catch {
        try {
            $Global:OutlookComMutex = [System.Threading.Mutex]::OpenExisting($mutexName)
            return $Global:OutlookComMutex
        } catch {
            $localName = "Local\MOS_Outlook_COM_Lock"
            try {
                $createdLocal = $false
                $Global:OutlookComMutex = New-Object System.Threading.Mutex($false, $localName, [ref]$createdLocal)
                return $Global:OutlookComMutex
            } catch {
                return $null
            }
        }
    }
}

$Global:CircuitOpen = $false
$Global:CircuitResetTime = [DateTime]::MinValue
$Global:BackoffIntervals = @(50, 150, 500, 1000)

function Invoke-OutlookMethod {
    param($ScriptBlock, $MaxRetries = 5)
    if ($Global:CircuitOpen) {
        if ([DateTime]::Now -lt $Global:CircuitResetTime) { 
            throw "Circuit Open: Outlook overloaded or unresponsive." 
        }
        $Global:CircuitOpen = $false
    }
    
    $mutex = Get-OutlookComMutex
    $hasLock = $false
    try {
        if ($null -ne $mutex) {
            try {
                $hasLock = $mutex.WaitOne(30000)
            } catch [System.Threading.AbandonedMutexException] {
                $hasLock = $true
            }
        }
        
        $retryCount = 0
        while ($retryCount -lt $MaxRetries) {
            try { 
                return & $ScriptBlock 
            }
            catch [System.Runtime.InteropServices.COMException] {
                $code = $_.Exception.ErrorCode
                # 0x8001010A (-2147417846 RPC_E_SERVERCALL_RETRYLATER), 0x80010001 (-2147418111 RPC_E_CALL_REJECTED), 0x800401EC (-2147220948 MAPI_E_BUSY)
                if ($code -eq -2147418111 -or $code -eq -2147417846 -or $code -eq -2147220948) {
                    $retryCount++
                    if ($retryCount -ge $MaxRetries) { break }
                    $baseDelay = if ($retryCount -le $Global:BackoffIntervals.Length) { 
                        $Global:BackoffIntervals[$retryCount - 1] 
                    } else { 1000 }
                    $jitter = Get-Random -Minimum 10 -Maximum 50
                    Start-Sleep -Milliseconds ($baseDelay + $jitter)
                }
                elseif ($code -eq -2147221164) { # 0x80040154 REGDB_E_CLASSNOTREG
                    Log-Progress "Outlook COM Error: Class not registered (0x80040154)."
                    throw $_
                }
                else {
                    throw $_
                }
            }
            catch {
                throw $_
            }
        }
        
        $Global:CircuitOpen = $true
        $Global:CircuitResetTime = [DateTime]::Now.AddSeconds(30)
        throw "Outlook busy timeout (RPC_E_SERVERCALL_RETRYLATER). Circuit broken for 30s."
    }
    finally {
        if ($hasLock -and $null -ne $mutex) {
            try { $mutex.ReleaseMutex() } catch {}
        }
    }
}

function Get-TargetFolder-Safe {
    param($item, [int]$folderType)
    $parent = $null; $store = $null; $target = $null
    try {
        $parent = $item.Parent
        if ($parent) {
            $store = $parent.Store
            if ($store) {
                $target = Invoke-OutlookMethod { $store.GetDefaultFolder($folderType) }
                return $target
            }
        }
        return $null
    } finally {
        Release-Com $store
        Release-Com $parent
    }
}

function Get-Outlook {
    $attempts = 0
    while ($attempts -lt 3) {
        try { 
            $obj = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") 
            if ($obj) { return $obj }
        } catch { 
            try { 
                $obj = New-Object -ComObject Outlook.Application 
                if ($obj) { return $obj }
            } catch { 
                [Console]::Error.WriteLine("Outlook Connection Attempt $($attempts + 1) failed: $($_.Exception.Message)")
            } 
        }
        $attempts++
        Start-Sleep -Seconds 2
    }
    [Console]::Error.WriteLine("CRITICAL: All 3 attempts to connect to Outlook failed.")
    return $null
}

function Init-Exclusions {
    param($Namespace)
    if ($null -eq $Global:ExcludedFolderIds) { $Global:ExcludedFolderIds = New-Object System.Collections.Generic.HashSet[string] }
    $folderIds = @(3, 4, 5, 16, 23)
    $stores = $null
    try {
        $stores = $Namespace.Stores
        foreach ($S in $stores) {
            foreach ($id in $folderIds) {
                $f = $null
                try { 
                    $f = Invoke-OutlookMethod { $S.GetDefaultFolder($id) }
                    if ($f) { [void]$Global:ExcludedFolderIds.Add($f.EntryID) } 
                } catch {}
                finally { Release-Com $f }
            }
            Release-Com $S
        }
    } catch {}
    finally { Release-Com $stores }
}

function Parse-Forensics {
    param($item)
    $headers = Get-Property-Safe $item "0x007D001E" "" # PR_TRANSPORT_MESSAGE_HEADERS
    $senderIp = "N/A"
    
    if ($headers) {
        $received = $headers -split "`r`n" | Where-Object { $_ -match "^Received:" }
        foreach ($line in $received) {
            if ($line -match "\[(?<ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\]" -or $line -match "\((?<ip>\d{1,3}\.\d{1,3}\.\d{1,3})\)") { 
                $cand = $Matches['ip']
                if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") { $senderIp = $cand; break }
            }
        }
    }
    
    $from = "Unknown"
    $Snd = $null
    try { 
        $Snd = $item.Sender
        if ($Snd) { $from = Resolve-Email -Recipient $Snd } 
        else { $from = Get-Property-Safe $item "0x0042001E" "Unknown" } # PR_SENDER_NAME
    } catch { 
        $from = Get-Property-Safe $item "0x0042001E" "Unknown" 
    } finally {
        Release-Com $Snd
    }
    
    $body = Get-Property-Safe $item "0x1000001E" "" # PR_BODY
    if ([string]::IsNullOrWhiteSpace($body)) { $body = "Internal Message: No plain text body content." }
    
    $atts = New-Object System.Collections.Generic.List[object]
    $AttsObj = $null
    try {
        $AttsObj = $item.Attachments
        if ($AttsObj -and $AttsObj.Count -gt 0) {
            foreach ($at in $AttsObj) {
                $hash = "N/A"; try { $hash = Get-SHA256 "$($at.FileName)|$($at.Size)" } catch {}
                [void]$atts.Add(@{ name=$at.FileName; hash=$hash; size=$at.Size })
                Release-Com $at
            }
        }
    } catch {}
    finally {
        Release-Com $AttsObj
    }
    
    return @{ ip=$senderIp; headers=$headers; from=$from; body=$body; attachments=$atts }
}

function Robust-Move {
    param($item, $targetFolder)
    if (!$item -or !$targetFolder) { return $null }
    try {
        $origUnread = $item.UnRead
        $m = Invoke-OutlookMethod { $item.Move($targetFolder) }
        if ($null -ne $m) { 
            Invoke-OutlookMethod {
                $m.UnRead = $origUnread
                $m.Save()
            }
            return $m 
        }
    } catch { Log-Progress "Robust-Move Error: $($_.Exception.Message)" }
    return $null
}

# --- WORKER MODE ---

if ($Mode -eq "Worker") {
    $O = Get-Outlook; if (!$O) { exit 1 }
    $N = $null
    try {
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
                if (Test-Path $pauseFile) { Remove-Item $pauseFile -Force }
                Log-Progress "Worker: Starting Duplicate Email Discovery phase..."
                
                if ($null -eq $Global:DupStack) {
                    $Global:DupStack = New-Object System.Collections.Generic.Stack[object]
                    $Global:DupHashes = @{}
                    $Global:DupResults = New-Object System.Collections.Generic.List[object]
                    $Global:DupScannedCount = 0
                    
                    $stores = $null
                    try {
                        $stores = $N.Stores
                        foreach ($S in $stores) {
                            $root = $null; $storeSize = 0
                            try { $root = $S.GetRootFolder(); $storeSize = $root.Size } catch {} finally { Release-Com $root }
                            
                            Send-Structured-Message @{type="duplicate-update"; status="StoreStart"; store=$S.DisplayName; size=$storeSize}
                            
                            $storeItemsCount = 0
                            $foldersToProcess = New-Object System.Collections.Generic.Stack[object]
                            @(6, 5) | ForEach-Object { 
                                $df = $null
                                try { 
                                    $df = Invoke-OutlookMethod { $S.GetDefaultFolder($_) }
                                    if ($df) { $foldersToProcess.Push($df) } 
                                } catch {}
                            }
                            
                            while ($foldersToProcess.Count -gt 0) {
                                $f = $foldersToProcess.Pop()
                                $fItems = $null; $fSubs = $null
                                try { 
                                    $fItems = $f.Items
                                    $storeItemsCount += $fItems.Count 
                                    $Global:DupStack.Push(@{
                                        FolderId = $f.EntryID
                                        StoreId = $S.StoreID
                                        StoreName = $S.DisplayName
                                        FolderName = $f.Name
                                        DefaultItemType = $f.DefaultItemType
                                    })
                                    $fSubs = $f.Folders
                                    foreach ($sub in $fSubs) { $foldersToProcess.Push($sub) }
                                } catch {} 
                                finally { 
                                    Release-Com $fItems
                                    Release-Com $fSubs
                                    Release-Com $f
                                }
                            }
                            Send-Structured-Message @{type="duplicate-update"; status="StoreMeta"; store=$S.DisplayName; totalItems=$storeItemsCount}
                            Release-Com $S
                        }
                    } catch { 
                        Log-Progress "Worker: Error during store indexing: $($_.Exception.Message)" 
                    } finally {
                        Release-Com $stores
                    }
                }

                $currentStoreName = ""
                $currentStoreScannedItems = 0
                $currentStoreScannedSize = 0

                while ($Global:DupStack.Count -gt 0) {
                    if (Test-Path $pauseFile) {
                        Send-Structured-Message @{type="duplicate-update"; status="Paused"}
                        break
                    }

                    $entry = $Global:DupStack.Pop()
                    $folderId = $entry.FolderId
                    $storeId = $entry.StoreId
                    $storeName = $entry.StoreName
                    $folderName = $entry.FolderName
                    
                    if ($storeName -ne $currentStoreName) {
                        if ($currentStoreName -ne "") {
                            Send-Structured-Message @{type="duplicate-update"; status="StoreFinish"; store=$currentStoreName; found=$Global:DupResults.Count; scanned=$currentStoreScannedItems; size=$currentStoreScannedSize}
                        }
                        $currentStoreName = $storeName
                        $currentStoreScannedItems = 0
                        $currentStoreScannedSize = 0
                    }

                    if ($entry.DefaultItemType -eq 0) {
                        $f = $null
                        $items = $null
                        try {
                            $f = Invoke-OutlookMethod { $N.GetFolderFromID($folderId, $storeId) }
                            if ($f) {
                                $items = $f.Items
                                $totalInFolder = try { $items.Count } catch { 0 }
                                $folderScanned = 0
                                foreach ($t in $items) {
                                    try {
                                        $currentStoreScannedItems++
                                        $Global:DupScannedCount++
                                        $folderScanned++
                                        $itemSize = try { $t.Size } catch { 0 }
                                        $currentStoreScannedSize += $itemSize
                                        
                                        $subject = Get-Property-Safe $t "0x0037001E" "No Subject"
                                        
                                        Send-Structured-Message @{
                                            type = "duplicate-update"
                                            status = "Scanned"
                                            scanned = $Global:DupScannedCount
                                            found = $Global:DupResults.Count
                                            folderProgress = "$folderScanned/$totalInFolder"
                                            currentFolder = $folderName
                                            currentItem = $subject
                                            store = $storeName
                                            storeScanned = $currentStoreScannedItems
                                            storeScannedSize = $currentStoreScannedSize
                                        }

                                        $sender = "Unknown"
                                        $Snd = $null
                                        try { 
                                            $Snd = $t.Sender
                                            $sender = Resolve-Email -Recipient $Snd 
                                        } catch { 
                                            try { $sender = $t.SenderEmailAddress } catch {} 
                                        } finally { 
                                            Release-Com $Snd 
                                        }
                                        $cleanSubject = ($subject -replace "^(Re:|Fwd:|FW:|RE:)\s*", "").Trim()
                                        $received = "00000000000000"
                                        try { if ($t.ReceivedTime) { $received = $t.ReceivedTime.ToString("yyyyMMddHHmmss") } } catch {}
                                        $bodySample = Get-Property-Safe $t "0x1000001E" ""
                                        if ($bodySample.Length -gt 300) { $bodySample = $bodySample.Substring(0, 300) }
                                        
                                        $dna = Get-SHA256 "$sender|$cleanSubject|$received|$bodySample"
                                        $itemObj = @{ entryId=$t.EntryID; subject=$subject; sender=$sender; timestamp=$t.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"); size=$itemSize; folder=$folderName; store=$storeName; dna=$dna }
                                        
                                        if (!$Global:DupHashes.ContainsKey($dna)) { 
                                            $Global:DupHashes[$dna] = $itemObj 
                                        }
                                        else {
                                            $survivor = $Global:DupHashes[$dna]
                                            if ($itemObj.size -gt $survivor.size) { 
                                                $dupItem = $survivor
                                                $dupItem.survivorId = $itemObj.entryId
                                                $dupItem.survivorFolder = $itemObj.folder
                                                $dupItem.survivorStore = $itemObj.store
                                                [void]$Global:DupResults.Add($dupItem)
                                                $Global:DupHashes[$dna] = $itemObj 
                                                foreach ($res in $Global:DupResults) {
                                                    if ($res.dna -eq $dna) {
                                                        $res.survivorId = $itemObj.entryId
                                                        $res.survivorFolder = $itemObj.folder
                                                        $res.survivorStore = $itemObj.store
                                                    }
                                                }
                                            }
                                            else { 
                                                $itemObj.survivorId = $survivor.entryId
                                                $itemObj.survivorFolder = $survivor.folder
                                                $itemObj.survivorStore = $survivor.store
                                                [void]$Global:DupResults.Add($itemObj) 
                                            }
                                            Send-Structured-Message @{type="duplicate-update"; status="Found"; count=$Global:DupResults.Count; current=$itemObj.subject; scanned=$Global:DupScannedCount}
                                        }
                                    } catch {} finally { Release-Com $t }
                                    if ($folderScanned % 100 -eq 0) { Start-Sleep -Milliseconds 10 }
                                }
                            }
                        } catch {
                            Log-Progress "Worker: Error scanning folder $folderName : $($_.Exception.Message)"
                        } finally {
                            Release-Com $items
                            Release-Com $f
                        }
                    }
                }

                if ($Global:DupStack.Count -eq 0) {
                    if ($currentStoreName -ne "") {
                        Send-Structured-Message @{type="duplicate-update"; status="StoreFinish"; store=$currentStoreName; found=$Global:DupResults.Count; scanned=$currentStoreScannedItems; size=$currentStoreScannedSize}
                    }
                    Send-Structured-Message @{type="duplicate-update"; status="Finished"; items=$Global:DupResults}
                    $Global:DupStack = $null
                }
            }
            elseif ($Action -eq "ResetDuplicateStack" -or $Action -eq "ResetStack") {
                $Global:DupStack = $null
                $Global:DupHashes = @{}
                $Global:DupResults = New-Object System.Collections.Generic.List[object]
                $Global:DupScannedCount = 0
                Log-Progress "Worker: Duplicate detection engine reset."
            }
            elseif ($Action -eq "CleanDuplicates" -or ($Action -eq "Delete" -and ($Ex.items -or $Ex.isDuplicateCleanup))) {
                $targetItems = @()
                if ($Ex.items) {
                    $targetItems = @($Ex.items)
                } elseif ($Ex.entryIds) {
                    foreach ($eid in $Ex.entryIds) {
                        $found = $null
                        if ($null -ne $Global:DupResults) {
                            $found = $Global:DupResults | Where-Object { $_.entryId -eq $eid } | Select-Object -First 1
                        }
                        if ($found) { $targetItems += $found } else { $targetItems += @{ entryId = $eid } }
                    }
                } elseif ($Ex.entryId) {
                    $found = $null
                    if ($null -ne $Global:DupResults) {
                        $found = $Global:DupResults | Where-Object { $_.entryId -eq $Ex.entryId } | Select-Object -First 1
                    }
                    if ($found) { $targetItems += $found } else { $targetItems += @{ entryId = $Ex.entryId } }
                }

                $mode = if ($Ex.mode) { [string]$Ex.mode.ToLower() } else { "safe" }
                $successCount = 0
                $skippedCount = 0
                $errors = New-Object System.Collections.Generic.List[object]
                $skippedItems = New-Object System.Collections.Generic.List[object]

                Log-Progress "Worker: Starting Pre-Flight Survivor Validation and Duplicate Cleanup (Mode: $mode, Items: $($targetItems.Count))..."

                foreach ($entry in $targetItems) {
                    $dupId = $entry.entryId
                    if ([string]::IsNullOrEmpty($dupId)) { continue }

                    $survivorId = $entry.survivorId
                    $expectedDna = $entry.dna

                    if ([string]::IsNullOrEmpty($survivorId) -and $null -ne $Global:DupResults) {
                        $match = $Global:DupResults | Where-Object { $_.entryId -eq $dupId } | Select-Object -First 1
                        if ($match) {
                            $survivorId = $match.survivorId
                            if ([string]::IsNullOrEmpty($expectedDna)) { $expectedDna = $match.dna }
                        }
                    }

                    # Pre-flight Check 1: Survivor record exists and is distinct
                    if ([string]::IsNullOrEmpty($survivorId)) {
                        $skippedCount++
                        $msg = "Survivor missing or modified"
                        $skippedItems.Add(@{ entryId=$dupId; status="Skipped"; reason=$msg; subject=$entry.subject })
                        Log-Progress "SAFETY GUARD: No survivor mapped for duplicate $dupId. Skipping deletion to prevent data loss."
                        continue
                    }

                    if ($dupId -eq $survivorId) {
                        $skippedCount++
                        $msg = "Target item is the survivor itself"
                        $skippedItems.Add(@{ entryId=$dupId; status="Skipped"; reason=$msg; subject=$entry.subject })
                        Log-Progress "SAFETY GUARD: Target $dupId is the survivor. Deletion halted."
                        continue
                    }

                    $survivorItem = $null
                    $targetItem = $null
                    $survivorParent = $null
                    $survivorStore = $null
                    $deletedFolder = $null

                    try {
                        # Pre-flight Check 2: Fetch survivor and verify existence in Outlook
                        $survivorItem = Invoke-OutlookMethod { $N.GetItemFromID($survivorId) }
                        if (!$survivorItem) {
                            $skippedCount++
                            $msg = "Survivor missing or modified"
                            $skippedItems.Add(@{ entryId=$dupId; survivorId=$survivorId; status="Skipped"; reason=$msg; subject=$entry.subject })
                            Log-Progress "SAFETY GUARD: Survivor item $survivorId does not exist in Outlook. Skipping duplicate $dupId."
                            continue
                        }

                        # Pre-flight Check 3: Verify survivor is not in Deleted Items (olFolderDeletedItems = 3)
                        $isSurvivorInDeleted = $false
                        try {
                            $survivorParent = $survivorItem.Parent
                            $survivorStore = $survivorParent.Store
                            $deletedFolder = Invoke-OutlookMethod { $survivorStore.GetDefaultFolder(3) }
                            if ($survivorParent -and $deletedFolder -and ($survivorParent.EntryID -eq $deletedFolder.EntryID)) {
                                $isSurvivorInDeleted = $true
                            }
                        } catch {}
                        if ($isSurvivorInDeleted) {
                            $skippedCount++
                            $msg = "Survivor is in Deleted Items folder"
                            $skippedItems.Add(@{ entryId=$dupId; survivorId=$survivorId; status="Skipped"; reason=$msg; subject=$entry.subject })
                            Log-Progress "SAFETY GUARD: Survivor $survivorId is in Deleted Items. Preserving duplicate $dupId."
                            continue
                        }

                        # Pre-flight Check 4: Survivor DNA must match expected DNA
                        $survivorDna = Get-ItemDna $survivorItem
                        if ([string]::IsNullOrEmpty($survivorDna) -or (![string]::IsNullOrEmpty($expectedDna) -and $survivorDna -ne $expectedDna)) {
                            $skippedCount++
                            $msg = "Survivor missing or modified"
                            $skippedItems.Add(@{ entryId=$dupId; survivorId=$survivorId; status="Skipped"; reason=$msg; subject=$entry.subject })
                            Log-Progress "SAFETY GUARD: Survivor $survivorId DNA mismatch (Expected: $expectedDna, Actual: $survivorDna). Preserving duplicate $dupId."
                            continue
                        }

                        # Pre-flight Check 5: Fetch duplicate target item and verify distinct + matching DNA
                        $targetItem = Invoke-OutlookMethod { $N.GetItemFromID($dupId) }
                        if (!$targetItem) {
                            $skippedCount++
                            $msg = "Target duplicate item not found"
                            $skippedItems.Add(@{ entryId=$dupId; status="Skipped"; reason=$msg; subject=$entry.subject })
                            continue
                        }

                        if ($targetItem.EntryID -eq $survivorItem.EntryID) {
                            $skippedCount++
                            $msg = "Target item resolved to Survivor"
                            $skippedItems.Add(@{ entryId=$dupId; status="Skipped"; reason=$msg; subject=$entry.subject })
                            Log-Progress "SAFETY GUARD: Target resolved to Survivor ($($survivorItem.EntryID)). Skipping."
                            continue
                        }

                        $targetDna = Get-ItemDna $targetItem
                        if ([string]::IsNullOrEmpty($targetDna) -or (![string]::IsNullOrEmpty($expectedDna) -and $targetDna -ne $expectedDna)) {
                            $skippedCount++
                            $msg = "Target duplicate DNA mismatch"
                            $skippedItems.Add(@{ entryId=$dupId; status="Skipped"; reason=$msg; subject=$entry.subject })
                            Log-Progress "SAFETY GUARD: Target $dupId DNA mismatch. Skipping."
                            continue
                        }

                        # Execution Phase: Safe Move vs Permanent Purge
                        if ($mode -eq "purge" -or $mode -eq "permanent") {
                            Invoke-OutlookMethod { $targetItem.Delete() }
                            $successCount++
                            Log-Progress "Duplicate Cleanup [PURGE]: Permanently deleted duplicate '$($entry.subject)' (ID: $dupId, Sender: $($entry.sender), TS: $($entry.timestamp)). Preserved Survivor: $survivorId"
                        } else {
                            # Safe Mode: Move to Deleted Items via Robust-Move
                            $targetStore = $null
                            $destFolder = $null
                            try {
                                $targetStore = $targetItem.Parent.Store
                                $destFolder = Invoke-OutlookMethod { $targetStore.GetDefaultFolder(3) } # olFolderDeletedItems = 3
                                if ($destFolder) {
                                    $moved = Robust-Move -item $targetItem -targetFolder $destFolder
                                    if ($moved) {
                                        $successCount++
                                        Log-Progress "Duplicate Cleanup [SAFE]: Moved duplicate '$($entry.subject)' to Deleted Items (ID: $dupId, Sender: $($entry.sender), TS: $($entry.timestamp)). Preserved Survivor: $survivorId"
                                        Release-Com $moved
                                    } else {
                                        $skippedCount++
                                        $errors.Add(@{ entryId=$dupId; error="Robust-Move returned null" })
                                    }
                                } else {
                                    $skippedCount++
                                    $errors.Add(@{ entryId=$dupId; error="Could not locate Deleted Items folder" })
                                }
                            } finally {
                                Release-Com $destFolder
                                Release-Com $targetStore
                            }
                        }
                    } catch {
                        $skippedCount++
                        $errors.Add(@{ entryId=$dupId; error=$_.Exception.Message })
                        Log-Progress "Worker: Error processing duplicate $dupId : $($_.Exception.Message)"
                    } finally {
                        Release-Com $deletedFolder
                        Release-Com $survivorStore
                        Release-Com $survivorParent
                        Release-Com $targetItem
                        Release-Com $survivorItem
                    }
                }

                $report = @{
                    successCount = $successCount
                    skippedCount = $skippedCount
                    count = $successCount
                    mode = $mode
                    errors = $errors
                    skipped = $skippedItems
                }

                Send-Structured-Message @{type="delete-summary"; count=$successCount; skipped=$skippedCount; mode=$mode}
                if ($Ex.rid) {
                    Send-Structured-Message @{type="cmd-response"; rid=$Ex.rid; data=$report}
                }
            }
            elseif ($Action -eq "Delete") {
                $count = 0
                $ids = @()
                if ($Ex.entryIds) { $ids = $Ex.entryIds }
                elseif ($Ex.entryId) { $ids = @($Ex.entryId) }
                
                foreach ($id in $ids) {
                    $item = $null
                    try {
                        $item = Invoke-OutlookMethod { $N.GetItemFromID($id) }
                        if ($item) { 
                            Invoke-OutlookMethod { $item.Delete() }
                            $count++ 
                        }
                    } catch {} 
                    finally { Release-Com $item }
                }
                Send-Structured-Message @{type="delete-summary"; count=$count}
                if ($Ex.rid) {
                    Send-Structured-Message @{type="cmd-response"; rid=$Ex.rid; data=@{count=$count; success=$true}}
                }
            }
            elseif ($Action -eq "Release") {
                $id = if ($Ex.entryId) { $Ex.entryId } elseif ($Ex.data -and $Ex.data.entryId) { $Ex.data.entryId } else { $null }
                $targetFolderType = if ($Ex.targetFolder) { [int]$Ex.targetFolder } elseif ($Ex.data -and $Ex.data.targetFolder) { [int]$Ex.data.targetFolder } else { 6 } # olFolderInbox = 6
                $fp = if ($Ex.fingerprint) { $Ex.fingerprint } elseif ($Ex.data -and $Ex.data.fingerprint) { $Ex.data.fingerprint } else { "" }
                
                $item = $null
                $targetFolder = $null
                $m = $null
                $success = $false
                $newId = ""
                $errMsg = ""

                try {
                    if ([string]::IsNullOrEmpty($id)) {
                        throw "Invalid parameter: entryId is required"
                    }
                    $item = Invoke-OutlookMethod { $N.GetItemFromID($id) }
                    if (!$item) {
                        throw "Item not found"
                    }
                    if ([string]::IsNullOrEmpty($fp)) {
                        $fData = Parse-Forensics $item
                        $fp = Get-Fingerprint -item $item -ip $fData.ip
                    }
                    if (![string]::IsNullOrEmpty($fp)) {
                        [void]$Global:ReleasedFingerprints.Add($fp)
                        Send-Structured-Message @{ type="store-update"; key="releasedFingerprints"; value=$fp }
                    }

                    $targetFolder = Get-TargetFolder-Safe $item $targetFolderType
                    if (!$targetFolder) {
                        $targetFolder = Invoke-OutlookMethod { $N.GetDefaultFolder($targetFolderType) }
                    }
                    if (!$targetFolder) {
                        throw "Target folder $targetFolderType not available"
                    }

                    $m = Robust-Move $item $targetFolder
                    if ($m) {
                        $success = $true
                        $newId = $m.EntryID
                    } else {
                        throw "Move operation failed"
                    }
                }
                catch [System.Runtime.InteropServices.COMException] {
                    $errMsg = if ($_.Exception.ErrorCode -eq -2147221233 -or $_.Exception.ErrorCode -eq 0x8004010F -or $_.Exception.Message -match "not found|could not open") { "Item not found" } else { $_.Exception.Message }
                }
                catch {
                    $errMsg = $_.Exception.Message
                }
                finally {
                    Release-Com $m
                    Release-Com $targetFolder
                    Release-Com $item
                }

                if ($Ex.rid) {
                    if ($success) {
                        Send-Structured-Message @{ type="cmd-response"; rid=$Ex.rid; success=$true; newEntryId=$newId; fingerprint=$fp; data=@{ success=$true; newEntryId=$newId; fingerprint=$fp } }
                    } else {
                        Send-Structured-Message @{ type="cmd-response"; rid=$Ex.rid; success=$false; error=$errMsg; data=@{ success=$false; error=$errMsg } }
                    }
                }
                Send-Structured-Message @{ type="release-complete"; entryId=$newId; originalEntryId=$id; success=$success; fingerprint=$fp }
            }
            elseif ($Action -eq "Quarantine") {
                $id = if ($Ex.entryId) { $Ex.entryId } elseif ($Ex.data -and $Ex.data.entryId) { $Ex.data.entryId } else { $null }
                $targetFolderType = if ($Ex.targetFolder) { [int]$Ex.targetFolder } elseif ($Ex.data -and $Ex.data.targetFolder) { [int]$Ex.data.targetFolder } else { 23 } # olFolderJunk = 23
                
                $item = $null
                $targetFolder = $null
                $m = $null
                $success = $false
                $newId = ""
                $errMsg = ""

                try {
                    if ([string]::IsNullOrEmpty($id)) {
                        throw "Invalid parameter: entryId is required"
                    }
                    $item = Invoke-OutlookMethod { $N.GetItemFromID($id) }
                    if (!$item) {
                        throw "Item not found"
                    }

                    $targetFolder = Get-TargetFolder-Safe $item $targetFolderType
                    if (!$targetFolder -and $targetFolderType -eq 23) {
                        $targetFolder = Get-TargetFolder-Safe $item 3 # Fallback to Deleted Items
                    }
                    if (!$targetFolder) {
                        $targetFolder = Invoke-OutlookMethod { $N.GetDefaultFolder($targetFolderType) }
                    }
                    if (!$targetFolder) {
                        throw "Target quarantine folder not available"
                    }

                    $m = Robust-Move $item $targetFolder
                    if ($m) {
                        $success = $true
                        $newId = $m.EntryID
                    } else {
                        throw "Move operation failed"
                    }
                }
                catch [System.Runtime.InteropServices.COMException] {
                    $errMsg = if ($_.Exception.ErrorCode -eq -2147221233 -or $_.Exception.ErrorCode -eq 0x8004010F -or $_.Exception.Message -match "not found|could not open") { "Item not found" } else { $_.Exception.Message }
                }
                catch {
                    $errMsg = $_.Exception.Message
                }
                finally {
                    Release-Com $m
                    Release-Com $targetFolder
                    Release-Com $item
                }

                if ($Ex.rid) {
                    if ($success) {
                        Send-Structured-Message @{ type="cmd-response"; rid=$Ex.rid; success=$true; newEntryId=$newId; data=@{ success=$true; newEntryId=$newId } }
                    } else {
                        Send-Structured-Message @{ type="cmd-response"; rid=$Ex.rid; success=$false; error=$errMsg; data=@{ success=$false; error=$errMsg } }
                    }
                }
                Send-Structured-Message @{ type="quarantine-complete"; entryId=$newId; originalEntryId=$id; success=$success }
            }
            elseif ($Action -eq "Check-Existence") {
                $itemsList = if ($Ex.items) { $Ex.items } elseif ($Ex.data -and $Ex.data.items) { $Ex.data.items } else { @() }
                $removedList = New-Object System.Collections.Generic.List[object]
                foreach ($entry in $itemsList) {
                    $id = if ($entry.entryId) { $entry.entryId } else { $entry }
                    if ([string]::IsNullOrEmpty($id)) { continue }
                    $item = $null
                    $exists = $false
                    try {
                        $item = Invoke-OutlookMethod { $N.GetItemFromID($id) }
                        if ($null -ne $item) { $exists = $true }
                    } catch {
                        $exists = $false
                    } finally {
                        Release-Com $item
                    }
                    if (!$exists) {
                        [void]$removedList.Add(@{ entryId=$id })
                    }
                }
                if ($Ex.rid) {
                    Send-Structured-Message @{ type="cmd-response"; rid=$Ex.rid; success=$true; removed=$removedList; data=@{ removed=$removedList } }
                }
            }
            elseif ($Action -eq "CloudVirusScan") {
                $id = if ($Ex.entryId) { $Ex.entryId } elseif ($Ex.data -and $Ex.data.entryId) { $Ex.data.entryId } else { $null }
                $til = if ($Ex.threatIntelligenceLevel -ne $null) { $Ex.threatIntelligenceLevel } elseif ($Ex.data -and $Ex.data.threatIntelligenceLevel -ne $null) { $Ex.data.threatIntelligenceLevel } else { 1 }
                $vtKey = if ($Ex.vtKey) { $Ex.vtKey } elseif ($Ex.data -and $Ex.data.vtKey) { $Ex.data.vtKey } else { "" }
                
                $item = $null
                $success = $false
                $vtReport = $null
                $errMsg = ""

                try {
                    if ([string]::IsNullOrEmpty($id)) {
                        throw "Invalid parameter: entryId is required"
                    }
                    $item = Invoke-OutlookMethod { $N.GetItemFromID($id) }
                    if (!$item) {
                        throw "Item not found"
                    }

                    $fData = Parse-Forensics $item
                    $bodyHash = if (![string]::IsNullOrEmpty($fData.body)) { Get-SHA256 $fData.body } else { "N/A" }
                    $subject = Get-Property-Safe $item "0x0037001E" "No Subject"

                    $hashesToAudit = New-Object System.Collections.Generic.List[object]
                    if ($bodyHash -ne "N/A") {
                        [void]$hashesToAudit.Add(@{ type="body"; name="Message Body"; hash=$bodyHash })
                    }
                    foreach ($at in $fData.attachments) {
                        if ($at.hash -and $at.hash -ne "N/A") {
                            [void]$hashesToAudit.Add(@{ type="attachment"; name=$at.name; hash=$at.hash; size=$at.size })
                        }
                    }

                    $auditResults = New-Object System.Collections.Generic.List[object]
                    $totalThreats = 0
                    $vtQueried = $false

                    foreach ($hObj in $hashesToAudit) {
                        $vtStatus = "Clean / Undetected"
                        $maliciousVotes = 0
                        $suspiciousVotes = 0
                        $harmlessVotes = 0
                        
                        if (![string]::IsNullOrEmpty($vtKey) -and $vtKey -ne "MASKED_FOR_SECURITY" -and $vtKey.Length -ge 16) {
                            try {
                                $vtQueried = $true
                                $vtUrl = "https://www.virustotal.com/api/v3/files/$($hObj.hash)"
                                $resp = Invoke-RestMethod -Uri $vtUrl -Headers @{ "x-apikey" = $vtKey } -Method Get -TimeoutSec 5 -ErrorAction Stop
                                if ($resp -and $resp.data -and $resp.data.attributes -and $resp.data.attributes.last_analysis_stats) {
                                    $stats = $resp.data.attributes.last_analysis_stats
                                    $maliciousVotes = [int]$stats.malicious
                                    $suspiciousVotes = [int]$stats.suspicious
                                    $harmlessVotes = [int]$stats.harmless
                                    if ($maliciousVotes -gt 0) {
                                        $vtStatus = "MALICIOUS ($maliciousVotes vendor alerts)"
                                        $totalThreats++
                                    } elseif ($suspiciousVotes -gt 0) {
                                        $vtStatus = "SUSPICIOUS ($suspiciousVotes vendor alerts)"
                                    } else {
                                        $vtStatus = "Clean ($harmlessVotes clean verdicts)"
                                    }
                                }
                            }
                            catch {
                                if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 404) {
                                    $vtStatus = "Not seen in VirusTotal database"
                                } else {
                                    $vtStatus = "API query deferred ($($_.Exception.Message))"
                                }
                            }
                        } else {
                            $vtStatus = "Local Hash Computed (No VT API key configured)"
                        }

                        [void]$auditResults.Add(@{
                            name = $hObj.name
                            type = $hObj.type
                            hash = $hObj.hash
                            status = $vtStatus
                            malicious = $maliciousVotes
                            suspicious = $suspiciousVotes
                        })
                    }

                    $success = $true
                    $statusSummary = if ($totalThreats -gt 0) { "Threats Detected ($totalThreats)" } elseif ($vtQueried) { "Clean (Verified with VirusTotal)" } else { "Audited (Local Signatures)" }

                    $vtReport = @{
                        entryId = $id
                        subject = $subject
                        sender = $fData.from
                        ip = $fData.ip
                        bodyHash = $bodyHash
                        threatIntelligenceLevel = $til
                        vtQueried = $vtQueried
                        threatsFound = $totalThreats
                        status = $statusSummary
                        items = $auditResults
                    }
                }
                catch [System.Runtime.InteropServices.COMException] {
                    $errMsg = if ($_.Exception.ErrorCode -eq -2147221233 -or $_.Exception.ErrorCode -eq 0x8004010F -or $_.Exception.Message -match "not found|could not open") { "Item not found" } else { $_.Exception.Message }
                }
                catch {
                    $errMsg = $_.Exception.Message
                }
                finally {
                    Release-Com $item
                }

                if ($Ex.rid) {
                    if ($success) {
                        Send-Structured-Message @{ type="cmd-response"; rid=$Ex.rid; success=$true; data=$vtReport }
                    } else {
                        Send-Structured-Message @{ type="cmd-response"; rid=$Ex.rid; success=$false; error=$errMsg; data=@{ success=$false; error=$errMsg } }
                    }
                }
            }
        }
    } finally {
        Release-Com $N
        Release-Com $O
        if ($null -ne $Global:OutlookComMutex) {
            try { $Global:OutlookComMutex.Dispose() } catch {}
        }
    }
    exit 0
}

# --- SCANNER MODE ---

$O = Get-Outlook
if (!$O) { 
    Send-Status -status "ERROR" -details "CRITICAL: Security Engine cannot establish connection with Microsoft Outlook. Please ensure Outlook is open and responsive."
    exit 1 
}

$N = $null
$RunspacePool = $null
$Global:Watchers = New-Object System.Collections.Generic.List[object]
$Global:ScanQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$CurrentBatch = New-Object System.Collections.Generic.List[object]

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
                    $t = $null
                    try {
                        $t = Invoke-OutlookMethod { $N.GetItemFromID($itemData.Id) }
                        if ($t) {
                            if ($Global:ReleasedFingerprints.Contains($itemData.Finger)) { $R.mv = "CLEAN"; $R.verdict = "Safe" }
                            if ($R.mv -eq "MALICIOUS") {
                                $def3 = $null
                                try {
                                    $def3 = Get-TargetFolder-Safe $t 3
                                    if ($def3) { 
                                        $m = Robust-Move $t $def3
                                        if ($m) { 
                                            [void]$ps.Add($itemData.Finger)
                                            Send-Status -status "THREAT BLOCKED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -fingerprint $itemData.Finger
                                            Release-Com $m 
                                        } 
                                    }
                                } finally { Release-Com $def3 }
                            } elseif ($R.mv -eq "SPAM") {
                                $def23 = $null
                                try {
                                    $def23 = Get-TargetFolder-Safe $t 23
                                    if ($def23) { 
                                        $m = Robust-Move $t $def23
                                        if ($m) { 
                                            [void]$ps.Add($itemData.Finger)
                                            Send-Status -status "SPAM FILTERED" -details $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $m.UnRead -fingerprint $itemData.Finger
                                            Release-Com $m 
                                        } 
                                    }
                                } finally { Release-Com $def23 }
                            } else { 
                                [void]$ps.Add($itemData.Finger)
                                Send-Status -status "Finished" -details $itemData.Su -verdict "Safe" -entryId $itemData.Id -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $t.UnRead -fingerprint $itemData.Finger 
                            }
                        }
                    } catch {} finally { Release-Com $t }
                }
                try { $job.PS.Dispose() } catch {}
            }
        }
        foreach ($r in $toRemove) { [void]$remaining.Remove($r) }
        if ($remaining.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }
}

$AnalysisScript = {
    param($itemData, $sk, $ru, $wl, $bl, $Vk, $Til)
    
    $ID = if ([string]::IsNullOrEmpty($itemData.Su)) { "NoSub" } else { $itemData.Su.Substring(0, [Math]::Min(20, $itemData.Su.Length)) }
    function Trace($m) { Send-Structured-Message @{status="INFO"; details="TRACE [$ID] -> $m"} }

    $sc = 0.0; $hits = New-Object System.Collections.Generic.List[string]; $W = $ru.weights; $T = $ru.toggles
    $Se = $itemData.Se; $IP = $itemData.IP; $Do = $itemData.Do; $bare = if ($Se -match "<(.+)>$") { $Matches[1] } else { $Se }; $combo = "$IP|$Do"
    
    Trace "INIT: Multi-Stage Security Analysis Pipeline Started."
    Log-Progress "Engine: Analyzing [$ID] - Stage 1 (Reputation)"
    
    # STAGE 1: REPUTATION API (LOCAL)
    Trace "STAGE 1: Querying Local Reputation APIs..."
    Trace "Checking Whitelist for: $bare | IP: $IP | Domain: $Do"
    if ($wl.emails -contains $bare -or $wl.ips -contains $IP -or $wl.domains -contains $Do -or $wl.combos -contains $combo) { 
        Trace "RESULT: Positive Match in Whitelist. Logic: Short-circuit to CLEAN."
        Log-Progress "Engine: [$ID] Whitelisted Sender detected."
        return @{ mv = "CLEAN"; verdict = "Safe"; score = 100; tier = "Whitelisted" } 
    }
    if ($bl.emails -contains $bare -or $bl.ips -contains $IP -or $bl.domains -contains $Do -or $bl.combos -contains $combo) { 
        Trace "RESULT: Positive Match in Blacklist. Logic: Short-circuit to SPAM."
        Log-Progress "Engine: [$ID] Blacklisted Sender detected."
        return @{ mv = "SPAM"; verdict = "Spam"; score = 0; tier = "Blacklisted" } 
    }
    
    # STAGE 2: TRANSPORT SECURITY (RFC COMPLIANCE)
    Log-Progress "Engine: Analyzing [$ID] - Stage 2 (RFC Compliance)"
    Trace "STAGE 2: Analyzing RFC Transport Compliance..."
    if ($T.dmarc) { 
        Trace "Logic: Evaluating DMARC Alignment..."
        if ($itemData.Hs -match "dmarc=(?<res>fail|bestguesspass|none)") { 
            $res = $Matches['res']; $sc += ($W.dmarc / 10.0); [void]$hits.Add("DMARC:$res"); Trace "COMPLIANCE: DMARC $res detected. Score Penalty: +$($W.dmarc/10.0)"
        }
    }
    if ($T.spf) { 
        Trace "Logic: Evaluating SPF (Sender Policy Framework)..."
        if ($itemData.Hs -match "spf=(?<res>fail|softfail|none)") { 
            $res = $Matches['res']; $sc += ($W.spf / 10.0); [void]$hits.Add("SPF:$res"); Trace "COMPLIANCE: SPF $res detected. Score Penalty: +$($W.spf/10.0)"
        }
    }
    
    # STAGE 3: MIME & HEADER FORENSICS
    Log-Progress "Engine: Analyzing [$ID] - Stage 3 (Header Forensics)"
    Trace "STAGE 3: MIME Structure & Forensic Header Analysis..."
    if ($itemData.Hs -match "X-Spam-Flag:\s*YES") { $sc += 2.0; [void]$hits.Add("HEADER:X-Spam-Flag") }
    if ($itemData.Hs -match "Content-Type:\s*application/(x-executable|x-msdownload|x-bat|x-vbs)") { $sc += 3.0; [void]$hits.Add("MIME:DangerousAttachment") }

    # STAGE 4: HEURISTIC ENGINE (CONTENT)
    Log-Progress "Engine: Analyzing [$ID] - Stage 4 (Heuristics)"
    Trace "STAGE 4: Executing Local Heuristic Pattern Matching..."
    if ($T.heuristics) { 
        $kwMatch = 0
        foreach ($kw in $sk) { 
            if ($itemData.Su -match "\b$([regex]::Escape($kw))\b" -or $itemData.by -match "\b$([regex]::Escape($kw))\b") { 
                $kwMatch++; if ($kwMatch -ge 3) { break }
            } 
        }
        if ($kwMatch -gt 0) { 
            $penalty = (($W.heuristics * $kwMatch) / 10.0)
            $sc += $penalty; [void]$hits.Add("HEURISTICS:MATCHx$kwMatch")
        }
    }

    # STAGE 5: ADVANCED INTEL (MALWARE)
    Log-Progress "Engine: Analyzing [$ID] - Stage 5 (Malware Intelligence)"
    $score = [Math]::Max(0, (100 - ($sc * 10)))
    $triggerMalwareScan = $false
    if ($Til -eq 2) { $triggerMalwareScan = $true; Trace "INTEL: Global Enforcement Mode (Tier 2). Triggering full malware audit." }
    elseif ($Til -eq 1 -and $score -lt 65) { $triggerMalwareScan = $true; Trace "INTEL: Low Confidence Score ($score%). Triggering targeted malware audit." }

    if ($triggerMalwareScan) {
        Trace "Starting Crypto-Analysis of message components..."
        $malwareHits = New-Object System.Collections.Generic.List[string]
        if (![string]::IsNullOrEmpty($itemData.by)) { 
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($itemData.by)
            $h = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
            $bodyHash = [System.BitConverter]::ToString($h).Replace("-", "").ToLower()
            Trace "INTEL: Message Body SHA256: $bodyHash (Sent to Cloud Reputation API)"
            [void]$malwareHits.Add("BODY_HASH:$bodyHash")
        }
        if ($itemData.Attachments -and $itemData.Attachments.Count -gt 0) {
            foreach ($at in $itemData.Attachments) {
                if ($at.hash -and $at.hash -ne "N/A") { 
                    Trace "INTEL: Attachment [$( $at.name )] SHA256: $( $at.hash ) (Sent to Malware Sandbox API)"
                    [void]$malwareHits.Add("FILE_HASH:$($at.name):$($at.hash)") 
                }
            }
        }
        if ($malwareHits.Count -gt 0) { [void]$hits.Add("ADVANCED_INTEL(" + ([string]::Join(", ", $malwareHits)) + ")") }
    }

    Trace "ANALYSIS COMPLETE: Consolidating Verdict..."
    $verdict = if ($score -le 20) { "Malicious" } elseif ($score -le $ru.spamThresholdPercent) { "Spam" } else { "Safe" }
    Trace "FINAL VERDICT: [$verdict] | Integrity Score: $score%"
    
    return @{ mv = if ($verdict -eq "Malicious") { "MALICIOUS" } elseif ($verdict -eq "Spam") { "SPAM" } else { "CLEAN" }; verdict = $verdict; score = $score; tier = ([string]::Join(" | ", $hits) -replace "^$", "Analysis Complete") }
}

try {
    $N = $O.GetNamespace("MAPI")
    if ($null -eq $Global:ExcludedFolderIds) { $Global:ExcludedFolderIds = New-Object System.Collections.Generic.HashSet[string] }
    Init-Exclusions $N
    if ($null -eq $Global:ReleasedFingerprints) { $Global:ReleasedFingerprints = New-Object System.Collections.Generic.HashSet[string] }

    # ON-ACCESS: Listen for new items AND perform initial unread sweep
    Log-Progress "On-Access: Initializing protection. Performing sweep of unread items..."
    $stores = $null
    try {
        $stores = $N.Stores
        foreach ($S in $stores) {
            $inbox = $null
            try { 
                $inbox = Invoke-OutlookMethod { $S.GetDefaultFolder(6) } 
            } catch {}
            if ($inbox) {
                # Initial Sweep
                $unread = $null
                try { 
                    $unread = Invoke-OutlookMethod { $inbox.Items.Restrict("[Unread] = true") }
                    if ($unread) {
                        foreach ($u in $unread) {
                            $Global:ScanQueue.Enqueue($u.EntryID)
                            Release-Com $u
                        }
                    }
                } catch {}
                finally { Release-Com $unread }

                # Live Listener: Anchor BOTH $inbox and $items to prevent GC disconnect of COM connection point
                $items = $inbox.Items
                Register-ObjectEvent -InputObject $items -EventName "ItemAdd" -Action {
                    param($item)
                    if ($item.MessageClass -eq "IPM.Note") {
                        $Global:ScanQueue.Enqueue($item.EntryID)
                    }
                } | Out-Null
                [void]$Global:Watchers.Add(@{ Folder=$inbox; Items=$items })
            } else {
                Release-Com $inbox
            }
            Release-Com $S
        }
    } catch {
        Log-Progress "On-Access: Event registration failed. Falling back to poll."
    } finally {
        Release-Com $stores
    }

    Send-Heartbeat
    $C = [Console]::In.ReadLine()
    if (!$C) { exit 0 }
    $Ex = $C | ConvertFrom-Json
    $sk = $Ex.spamKeywords
    $ru = $Ex.rubrics
    $wl = $Ex.whitelist
    $bl = $Ex.blacklist
    $Vk = $Ex.vtKey
    $Til = $Ex.threatIntelligenceLevel
    if ($Ex.releasedFingerprints) { foreach ($fp in $Ex.releasedFingerprints) { [void]$Global:ReleasedFingerprints.Add($fp) } }
    $ps = New-Object System.Collections.Generic.HashSet[string]
    if ($Ex.processedIds) { foreach ($id in $Ex.processedIds) { [void]$ps.Add($id) } }

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new("Global:StdoutLock", $Global:StdoutLock, "Stdout synchronization lock"))
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new("StdoutLock", $Global:StdoutLock, "Stdout synchronization lock"))
    $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new("Send-Structured-Message", (Get-Item function:Send-Structured-Message).Definition))
    $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new("Log-Progress", (Get-Item function:Log-Progress).Definition))

    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 16, $iss, $Host)
    $RunspacePool.Open()

    # FULL SCAN / HISTORY LOGIC
    if ($Ex.mode -eq "History") {
        Log-Progress "Forensic Discovery: Starting master mailbox crawl..."
        $stack = New-Object System.Collections.Generic.Stack[object]
        $totalCount = 0
        $historyStores = $null
        try {
            $historyStores = $N.Stores
            foreach ($S in $historyStores) { 
                @(6, 5) | ForEach-Object { 
                    $sf = $null
                    try { 
                        $sf = Invoke-OutlookMethod { $S.GetDefaultFolder($_) }
                        if ($sf) { 
                            $sfItems = $null
                            try {
                                $sfItems = $sf.Items
                                $totalCount += $sfItems.Count
                            } finally { Release-Com $sfItems }
                            $stack.Push(@{
                                FolderId = $sf.EntryID
                                StoreId = $S.StoreID
                                FolderName = $sf.Name
                                DefaultItemType = $sf.DefaultItemType
                            })
                        } 
                    } catch {}
                    finally { Release-Com $sf }
                } 
                Release-Com $S
            }
        } finally {
            Release-Com $historyStores
        }

        $currentCount = 0
        while ($stack.Count -gt 0) {
            $fDesc = $stack.Pop()
            $f = $null
            $fItems = $null
            $fSubs = $null
            try {
                $f = Invoke-OutlookMethod { $N.GetFolderFromID($fDesc.FolderId, $fDesc.StoreId) }
                if ($f) {
                    if ($fDesc.DefaultItemType -eq 0) {
                        $fItems = $f.Items
                        foreach ($t in $fItems) {
                            try {
                                $currentCount++
                                $fData = Parse-Forensics $t
                                $fp = Get-Fingerprint -item $t -ip $fData.ip
                                if ($ps.Contains($fp) -or $Global:ReleasedFingerprints.Contains($fp)) { continue }
                                $itemData = @{ Id=$t.EntryID; Su=(Get-Property-Safe $t "0x0037001E" "No Subject"); Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$fData.body; Finger=$fp; Attachments=$fData.attachments }
                                $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk).AddArgument($Til)
                                $psi.RunspacePool = $RunspacePool
                                [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
                                
                                if ($currentCount % 5 -eq 0) {
                                    Send-Status -status "SCANNING" -details "Analyzing mailbox: $currentCount / $totalCount" -count $currentCount -total $totalCount -currentFolder $fDesc.FolderName
                                }

                                if ($CurrentBatch.Count -ge 8) { Process-Batch }
                            } catch {} finally { Release-Com $t }
                        }
                    }

                    $fSubs = $f.Folders
                    if ($fSubs) {
                        foreach ($sub in $fSubs) {
                            $subItems = $null
                            try {
                                $subItems = $sub.Items
                                $totalCount += $subItems.Count
                                $stack.Push(@{
                                    FolderId = $sub.EntryID
                                    StoreId = $fDesc.StoreId
                                    FolderName = $sub.Name
                                    DefaultItemType = $sub.DefaultItemType
                                })
                            } catch {}
                            finally {
                                Release-Com $subItems
                                Release-Com $sub
                            }
                        }
                    }
                }
            } catch {}
            finally {
                Release-Com $fItems
                Release-Com $fSubs
                Release-Com $f
            }
        }
        if ($CurrentBatch.Count -gt 0) { Process-Batch }
        Send-Status -status "Finished" -details "Audit complete. Processed $currentCount items." -count $currentCount -total $totalCount
    }

    Send-Status -status "MONITORING" -details "Live protection active."

    $StdInReader = [System.IO.StreamReader]::new([Console]::OpenStandardInput())
    $ReadTask = $StdInReader.ReadLineAsync()

    while ($true) {
        Send-Heartbeat
        
        if ($ReadTask.IsCompleted) {
            $line = $ReadTask.Result
            if (![string]::IsNullOrEmpty($line)) {
                $upd = try { $line | ConvertFrom-Json } catch { $null }
                if ($upd -and $upd.type -eq "config-update") {
                    if ($null -ne $upd.whitelist) { $wl = $upd.whitelist }
                    if ($null -ne $upd.blacklist) { $bl = $upd.blacklist }
                    if ($null -ne $upd.spamKeywords) { $sk = $upd.spamKeywords }
                    if ($null -ne $upd.rubrics) { $ru = $upd.rubrics }
                    if ($null -ne $upd.vtKey) { $Vk = $upd.vtKey }
                    if ($null -ne $upd.threatIntelligenceLevel) { $Til = $upd.threatIntelligenceLevel }
                    Log-Progress "Engine: Applied Live Policy Update without restarting."
                }
            }
            $ReadTask = $StdInReader.ReadLineAsync()
        }

        $id = $null
        while ($Global:ScanQueue.TryDequeue([ref]$id)) {
            $lt = $null
            try {
                $lt = Invoke-OutlookMethod { $N.GetItemFromID($id) }
                if ($lt) {
                    $fData = Parse-Forensics $lt
                    $fp = Get-Fingerprint -item $lt -ip $fData.ip
                    if ($ps.Contains($fp) -or $Global:ReleasedFingerprints.Contains($fp)) { continue }
                    $itemData = @{ Id=$id; Su=(Get-Property-Safe $lt "0x0037001E" "No Subject"); Se=$fData.from; IP=$fData.ip; Hs=$fData.headers; by=$fData.body; Finger=$fp; Attachments=$fData.attachments }
                    $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk).AddArgument($Til)
                    $psi.RunspacePool = $RunspacePool
                    [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
                }
            } catch {} finally { Release-Com $lt }
        }
        if ($CurrentBatch.Count -gt 0) { Process-Batch }
        Start-Sleep -Milliseconds 500
    }
} finally {
    if ($null -ne $Global:Watchers) {
        foreach ($w in $Global:Watchers) {
            try { Release-Com $w.Items } catch {}
            try { Release-Com $w.Folder } catch {}
        }
        $Global:Watchers.Clear()
    }
    Release-Com $N
    Release-Com $O
    if ($null -ne $RunspacePool) {
        try { $RunspacePool.Close(); $RunspacePool.Dispose() } catch {}
    }
    if ($null -ne $Global:OutlookComMutex) {
        try { $Global:OutlookComMutex.Dispose() } catch {}
    }
}

