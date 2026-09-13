param([string]$Mode = "", [int]$ParentPid = 0)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- SINGLE-INSTANCE CONCURRENCY GUARD ---
$Global:InstanceMutex = $null
if ($Mode -ne "Lib") {
    $instanceMutexName = if ($Mode -eq "Worker") { "Local\DeskGuard_Worker_Instance" } else { "Local\DeskGuard_Scanner_Instance" }
    $isCreatedNew = $false
    try {
        $Global:InstanceMutex = New-Object System.Threading.Mutex($true, $instanceMutexName, [ref]$isCreatedNew)
        if (!$isCreatedNew) {
            [Console]::Error.WriteLine("CRITICAL: Another DeskGuard instance for role '$Mode' is already active. Terminating duplicate to prevent COM contention.")
            [System.Environment]::Exit(0)
        }
    } catch {
        # If error initializing mutex, continue cautiously
    }
}

# --- GLOBAL STATE PERSISTENCE & SYNCHRONIZATION ---
if ($null -eq $Global:StdoutLock) { $Global:StdoutLock = [System.Object]::new() }
if ($null -eq $Global:DupStack) { $Global:DupStack = $null }
if ($null -eq $Global:DupHashes) { $Global:DupHashes = @{} }
if ($null -eq $Global:DupResults) { $Global:DupResults = New-Object System.Collections.Generic.List[object] }
if ($null -eq $Global:DupScannedCount) { $Global:DupScannedCount = 0 }
if ($null -eq $Global:EventSubscriberIds) { $Global:EventSubscriberIds = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList)) }

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
        if ($len -le 0 -or $len -gt 10485760) { return }
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

function Get-MD5 {
    param($string)
    if ([string]::IsNullOrEmpty($string)) { return "N/A" }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($string)
    $hash = [System.Security.Cryptography.MD5]::Create().ComputeHash($bytes)
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

function Get-ItemSubject {
    param($mailItem)
    if (!$mailItem) { return "(No Subject)" }
    $su = ""
    try { $su = $mailItem.Subject } catch {}
    if ([string]::IsNullOrWhiteSpace($su)) {
        $su = Get-Property-Safe $mailItem "0x0037001F" ""
    }
    if ([string]::IsNullOrWhiteSpace($su)) {
        $su = Get-Property-Safe $mailItem "0x0037001E" ""
    }
    if ([string]::IsNullOrWhiteSpace($su)) {
        $su = Get-Property-Safe $mailItem "0x0070001F" ""
    }
    if ([string]::IsNullOrWhiteSpace($su)) {
        $su = Get-Property-Safe $mailItem "0x0070001E" ""
    }
    if ([string]::IsNullOrWhiteSpace($su)) {
        $hdrs = Get-Property-Safe $mailItem "0x007D001E" ""
        if ([string]::IsNullOrWhiteSpace($hdrs)) { $hdrs = Get-Property-Safe $mailItem "0x007D001F" "" }
        if (![string]::IsNullOrWhiteSpace($hdrs) -and $hdrs -match "(?im)^Subject:\s*(.*?)$") {
            $su = $Matches[1].Trim()
        }
    }
    if ([string]::IsNullOrWhiteSpace($su)) {
        return "(No Subject)"
    }
    return $su
}

function Get-ItemSender {
    param($mailItem)
    if (!$mailItem) { return "Unknown" }
    $sender = ""
    
    # 0. Primary SMTP Address property tag: PR_SENDER_SMTP_ADDRESS (0x5D01001F / 0x5D01001E)
    $smtpSender = Get-Property-Safe $mailItem "0x5D01001F" ""
    if ([string]::IsNullOrWhiteSpace($smtpSender)) {
        $smtpSender = Get-Property-Safe $mailItem "0x5D01001E" ""
    }
    if (![string]::IsNullOrWhiteSpace($smtpSender) -and $smtpSender -match "@") {
        return $smtpSender.Trim()
    }

    # 1. Transport RFC Headers (PR_TRANSPORT_MESSAGE_HEADERS: 0x007D001E / 0x007D001F)
    $headers = Get-Property-Safe $mailItem "0x007D001E" ""
    if ([string]::IsNullOrWhiteSpace($headers)) {
        $headers = Get-Property-Safe $mailItem "0x007D001F" ""
    }
    if (![string]::IsNullOrWhiteSpace($headers)) {
        if ($headers -match "(?im)^From:\s*(?:.*?<)?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})>?") {
            $sender = $Matches[1].ToLower().Trim()
        } elseif ($headers -match "(?im)^Return-Path:\s*<([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})>") {
            $sender = $Matches[1].ToLower().Trim()
        }
    }
    if (![string]::IsNullOrWhiteSpace($sender) -and $sender -ne "Unknown") {
        return $sender
    }

    # 2. PR_SENT_REPRESENTING_SMTP_ADDRESS (0x5D07001F / 0x5D07001E)
    $repSmtp = Get-Property-Safe $mailItem "0x5D07001F" ""
    if ([string]::IsNullOrWhiteSpace($repSmtp)) {
        $repSmtp = Get-Property-Safe $mailItem "0x5D07001E" ""
    }
    if (![string]::IsNullOrWhiteSpace($repSmtp) -and $repSmtp -match "@") {
        return $repSmtp.Trim()
    }

    # 3. Direct SenderEmailAddress if clean SMTP address
    try { 
        $sea = $mailItem.SenderEmailAddress
        if (![string]::IsNullOrWhiteSpace($sea) -and $sea -notmatch "/o=" -and $sea -match "@") {
            return $sea.Trim()
        }
    } catch {}

    # 4. Try $mailItem.Sender with Resolve-Email
    $Snd = $null
    try { 
        $Snd = $mailItem.Sender
        if ($Snd) { 
            $resolved = Resolve-Email -Recipient $Snd
            if (![string]::IsNullOrWhiteSpace($resolved) -and $resolved -ne "Unknown" -and $resolved -notmatch "/o=") {
                return $resolved
            }
        } 
    } catch {} 
    finally { 
        Release-Com $Snd 
    }

    # 5. Fall back to PR_SENDER_EMAIL_ADDRESS ('0x0065001E' / '0x0065001F')
    $pSender = Get-Property-Safe $mailItem "0x0065001E" ""
    if ([string]::IsNullOrWhiteSpace($pSender)) {
        $pSender = Get-Property-Safe $mailItem "0x0065001F" ""
    }
    if (![string]::IsNullOrWhiteSpace($pSender) -and $pSender -ne "N/A" -and $pSender -ne "Unknown" -and $pSender -notmatch "/o=") {
        return $pSender.Trim()
    }

    # 6. Fall back to SenderName and PR_SENDER_NAME ('0x0042001E' / '0x0042001F')
    try {
        $sName = $mailItem.SenderName
        if (![string]::IsNullOrWhiteSpace($sName) -and $sName -ne "Unknown" -and $sName -notmatch "/o=") {
            return $sName.Trim()
        }
    } catch {}

    $pName = Get-Property-Safe $mailItem "0x0042001E" ""
    if ([string]::IsNullOrWhiteSpace($pName)) {
        $pName = Get-Property-Safe $mailItem "0x0042001F" ""
    }
    if (![string]::IsNullOrWhiteSpace($pName) -and $pName -ne "N/A" -and $pName -ne "Unknown" -and $pName -notmatch "/o=") {
        return $pName.Trim()
    }

    # 7. Clean up residual /o= X.500 Distinguished Names if all else failed
    if (![string]::IsNullOrWhiteSpace($sender) -and $sender -match "/o=") {
        if ($sender -match "cn=([^/]+)$") { return $Matches[1] }
        else { return ($sender -split "=")[-1] }
    }

    return "Unknown"
}

function Get-Fingerprint {
    param($item, $ip)
    if (!$item) { return [guid]::NewGuid().ToString() }
    
    $se = Get-ItemSender $item
    $su = Get-ItemSubject $item
    $rt = "00000000000000"; try { if ($item.ReceivedTime) { $rt = $item.ReceivedTime.ToString("yyyyMMddHHmmss") } } catch {}
    
    return Get-SHA256 "$se|$su|$rt"
}

function Get-ItemDna {
    param($mailItem)
    if (!$mailItem) { return "" }
    try {
        $subject = Get-ItemSubject $mailItem
        $cleanSubject = ($subject -replace "^(Re:|Fwd:|FW:|RE:)\s*", "").Trim()
        $sender = Get-ItemSender $mailItem
        $received = "00000000000000"
        try { if ($mailItem.ReceivedTime) { $received = $mailItem.ReceivedTime.ToString("yyyyMMddHHmmss") } } catch {}
        $bodySample = ""
        try { $bodySample = $mailItem.Body } catch {}
        if ($bodySample.Length -gt 300) { $bodySample = $bodySample.Substring(0, 300) }
        return Get-SHA256 "$sender|$cleanSubject|$received|$bodySample"
    } catch {
        return ""
    }
}

function Send-Status {
    param([string]$status, [string]$details, [string]$verdict = "Pending", [string]$action = "None", [string]$entryId = "", [string]$originalEntryId = "", [string]$tier = "", [string]$phase = "", [string]$sender = "", [string]$ip = "", [string]$domain = "", [string]$originalFolder = "", [string]$fullHeaders = "", [float]$score = 0, [string]$body = "", [bool]$unread = $false, [string]$scanType = "", [string]$to = "", [string]$cc = "", [string]$fingerprint = "", [string]$timestamp = "", [int]$count = 0, [int]$total = 0, [string]$currentFolder = "", [string]$subject = "", [string]$date = "", [string]$time = "", [string]$storeId = "")
    if ($null -eq $fullHeaders) { $fullHeaders = "" }
    if ($null -eq $body) { $body = "" }
    if ($fullHeaders.Length -gt 1048576) { $fullHeaders = $fullHeaders.Substring(0, 1048576) }
    if ($body.Length -gt 4194304) { $body = $body.Substring(0, 4194304) }
    $h = ""; if (![string]::IsNullOrEmpty($fullHeaders)) { try { $h = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($fullHeaders)) } catch {} }
    $b = ""; if (![string]::IsNullOrEmpty($body)) { try { $b = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($body)) } catch {} }
    $ts = if ([string]::IsNullOrEmpty($timestamp)) { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") } else { $timestamp }
    $d = if (![string]::IsNullOrEmpty($date)) { $date } else { $ts.Split(" ")[0] }
    $t = if (![string]::IsNullOrEmpty($time)) { $time } else { if ($ts.Split(" ").Length -gt 1) { $ts.Split(" ")[1] } else { "" } }
    $resolvedSubject = if (![string]::IsNullOrWhiteSpace($subject)) { $subject } else { $details }
    Send-Structured-Message @{
        timestamp=$ts; date=$d; time=$t; status=$status; details=$details; subject=$resolvedSubject; verdict=$verdict; action=$action;
        entryId=$entryId; originalEntryId=$originalEntryId; storeId=$storeId; tier=$tier; phase=$phase; sender=$sender;
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
    if (!$item) {
        return @{ ip="N/A"; headers=""; from="Unknown"; body=""; attachments=@() }
    }
    
    # 1. Transport Headers (PR_TRANSPORT_MESSAGE_HEADERS: 0x007D001E / 0x007D001F)
    $headers = Get-Property-Safe $item "0x007D001E" ""
    if ([string]::IsNullOrWhiteSpace($headers)) {
        $headers = Get-Property-Safe $item "0x007D001F" ""
    }
    
    # 2. Extract Sender safely (RFC headers first, bypassing OOM guard)
    $from = Get-ItemSender $item

    # 2b. Extract Recipient and CC
    $toRecips = ""
    $ccRecips = ""
    try {
        $rList = New-Object System.Collections.Generic.List[string]
        $cList = New-Object System.Collections.Generic.List[string]
        $recips = $item.Recipients
        if ($recips -and $recips.Count -gt 0) {
            foreach ($r in $recips) {
                $rType = $r.Type  # 1 = olTo, 2 = olCC, 3 = olBCC
                $smtp = Get-Property-Safe $r "0x39FE001F" ""
                if ([string]::IsNullOrWhiteSpace($smtp) -or $smtp -eq "N/A") {
                    $smtp = Get-Property-Safe $r "0x39FE001E" ""
                }
                $addr = if (![string]::IsNullOrWhiteSpace($smtp) -and $smtp -ne "N/A" -and $smtp -match "@") {
                    $smtp.Trim()
                } else {
                    $rAddr = ""; try { $rAddr = $r.Address } catch {}
                    if (![string]::IsNullOrWhiteSpace($rAddr) -and $rAddr -notmatch "/o=" -and $rAddr -match "@") {
                        $rAddr.Trim()
                    } else {
                        $r.Name
                    }
                }
                if ($rType -eq 2) {
                    [void]$cList.Add($addr)
                } else {
                    [void]$rList.Add($addr)
                }
                Release-Com $r
            }
            Release-Com $recips
        }
        if ($rList.Count -gt 0) { $toRecips = ($rList -join "; ") }
        if ($cList.Count -gt 0) { $ccRecips = ($cList -join "; ") }
    } catch {}
    if ([string]::IsNullOrWhiteSpace($toRecips)) {
        try { if ($item.To) { $toRecips = $item.To } } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($ccRecips)) {
        try { if ($item.CC) { $ccRecips = $item.CC } } catch {}
    }

    # 2c. Extract Date and Time
    $recTime = $null
    try { $recTime = $item.ReceivedTime } catch {}
    if (!$recTime -or $recTime -eq [DateTime]::MinValue) {
        try { $recTime = $item.SentOn } catch {}
    }
    if (!$recTime -or $recTime -eq [DateTime]::MinValue) {
        $recTime = [DateTime]::Now
    }
    $recDate = $recTime.ToString("yyyy-MM-dd")
    $recClock = $recTime.ToString("HH:mm:ss")
    
    # 3. Extract Originating Sender IP
    $senderIp = "Internal Network"
    if (![string]::IsNullOrWhiteSpace($headers)) {
        # Strategy A: client-ip in Authentication-Results / Received-SPF
        if ($headers -match "(?i)client-ip=(?<ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
            $cand = $Matches['ip']
            if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") { $senderIp = $cand }
        }
        # Strategy B: "sender IP is <ip>" in Authentication-Results
        if ($senderIp -eq "Internal Network" -and $headers -match "(?i)sender\s+IP\s+is\s+(?<ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
            $cand = $Matches['ip']
            if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") { $senderIp = $cand }
        }
        # Strategy C: X-Originating-IP
        if ($senderIp -eq "Internal Network" -and $headers -match "(?im)^X-Originating-IP:\s*\[?(?<ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\]?") {
            $cand = $Matches['ip']
            if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") { $senderIp = $cand }
        }
        # Strategy D: Originating Received hop from bottom up (originating hop first)
        if ($senderIp -eq "Internal Network") {
            $rxHops = [regex]::Matches($headers, "(?s)Received:\s*from\s+.*?;\s*[\w, ]+\d{4}")
            for ($k = $rxHops.Count - 1; $k -ge 0; $k--) {
                $hop = $rxHops[$k].Value
                $ipMatches = [regex]::Matches($hop, "(?:\[|\()(?<ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})(?:\]|\))")
                foreach ($m in $ipMatches) {
                    $cand = $m.Groups['ip'].Value
                    if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") {
                        $senderIp = $cand
                        break
                    }
                }
                if ($senderIp -ne "Internal Network") { break }
            }
        }
        # Strategy E: Fallback to any public IP found in brackets or parentheses
        if ($senderIp -eq "Internal Network") {
            $allIps = [regex]::Matches($headers, "(?:\[|\()(?<ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})(?:\]|\))")
            for ($k = $allIps.Count - 1; $k -ge 0; $k--) {
                $cand = $allIps[$k].Groups['ip'].Value
                if ($cand -notmatch "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.|127\.|169\.254\.)") {
                    $senderIp = $cand
                    break
                }
            }
        }
    }
    
    $subj = Get-ItemSubject $item
    
    # If no transport headers (internal Exchange / MAPI store), construct synthetic envelope headers so forensics is never blank
    if ([string]::IsNullOrWhiteSpace($headers)) {
        $sentDate = ""; try { if ($item.SentOn) { $sentDate = $item.SentOn.ToString("r") } } catch {}
        $msgClass = "IPM.Note"; try { if ($item.MessageClass) { $msgClass = $item.MessageClass } } catch {}
        $toRecips = ""; try { if ($item.To) { $toRecips = $item.To } } catch {}
        $ccRecips = ""; try { if ($item.CC) { $ccRecips = $item.CC } } catch {}
        $hdrLines = @(
            "X-Delivery-Context: Internal / MAPI Store (No SMTP Transport Headers)",
            "Message-Class: $msgClass",
            "From: $from",
            "To: $toRecips",
            "CC: $ccRecips",
            "Subject: $subj",
            "Date: $sentDate"
        )
        $headers = ($hdrLines -join "`r`n")
    }
    
    # 4. Extract Body safely (native OOM Body first, fallback to HTMLBody text)
    $body = ""
    try { $body = $item.Body } catch {}
    if ([string]::IsNullOrWhiteSpace($body)) {
        try {
            $html = $item.HTMLBody
            if (![string]::IsNullOrWhiteSpace($html)) {
                $body = [System.Text.RegularExpressions.Regex]::Replace($html, "<[^>]+>", " ").Trim()
            }
        } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($body)) {
        $body = "(No message body content)"
    }
    
    # 5. Extract Attachments
    $atts = New-Object System.Collections.Generic.List[object]
    $AttsObj = $null
    try {
        $AttsObj = $item.Attachments
        if ($AttsObj -and $AttsObj.Count -gt 0) {
            foreach ($at in $AttsObj) {
                $hash = "N/A"; $md5 = "N/A"
                try {
                    $hash = Get-SHA256 "$($at.FileName)|$($at.Size)"
                    $md5 = Get-MD5 "$($at.FileName)|$($at.Size)"
                } catch {}
                [void]$atts.Add(@{ name=$at.FileName; hash=$hash; md5=$md5; size=$at.Size })
                Release-Com $at
            }
        }
    } catch {}
    finally {
        Release-Com $AttsObj
    }
    
    return @{ 
        ip = $senderIp; 
        headers = $headers; 
        from = $from; 
        to = $toRecips; 
        cc = $ccRecips; 
        date = $recDate; 
        time = $recClock; 
        timestamp = "$recDate $recClock"; 
        body = $body; 
        attachments = $atts;
        subject = $subj;
        Su = $subj;
        Se = $from;
        Hs = $headers;
        by = $body
    }
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
                                        
                                        $subject = Get-ItemSubject $t
                                        
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

                                        $dna = Get-ItemDna $t
                                        if ([string]::IsNullOrEmpty($dna)) {
                                            Log-Progress "Worker: Diagnostic warning - Unable to compute DNA for item '$subject' ($($t.EntryID)). Skipping."
                                            continue
                                        }
                                        $sender = Get-ItemSender $t
                                        $receivedTs = ""
                                        try { if ($t.ReceivedTime) { $receivedTs = $t.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") } } catch {}
                                        $itemObj = @{ entryId=$t.EntryID; subject=$subject; sender=$sender; timestamp=$receivedTs; size=$itemSize; folder=$folderName; store=$storeName; dna=$dna }
                                        
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
                                    if ($folderScanned % 15 -eq 0) { Start-Sleep -Milliseconds 25 }
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
                    $bodyMd5 = if (![string]::IsNullOrEmpty($fData.body)) { Get-MD5 $fData.body } else { "N/A" }
                    $subject = Get-ItemSubject $item

                    $hashesToAudit = New-Object System.Collections.Generic.List[object]
                    if ($bodyHash -ne "N/A") {
                        [void]$hashesToAudit.Add(@{ type="body"; name="Message Body"; hash=$bodyHash; md5=$bodyMd5 })
                    }
                    foreach ($at in $fData.attachments) {
                        if ($at.hash -and $at.hash -ne "N/A") {
                            $atMd5 = if ($at.md5 -and $at.md5 -ne "N/A") { $at.md5 } else { "" }
                            [void]$hashesToAudit.Add(@{ type="attachment"; name=$at.name; hash=$at.hash; md5=$atMd5; size=$at.size })
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
                            $vtStatus = "No VirusTotal API Key Configured"
                        }

                        # Threat Intelligence Redundancy: Team Cymru Malware Hash Database (DNS TXT Query)
                        if ($hObj.md5 -and $hObj.md5 -ne "N/A" -and $hObj.md5.Length -eq 32) {
                            try {
                                $cymruDns = Resolve-DnsName -Name "$($hObj.md5).malware.hash.cymru.com" -Type TXT -ErrorAction SilentlyContinue
                                if ($cymruDns -and ($cymruDns.Strings -match "^\d+\s+\d+" -or $cymruDns.Strings -match "127\.0\.0\.2")) {
                                    $totalThreats++
                                    $maliciousVotes = [Math]::Max(1, $maliciousVotes)
                                    if ($vtStatus -match "No VirusTotal|Clean|Not seen|deferred") {
                                        $vtStatus = "MALICIOUS (Team Cymru Malware Hash DB Hit)"
                                    } else {
                                        $vtStatus += " | Cymru MHR: Match"
                                    }
                                }
                            } catch {}
                        }

                        [void]$auditResults.Add(@{
                            name = $hObj.name
                            type = $hObj.type
                            hash = $hObj.hash
                            md5 = $hObj.md5
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
            elseif ($Action -eq "OpenInOutlook") {
                $id = if ($Ex.entryId) { $Ex.entryId } elseif ($Ex.data -and $Ex.data.entryId) { $Ex.data.entryId } else { $null }
                $storeId = if ($Ex.storeId) { $Ex.storeId } elseif ($Ex.data -and $Ex.data.storeId) { $Ex.data.storeId } else { $null }
                $item = $null
                $success = $false
                $errMsg = ""
                try {
                    if ([string]::IsNullOrEmpty($id)) {
                        throw "Invalid parameter: entryId is required"
                    }
                    if (![string]::IsNullOrEmpty($storeId)) {
                        $item = Invoke-OutlookMethod { $N.GetItemFromID($id, $storeId) }
                    } else {
                        $item = Invoke-OutlookMethod { $N.GetItemFromID($id) }
                    }
                    if (!$item) {
                        throw "Email item not found in Microsoft Outlook"
                    }
                    Invoke-OutlookMethod { $item.Display($false) }
                    $success = $true
                }
                catch {
                    $errMsg = $_.Exception.Message
                }
                finally {
                    Release-Com $item
                }

                if ($Ex.rid) {
                    Send-Structured-Message @{ type="cmd-response"; rid=$Ex.rid; success=$success; error=$errMsg; data=@{ success=$success; error=$errMsg } }
                }
            }
        }
    } finally {
        Release-Com $N
        Release-Com $O
        if ($null -ne $Global:OutlookComMutex) {
            try { $Global:OutlookComMutex.Dispose() } catch {}
        }
        if ($null -ne $Global:InstanceMutex) {
            try {
                $Global:InstanceMutex.ReleaseMutex()
                $Global:InstanceMutex.Dispose()
            } catch {}
        }
    }
    exit 0
}

# --- SCANNER MODE ---
if ($Mode -ne "Lib" -and $Mode -ne "Worker") {
$O = Get-Outlook
if (!$O) { 
    Send-Status -status "ERROR" -details "CRITICAL: Security Engine cannot establish connection with Microsoft Outlook. Please ensure Outlook is open and responsive."
    exit 1 
}

$N = $null
$RunspacePool = $null
$Global:Watchers = New-Object System.Collections.Generic.List[object]
$Global:EventSubscriberIds = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
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
                        if (![string]::IsNullOrWhiteSpace($itemData.StoreId)) {
                            $t = Invoke-OutlookMethod { $N.GetItemFromID($itemData.Id, $itemData.StoreId) }
                        }
                        if (!$t) {
                            $t = Invoke-OutlookMethod { $N.GetItemFromID($itemData.Id) }
                        }
                        if ($Global:ReleasedFingerprints.Contains($itemData.Finger)) { $R.mv = "CLEAN"; $R.verdict = "Safe" }
                        
                        if ($R.mv -eq "MALICIOUS" -and $t) {
                            $def3 = $null
                            try {
                                $def3 = Get-TargetFolder-Safe $t 3
                                if ($def3) { 
                                    $m = Robust-Move $t $def3
                                    if ($m) { 
                                        [void]$ps.Add($itemData.Finger)
                                        $unreadVal = try { $m.UnRead } catch { $false }
                                        Send-Status -status "THREAT BLOCKED" -details $itemData.Su -subject $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $unreadVal -fingerprint $itemData.Finger -fullHeaders $itemData.Hs -body $itemData.by -to $itemData.To -cc $itemData.Cc -date $itemData.Date -time $itemData.Time -timestamp $itemData.Timestamp -storeId $itemData.StoreId
                                        Release-Com $m 
                                    } 
                                }
                            } finally { Release-Com $def3 }
                        } elseif ($R.mv -eq "SPAM" -and $t) {
                            $def23 = $null
                            try {
                                $def23 = Get-TargetFolder-Safe $t 23
                                if ($def23) { 
                                    $m = Robust-Move $t $def23
                                    if ($m) { 
                                        [void]$ps.Add($itemData.Finger)
                                        $unreadVal = try { $m.UnRead } catch { $false }
                                        Send-Status -status "SPAM FILTERED" -details $itemData.Su -subject $itemData.Su -verdict $R.verdict -action $R.action -entryId $m.EntryID -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $unreadVal -fingerprint $itemData.Finger -fullHeaders $itemData.Hs -body $itemData.by -to $itemData.To -cc $itemData.Cc -date $itemData.Date -time $itemData.Time -timestamp $itemData.Timestamp -storeId $itemData.StoreId
                                        Release-Com $m 
                                    } 
                                }
                            } finally { Release-Com $def23 }
                        } else { 
                            [void]$ps.Add($itemData.Finger)
                            $unreadVal = if ($t) { try { $t.UnRead } catch { $false } } else { $false }
                            Send-Status -status "Finished" -details $itemData.Su -subject $itemData.Su -verdict "Safe" -entryId $itemData.Id -originalEntryId $itemData.Id -sender $itemData.Se -ip $itemData.IP -score $R.score -tier $R.tier -unread $unreadVal -fingerprint $itemData.Finger -fullHeaders $itemData.Hs -body $itemData.by -to $itemData.To -cc $itemData.Cc -date $itemData.Date -time $itemData.Time -timestamp $itemData.Timestamp -storeId $itemData.StoreId
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
    $Se = $itemData.Se; $IP = $itemData.IP
    $Do = if ($itemData.Do) { $itemData.Do } elseif ($itemData.Se -match "@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})") { $Matches[1].ToLower() } else { "unknown" }
    $bare = if ($Se -match "<(.+)>$") { $Matches[1] } else { $Se }; $combo = "$IP|$Do"
    
    Trace "INIT: Multi-Stage Security Analysis Pipeline Started."
    Log-Progress "Engine: Analyzing [$ID] - Stage 1 (Reputation)"
    
    # STAGE 1: REPUTATION API (LOCAL WHITELIST / BLACKLIST)
    Trace "STAGE 1: Querying Local Reputation Lists..."
    Trace "Checking Whitelist for: $bare | IP: $IP | Domain: $Do"
    if ($wl.emails -contains $bare -or $wl.ips -contains $IP -or $wl.domains -contains $Do -or $wl.combos -contains $combo) { 
        Trace "RESULT: Positive Match in Whitelist. Logic: Short-circuit to CLEAN."
        Log-Progress "Engine: [$ID] Whitelisted Sender detected."
        return @{ mv = "CLEAN"; verdict = "Safe"; score = 100; tier = "Whitelisted"; action = "None" } 
    }
    if ($bl.emails -contains $bare -or $bl.ips -contains $IP -or $bl.domains -contains $Do -or $bl.combos -contains $combo) { 
        Trace "RESULT: Positive Match in Blacklist. Logic: Short-circuit to SPAM."
        Log-Progress "Engine: [$ID] Blacklisted Sender detected."
        return @{ mv = "SPAM"; verdict = "Spam"; score = 0; tier = "Blacklisted"; action = "Quarantined" } 
    }
    
    # STAGE 2: TRANSPORT SECURITY & REPUTATION (8 SCORING ENGINES)
    Log-Progress "Engine: Analyzing [$ID] - Stage 2 (RFC Compliance & Blacklists)"
    Trace "STAGE 2: Analyzing RFC Transport Compliance & Blacklists..."

    # Engine 1: DMARC Authentication
    if ($T.dmarc) { 
        Trace "Logic: Evaluating DMARC Alignment..."
        if ($itemData.Hs -match "dmarc=(?<res>fail|bestguesspass|none)" -or $itemData.Hs -match "Authentication-Results:.*?dmarc=(?<res>fail|none)") { 
            $res = $Matches['res']; $sc += ($W.dmarc / 10.0); [void]$hits.Add("DMARC:$res"); Trace "COMPLIANCE: DMARC $res detected. Score Penalty: +$($W.dmarc/10.0)"
        }
    }

    # Engine 2: SPF Authorization
    if ($T.spf) { 
        Trace "Logic: Evaluating SPF (Sender Policy Framework)..."
        if ($itemData.Hs -match "spf=(?<res>fail|softfail|none)" -or $itemData.Hs -match "Authentication-Results:.*?spf=(?<res>fail|softfail|none)") { 
            $res = $Matches['res']; $sc += ($W.spf / 10.0); [void]$hits.Add("SPF:$res"); Trace "COMPLIANCE: SPF $res detected. Score Penalty: +$($W.spf/10.0)"
        }
    }

    # Engine 3: DKIM Signatures
    if ($T.dkim) {
        Trace "Logic: Evaluating DKIM Cryptographic Seal..."
        if ($itemData.Hs -match "dkim=(?<res>fail|none)" -or $itemData.Hs -match "Authentication-Results:.*?dkim=(?<res>fail|none)") {
            $res = $Matches['res']; $sc += ($W.dkim / 10.0); [void]$hits.Add("DKIM:$res"); Trace "COMPLIANCE: DKIM $res detected. Score Penalty: +$($W.dkim/10.0)"
        }
    }

    # Engine 4: Sender Envelope Alignment
    if ($T.alignment -and $itemData.Hs -match "Return-Path:.*?<(?<v>.*?)>") {
        Trace "Logic: Evaluating Sender Alignment (Header From vs Return-Path)..."
        $rp = $Matches['v']
        if ($Se -and $rp -and $Se.ToLower() -ne $rp.ToLower()) {
            $sc += ($W.alignment / 10.0); [void]$hits.Add("SENDER_MISALIGNMENT"); Trace "COMPLIANCE: Sender Misalignment detected ($Se vs $rp). Score Penalty: +$($W.alignment/10.0)"
        }
    }

    # Engine 5: Reverse DNS (PTR) Verification
    if ($T.rdns -and $IP -ne "N/A" -and $IP -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1)") {
        Trace "Logic: Evaluating Reverse DNS (PTR) for $IP..."
        try {
            $ptr = [System.Net.Dns]::GetHostEntry($IP).HostName
            if (!$ptr -or ($Do -and $ptr -notmatch [regex]::Escape($Do))) {
                $sc += ($W.rdns / 10.0); [void]$hits.Add("RDNS_MISMATCH"); Trace "COMPLIANCE: Reverse DNS Mismatch ($ptr vs $Do). Score Penalty: +$($W.rdns/10.0)"
            }
        } catch {
            $sc += ($W.rdns / 10.0); [void]$hits.Add("RDNS_MISSING"); Trace "COMPLIANCE: Reverse DNS PTR query failed. Score Penalty: +$($W.rdns/10.0)"
        }
    }

    # Engine 6: 3-Provider RBL Blacklists (Spamhaus, Spamcop, Barracuda + DBL)
    if ($T.rbl -and $IP -ne "N/A" -and $IP -notmatch "^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|fe80|::1)") {
        Trace "Logic: Evaluating 3-Provider RBL Blacklists..."
        $isRBL = $false
        foreach ($rbl in @("zen.spamhaus.org", "bl.spamcop.net", "b.barracudacentral.org")) {
            try {
                $rev = ($IP -split "\.")[3..0] -join "."
                $d = Resolve-DnsName -Name "$rev.$rbl" -Type A -ErrorAction SilentlyContinue
                if ($d -and $d.IPAddress -match "^127\.0\.0\.") { 
                    $isRBL = $true
                    Trace "RBL: Listed in $rbl ($($d.IPAddress))"
                    break 
                }
            } catch {}
        }
        if (!$isRBL -and $Do -and $Do -ne "unknown") {
            try {
                $d = Resolve-DnsName -Name "$Do.dbl.spamhaus.org" -Type A -ErrorAction SilentlyContinue
                if ($d -and $d.IPAddress -match "^127\.0\.1\.") { 
                    $isRBL = $true
                    Trace "RBL: Domain listed in dbl.spamhaus.org ($($d.IPAddress))"
                }
            } catch {}
        }
        if ($isRBL) {
            $sc += ($W.rbl / 10.0); [void]$hits.Add("GLOBAL_RBL_HIT"); Trace "REPUTATION: RBL hit detected. Score Penalty: +$($W.rbl/10.0)"
        }
    }
    
    # STAGE 3: MIME & HEADER FORENSICS
    Log-Progress "Engine: Analyzing [$ID] - Stage 3 (Header Forensics)"
    Trace "STAGE 3: MIME Structure & Forensic Header Analysis..."
    if ($itemData.Hs -match "X-Spam-Flag:\s*YES") { $sc += 2.0; [void]$hits.Add("HEADER:X-Spam-Flag") }
    if ($itemData.Hs -match "Content-Type:\s*application/(x-executable|x-msdownload|x-bat|x-vbs|x-msdos-program)") { $sc += 3.0; [void]$hits.Add("MIME:DangerousAttachment") }

    # STAGE 4: HEURISTIC ENGINE (SCAM KEYWORDS & ANTI-PHISHING BODY SHIELD)
    Log-Progress "Engine: Analyzing [$ID] - Stage 4 (Heuristics & Anti-Phishing)"
    Trace "STAGE 4: Executing Heuristics & Anti-Phishing Scan..."
    
    # Engine 7: Smart Scam Heuristics
    if ($T.heuristics) { 
        $kwMatch = 0
        if ($sk -and $sk.Count -gt 0) {
            foreach ($kw in $sk) { 
                if ($itemData.Su -match "\b$([regex]::Escape($kw))\b" -or $itemData.by -match "\b$([regex]::Escape($kw))\b") { 
                    $kwMatch++; if ($kwMatch -ge 3) { break }
                } 
            }
        }
        if ($kwMatch -gt 0) { 
            $penalty = (($W.heuristics * $kwMatch) / 10.0)
            $sc += $penalty; [void]$hits.Add("HEURISTICS:MATCHx$kwMatch")
        }
    }

    # Engine 8: Anti-Phishing Body Entropy Shield
    if ($T.body -and (![string]::IsNullOrEmpty($itemData.by))) {
        if ($itemData.by -match "<script" -or $itemData.by -match "display:\s*none" -or $itemData.by -match "visibility:\s*hidden" -or $itemData.by -match "font-size:\s*0" -or $itemData.by -match "color:\s*(transparent|#fff|#ffffff|white)\s*;\s*background(-color)?:\s*(#fff|#ffffff|white)") {
            $sc += ($W.body / 10.0)
            [void]$hits.Add("HIDDEN_BODY_ENTROPY")
            Trace "HEURISTICS: Hidden text/script entropy detected. Score Penalty: +$($W.body/10.0)"
        }
    }

    # STAGE 5: ADVANCED INTEL (MULTI-TIER MALWARE AUDIT: CYMRU MHR + VIRUSTOTAL)
    Log-Progress "Engine: Analyzing [$ID] - Stage 5 (Malware Intelligence)"
    $score = [Math]::Max(0, (100 - [int]($sc * 10)))
    $triggerMalwareScan = $false
    if ($Til -eq 2) { $triggerMalwareScan = $true; Trace "INTEL: Global Enforcement Mode (Tier 2). Triggering full malware audit." }
    elseif ($Til -eq 1 -and $score -lt 65) { $triggerMalwareScan = $true; Trace "INTEL: Low Confidence Score ($score%). Triggering targeted malware audit." }

    $malwareFound = $false
    if ($triggerMalwareScan) {
        Trace "Starting Multi-Engine Malware Intelligence Audit..."
        $malwareHits = New-Object System.Collections.Generic.List[string]
        $hashesToCheck = New-Object System.Collections.Generic.List[object]

        if (![string]::IsNullOrEmpty($itemData.by)) { 
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($itemData.by)
            $sha256 = [System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)).Replace("-", "").ToLower()
            $md5 = [System.BitConverter]::ToString([System.Security.Cryptography.MD5]::Create().ComputeHash($bytes)).Replace("-", "").ToLower()
            [void]$hashesToCheck.Add(@{ type = "Body"; name = "body"; sha256 = $sha256; md5 = $md5 })
        }
        if ($itemData.Attachments -and $itemData.Attachments.Count -gt 0) {
            foreach ($at in $itemData.Attachments) {
                if ($at.hash -and $at.hash -ne "N/A") { 
                    $md5At = if ($at.md5 -and $at.md5 -ne "N/A") { $at.md5 } else { "" }
                    [void]$hashesToCheck.Add(@{ type = "Attachment"; name = $at.name; sha256 = $at.hash; md5 = $md5At })
                }
            }
        }

        foreach ($hItem in $hashesToCheck) {
            # Redundancy 1: VirusTotal API (when user configured key)
            if (![string]::IsNullOrEmpty($Vk) -and $Vk -ne "MASKED_FOR_SECURITY" -and $Vk.Length -ge 16 -and $hItem.sha256) {
                try {
                    $u = "https://www.virustotal.com/api/v3/files/$($hItem.sha256)"
                    $hHeaders = @{ "x-apikey" = $Vk }
                    $r = Invoke-RestMethod -Uri $u -Headers $hHeaders -Method Get -TimeoutSec 4 -ErrorAction SilentlyContinue
                    if ($r -and $r.data -and $r.data.attributes -and $r.data.attributes.last_analysis_stats) {
                        if ($r.data.attributes.last_analysis_stats.malicious -gt 0) {
                            $malwareFound = $true
                            [void]$malwareHits.Add("VT_MALICIOUS($($hItem.name))")
                            Trace "INTEL: VirusTotal detected malware for $($hItem.name)!"
                        }
                    }
                } catch {}
            }
            # Redundancy 2: Team Cymru Malware Hash Database (Zero-Config DNS API, no key required)
            if (!$malwareFound -and $hItem.md5) {
                try {
                    $d = Resolve-DnsName -Name "$($hItem.md5).malware.hash.cymru.com" -Type TXT -ErrorAction SilentlyContinue
                    if ($d -and ($d.Strings -match "^\d+\s+\d+" -or $d.Strings -match "127\.0\.0\.2")) {
                        $malwareFound = $true
                        [void]$malwareHits.Add("CYMRU_MALWARE_DB($($hItem.name))")
                        Trace "INTEL: Team Cymru Malware Hash Database match for $($hItem.name)!"
                    }
                } catch {}
            }
        }
        if ($malwareHits.Count -gt 0) { [void]$hits.Add("ADVANCED_INTEL(" + ([string]::Join(", ", $malwareHits)) + ")") }
    }

    if ($malwareFound) {
        $score = [Math]::Min($score, 10)
    }

    Trace "ANALYSIS COMPLETE: Consolidating Verdict..."
    $threshold = if ($null -ne $ru.spamThresholdPercent) { [int]$ru.spamThresholdPercent } else { 50 }
    $verdict = if ($malwareFound -or $score -le 20) { "Malicious" } elseif ($score -le $threshold) { "Spam" } else { "Safe" }
    Trace "FINAL VERDICT: [$verdict] | Integrity Score: $score%"
    
    return @{ 
        mv = if ($verdict -eq "Malicious") { "MALICIOUS" } elseif ($verdict -eq "Spam") { "SPAM" } else { "CLEAN" }
        verdict = $verdict
        score = $score
        tier = ([string]::Join(" | ", $hits) -replace "^$", "Clean (Passed Security Checks)")
        action = if ($verdict -ne "Safe") { "Quarantined" } else { "None" }
    }
}

function Register-OnAccess-Watchers {
    param($Namespace)
    if ($Global:Watchers.Count -gt 0) { return }
    Log-Progress "On-Access: Initializing live mailbox event listener..."
    $stores = $null
    try {
        $stores = $Namespace.Stores
        foreach ($S in $stores) {
            Send-Heartbeat
            $inbox = $null
            try { 
                $inbox = Invoke-OutlookMethod { $S.GetDefaultFolder(6) } 
            } catch {}
            if ($inbox) {
                $items = $inbox.Items
                $subId = "MOS_Inbox_ItemAdd_$($S.StoreID)"
                Unregister-Event -SourceIdentifier $subId -ErrorAction SilentlyContinue
                [void](Register-ObjectEvent -InputObject $items -EventName "ItemAdd" -SourceIdentifier $subId -Action {
                    param($item)
                    try {
                        if ($item -and $item.MessageClass -eq "IPM.Note") {
                            $Global:ScanQueue.Enqueue($item.EntryID)
                        }
                    } finally {
                        Release-Com $item
                    }
                })
                [void]$Global:EventSubscriberIds.Add($subId)
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
}

function Sweep-Unread-Items {
    param($Namespace)
    Log-Progress "On-Access: Initializing protection. Performing sweep of unread items..."
    $stores = $null
    try {
        $stores = $Namespace.Stores
        foreach ($S in $stores) {
            Send-Heartbeat
            $inbox = $null
            try { 
                $inbox = Invoke-OutlookMethod { $S.GetDefaultFolder(6) } 
            } catch {}
            if ($inbox) {
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
                finally { Release-Com $unread; Release-Com $inbox }
            }
            Release-Com $S
        }
    } catch {}
    finally { Release-Com $stores }
}

$HeartbeatTimer = $null
try {
    # 1. SEND IMMEDIATE HEARTBEAT & READ CONFIG
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

    # 2. START AUTONOMOUS BACKGROUND HEARTBEAT TIMER
    $HeartbeatTimer = New-Object System.Timers.Timer(5000)
    $HeartbeatTimer.AutoReset = $true
    [void](Register-ObjectEvent -InputObject $HeartbeatTimer -EventName Elapsed -SourceIdentifier "HeartbeatTimerEvent" -Action {
        try { Send-Heartbeat } catch {}
    })
    $HeartbeatTimer.Start()

    # 3. INITIALIZE MAPI NAMESPACE & EXCLUSIONS
    $N = $O.GetNamespace("MAPI")
    if ($null -eq $Global:ExcludedFolderIds) { $Global:ExcludedFolderIds = New-Object System.Collections.Generic.HashSet[string] }
    Init-Exclusions $N

    # 4. ON-ACCESS PROTECTION: REGISTER LISTENERS
    if ($Ex.onAccessEnabled -ne $false) {
        Register-OnAccess-Watchers $N
        if ($Ex.mode -ne "History") {
            Sweep-Unread-Items $N
        }
    }

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
                $root = $null
                try {
                    $root = Invoke-OutlookMethod { $S.GetRootFolder() }
                    if ($root) {
                        $rfSubs = $null
                        try {
                            $rfSubs = $root.Folders
                            if ($rfSubs) {
                                foreach ($sub in $rfSubs) {
                                    $subItems = $null
                                    try {
                                        $subItems = $sub.Items
                                        $totalCount += $subItems.Count
                                        $stack.Push(@{
                                            FolderId = $sub.EntryID
                                            StoreId = $S.StoreID
                                            FolderName = $sub.Name
                                            DefaultItemType = $sub.DefaultItemType
                                        })
                                    } catch {}
                                    finally { Release-Com $subItems; Release-Com $sub }
                                }
                            }
                        } finally { Release-Com $rfSubs }
                    }
                } catch {}
                finally { Release-Com $root; Release-Com $S }
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
                        # MANDATORY CHRONOLOGICAL SORTING: NEWEST TO OLDEST
                        try {
                            Invoke-OutlookMethod { $fItems.Sort("[ReceivedTime]", $true) }
                        } catch {}
                        $entryIds = New-Object System.Collections.Generic.List[string]
                        $itemCount = 0
                        try { $itemCount = Invoke-OutlookMethod { $fItems.Count } } catch {}
                        $limit = if ($Ex.deepHistoryScanEnabled -eq $true) { $itemCount } elseif ($Ex.onDemandLimit -gt 0) { [int]$Ex.onDemandLimit } else { 1000 }
                        $scanTarget = [Math]::Min($itemCount, $limit)
                        if ($scanTarget -gt 0) {
                            for ($i = 1; $i -le $scanTarget; $i++) {
                                $it = $null
                                try {
                                    $it = Invoke-OutlookMethod { $fItems.Item($i) }
                                    if ($it) {
                                        $eid = $null
                                        try { $eid = $it.EntryID } catch {}
                                        if (![string]::IsNullOrEmpty($eid)) {
                                            [void]$entryIds.Add($eid)
                                        }
                                    }
                                } catch {}
                                finally { Release-Com $it }
                            }
                        }
                        if ($entryIds.Count -eq 0 -and $scanTarget -gt 0) {
                            $idx = 0
                            foreach ($it in $fItems) {
                                try {
                                    $idx++
                                    if ($idx -gt $scanTarget) { break }
                                    if ($it) {
                                        $eid = $null
                                        try { $eid = $it.EntryID } catch {}
                                        if (![string]::IsNullOrEmpty($eid)) {
                                            [void]$entryIds.Add($eid)
                                        }
                                    }
                                } catch {}
                                finally { Release-Com $it }
                            }
                        }
                        Release-Com $fItems
                        $fItems = $null

                        foreach ($entryId in $entryIds) {
                            # --- ON-ACCESS PRIORITY PREEMPTION ---
                            while ($Global:ScanQueue.Count -gt 0) {
                                $qId = $null
                                if (!$Global:ScanQueue.TryDequeue([ref]$qId)) { break }
                                if ($Global:ReleasedFingerprints.Contains($qId)) { continue }
                                $qt = $null
                                try {
                                    $qt = Invoke-OutlookMethod { $N.GetItemFromID($qId) }
                                    if ($qt) {
                                        $qfData = Parse-Forensics $qt
                                        $qfp = Get-Fingerprint -item $qt -ip $qfData.ip
                                        if ($ps.Contains($qfp) -or $Global:ReleasedFingerprints.Contains($qfp)) { continue }
                                        $qdomain = "unknown"
                                        if ($qfData.from -match "@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})") { $qdomain = $Matches[1].ToLower() }
                                        $qStoreId = ""
                                        try { if ($qt.Parent) { $qStoreId = $qt.Parent.StoreID } } catch {}
                                        $qitemData = @{ 
                                            Id=$qId; Su=(Get-ItemSubject $qt); Se=$qfData.from; IP=$qfData.ip; Do=$qdomain; 
                                            Hs=$qfData.headers; by=$qfData.body; Finger=$qfp; Attachments=$qfData.attachments;
                                            To=$qfData.to; Cc=$qfData.cc; Date=$qfData.date; Time=$qfData.time; Timestamp=$qfData.timestamp;
                                            StoreId=$qStoreId
                                        }
                                        $qpsi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($qitemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk).AddArgument($Til)
                                        $qpsi.RunspacePool = $RunspacePool
                                        [void]$CurrentBatch.Add(@{ PS=$qpsi; Handle=$qpsi.BeginInvoke(); Data=$qitemData })
                                        Process-Batch
                                    }
                                } catch {} finally { Release-Com $qt }
                            }

                            $t = $null
                            try {
                                $currentCount++
                                $t = Invoke-OutlookMethod { $N.GetItemFromID($entryId, $fDesc.StoreId) }
                                if (!$t) { $t = Invoke-OutlookMethod { $N.GetItemFromID($entryId) } }
                                if (!$t) { continue }
                                $fData = Parse-Forensics $t
                                $fp = Get-Fingerprint -item $t -ip $fData.ip
                                if ($ps.Contains($fp) -or $Global:ReleasedFingerprints.Contains($fp)) { continue }
                                $domain = "unknown"
                                if ($fData.from -match "@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})") { $domain = $Matches[1].ToLower() }
                                $itemData = @{ 
                                    Id=$entryId; Su=(Get-ItemSubject $t); Se=$fData.from; IP=$fData.ip; Do=$domain; 
                                    Hs=$fData.headers; by=$fData.body; Finger=$fp; Attachments=$fData.attachments;
                                    To=$fData.to; Cc=$fData.cc; Date=$fData.date; Time=$fData.time; Timestamp=$fData.timestamp;
                                    StoreId=$fDesc.StoreId
                                }
                                $psi = [powershell]::Create().AddScript($AnalysisScript).AddArgument($itemData).AddArgument($sk).AddArgument($ru).AddArgument($wl).AddArgument($bl).AddArgument($Vk).AddArgument($Til)
                                $psi.RunspacePool = $RunspacePool
                                [void]$CurrentBatch.Add(@{ PS=$psi; Handle=$psi.BeginInvoke(); Data=$itemData })
                                
                                if ($currentCount % 5 -eq 0) {
                                    Send-Status -status "SCANNING" -details "Analyzing mailbox: $currentCount / $totalCount" -count $currentCount -total $totalCount -currentFolder $fDesc.FolderName
                                }

                                if ($CurrentBatch.Count -ge 8) { Process-Batch; Start-Sleep -Milliseconds 25 }
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
        if ($Ex.onAccessEnabled -ne $false) {
            Register-OnAccess-Watchers $N
        }
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
                    if ($upd.whitelist) { $wl = $upd.whitelist }
                    if ($upd.blacklist) { $bl = $upd.blacklist }
                    if ($upd.spamKeywords) { $sk = $upd.spamKeywords }
                    if ($upd.rubrics) { $ru = $upd.rubrics }
                    if ($upd.vtKey) { $Vk = $upd.vtKey }
                    if ($upd.threatIntelligenceLevel) { $Til = $upd.threatIntelligenceLevel }
                }
            }
            $ReadTask = $StdInReader.ReadLineAsync()
        }

        while ($Global:ScanQueue.Count -gt 0) {
            $id = $null
            if (!$Global:ScanQueue.TryDequeue([ref]$id)) { break }
            if ($Global:ReleasedFingerprints.Contains($id)) { continue }
            $lt = $null
            try {
                $lt = Invoke-OutlookMethod { $N.GetItemFromID($id) }
                if ($lt) {
                    $fData = Parse-Forensics $lt
                    $fp = Get-Fingerprint -item $lt -ip $fData.ip
                    if ($ps.Contains($fp) -or $Global:ReleasedFingerprints.Contains($fp)) { continue }
                    $domain = "unknown"
                    if ($fData.from -match "@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})") { $domain = $Matches[1].ToLower() }
                    $qStoreId = ""
                    try { if ($lt.Parent) { $qStoreId = $lt.Parent.StoreID } } catch {}
                    $itemData = @{ 
                        Id=$id; Su=(Get-ItemSubject $lt); Se=$fData.from; IP=$fData.ip; Do=$domain; 
                        Hs=$fData.headers; by=$fData.body; Finger=$fp; Attachments=$fData.attachments;
                        To=$fData.to; Cc=$fData.cc; Date=$fData.date; Time=$fData.time; Timestamp=$fData.timestamp;
                        StoreId=$qStoreId
                    }
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
    if ($null -ne $HeartbeatTimer) {
        try {
            $HeartbeatTimer.Stop()
            Unregister-Event -SourceIdentifier "HeartbeatTimerEvent" -ErrorAction SilentlyContinue
            $HeartbeatTimer.Dispose()
        } catch {}
        $HeartbeatTimer = $null
    }
    if ($null -ne $Global:EventSubscriberIds) {
        foreach ($id in $Global:EventSubscriberIds) {
            try { Unregister-Event -SourceIdentifier $id -ErrorAction SilentlyContinue } catch {}
        }
    }
    if ($null -ne $Global:Watchers) {
        foreach ($w in $Global:Watchers) {
            try { Release-Com $w.Items } catch {}
            try { Release-Com $w.Folder } catch {}
        }
        $Global:Watchers.Clear()
    }
    if ($null -ne $Global:EventSubscriberIds) {
        $Global:EventSubscriberIds.Clear()
    }
    Release-Com $N
    Release-Com $O
    if ($null -ne $RunspacePool) {
        try { $RunspacePool.Close(); $RunspacePool.Dispose() } catch {}
    }
    if ($null -ne $Global:OutlookComMutex) {
        try { $Global:OutlookComMutex.Dispose() } catch {}
    }
    if ($null -ne $Global:InstanceMutex) {
        try {
            $Global:InstanceMutex.ReleaseMutex()
            $Global:InstanceMutex.Dispose()
        } catch {}
    }
}
}
