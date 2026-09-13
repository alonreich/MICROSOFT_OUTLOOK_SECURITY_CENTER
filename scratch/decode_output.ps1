$bytes = [System.IO.File]::ReadAllBytes("$PSScriptRoot\scanner_output.bin")
$offset = 0
$packets = 0
$statusCounts = @{}
while ($offset + 4 -le $bytes.Length) {
    $len = [System.BitConverter]::ToInt32($bytes, $offset)
    $offset += 4
    if ($offset + $len -gt $bytes.Length) { break }
    $str = [System.Text.Encoding]::UTF8.GetString($bytes, $offset, $len)
    $offset += $len
    $packets++
    $obj = $str | ConvertFrom-Json
    $st = if ($obj.status) { $obj.status } elseif ($obj.type) { $obj.type } else { "UNKNOWN" }
    if (!$statusCounts.ContainsKey($st)) { $statusCounts[$st] = 0 }
    $statusCounts[$st]++
    if ($st -ne "heartbeat" -and $st -ne "INFO") {
        Write-Output "Packet $packets : Status=$st Details=$($obj.details) Count=$($obj.count)/$($obj.total) Subject=$($obj.subject)"
    }
}
Write-Output "--- SUMMARY ---"
Write-Output "Total packets: $packets"
$statusCounts.GetEnumerator() | ForEach-Object { Write-Output "$($_.Key): $($_.Value)" }
