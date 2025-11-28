
$text = @"
EloDe Farm 5 was atta 23:04:46 23:04:22 23:04: 16 23:03:47 23:03:39 VARANASI (X:309 y:457). Globe Farm CCC5 was attacked by GLOBE VARANASI (X:309 Y:457). Globe Farm CCC5 was attacked by GLOBE VARANASI (X:309 y:457). Globe Farm CCC5 was attacked by GLOBE VARANASI (X:309 Y:457). Globe Farm CCC5 was attacked by GLOBE VARANASI (X:309 y:457). Globe Farm CCC4 was attacked by GLOBE VARANASI (X:312 Y:455). Globe Farm CCC4 was attacked by GLOBE VARANASI (X:312 Y:455). Globe Farm CCC4 was attacked by GLOBE VARANASI (X:312 Y:455). KO J Farm FFF3 was attacked by ssolove y:537). Globe Farm CCC4 was attacked by GLOBE VARANASI (X:312 Y:455). Globe Farm CCC4 was attacked by GLOBE VARANASI (X:312 Y:455). Globe Farm CCC4 was attacked by GLOBE VARANASI (X:312 Y:455). Globe Farm CCC4 was attacked by GLOBE VARANASI (X:312 Y:455).
"@

function Parse-AttackLogs {
    param([string]$Text)

    $clean = $Text -replace '[^\u0020-\u007E]', ' '
    $clean = $clean -replace '\s+', ' '

    Write-Host "Cleaned Text: $clean"

    $logDate = ''
    $dateMatch = [regex]::Match($clean, '\d{4}-\d{2}-\d{2}')
    if ($dateMatch.Success) { $logDate = $dateMatch.Value }

    $entries = New-Object System.Collections.Generic.List[pscustomobject]

    $timeTokens = [regex]::Matches($clean, '\b\d{2}:\d{2}:\d{2}\b') | ForEach-Object { $_.Value.Trim() }
    $timeQueue = New-Object System.Collections.Generic.Queue[string]
    foreach ($t in $timeTokens) { $timeQueue.Enqueue($t) }

    Write-Host "Time Tokens: $($timeTokens -join ', ')"

    $attackPattern = '(?:[\u0000-\u001f·•◆▶▷►]*\s*)?(?:(?<time>\d{2}:\d{2}:\d{2})\s*)?(?:Logs\s+)?(?<farm>[A-Za-z0-9\s]+?)\s+was attacked by\s+(?<attacker>[A-Za-z0-9\s]+?)\s*\(X[:：]?[^\d]{0,3}(?<x>\d+)\s*Y[:：]?[^\d]{0,3}(?<y>\d+)\)'
    $matches = [regex]::Matches($clean, $attackPattern, 'IgnoreCase')
    
    Write-Host "Match Count: $($matches.Count)"

    foreach ($m in $matches) {
        $timeVal = $m.Groups['time'].Value
        if (-not $timeVal) {
            $timeVal = if ($timeQueue.Count -gt 0) { $timeQueue.Dequeue() } else { '' }
        }
        if ($logDate) {
            if ($timeVal) { $timeVal = "$logDate $timeVal" } else { $timeVal = $logDate }
        }
        $entries.Add([pscustomobject]@{
            Time     = $timeVal
            Farm     = $m.Groups['farm'].Value.Trim()
            Attacker = $m.Groups['attacker'].Value.Trim()
            X        = $m.Groups['x'].Value.Trim()
            Y        = $m.Groups['y'].Value.Trim()
        }) | Out-Null
    }

    $dedup = $entries | Select-Object Time,Farm,Attacker,X,Y -Unique
    return $dedup
}

$result = Parse-AttackLogs -Text $text
$result | Format-Table
