function Get-DSRegStatus {
    $raw = dsregcmd /status | Select-String ":" |
        ForEach-Object {
            $parts = $_.Line.Trim() -split "\s*:\s*", 2
            if ($parts.Count -eq 2) {
                $key   = $parts[0].Trim() -replace '\s+', ''  # Remove all spaces from key
                $value = $parts[1].Trim()
                [PSCustomObject]@{ Key = $key; Value = $value }
            }
        }

    $obj = [PSCustomObject]@{}
    foreach ($item in $raw) {
        # Skip duplicates — first occurrence wins
        if (-not ($obj.PSObject.Properties.Name -contains $item.Key)) {
            $obj | Add-Member -NotePropertyName $item.Key -NotePropertyValue $item.Value
        }
    }
    return $obj
}
