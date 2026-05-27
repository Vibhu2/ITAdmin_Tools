
<#
.SYNOPSIS
    Flattens a complex PowerShell object into a single-level PSCustomObject with scalar property values.

.DESCRIPTION
    ConvertTo-FlatObject accepts any PowerShell object via the pipeline or direct parameter and
    converts all of its properties into flat, scalar-compatible values. This makes the output
    safe for tabular consumers such as Export-Csv, Format-Table, or ConvertTo-Json where nested
    objects and collections would otherwise produce unhelpful representations like
    "System.Object[]" or "@{Key=Value}".

    The function applies the following serialization rules to each property value:

      - Collections (arrays, lists, any IEnumerable except strings):
            Joined into a single semicolon-delimited string using '; ' as the separator.

      - Nested objects (hashtables or PSCustomObjects):
            Serialized to a compact single-line JSON string with a maximum depth of 3.

      - Scalar values (strings, integers, booleans, DateTimes, nulls, etc.):
            Passed through as-is without modification.

    Strings are explicitly excluded from collection handling because the .NET string type
    implements IEnumerable<char>. Without this guard, a string value would be character-joined
    (e.g., "hello" -> "h; e; l; l; o").

    Null input objects are silently skipped with no output and no error.

.PARAMETER InputObject
    The object to flatten. Accepts any PowerShell object type including PSCustomObject,
    deserialized objects (from Import-Csv, ConvertFrom-Json, Get-ADUser, etc.), and
    strongly-typed .NET objects. Accepts pipeline input.

.INPUTS
    System.Object
        Any object passed via the pipeline or the InputObject parameter.

.OUTPUTS
    System.Management.Automation.PSCustomObject
        A flat PSCustomObject where every property value is a scalar (string, number,
        boolean, or null). Property order is preserved from the source object.

.EXAMPLE
    # Flatten a single object with mixed property types
    $obj = [PSCustomObject]@{
        Name    = "Server01"
        Tags    = @("web", "prod", "east")
        Network = @{ IP = "10.0.0.1"; VLAN = 100 }
        Uptime  = 99.9
    }

    $obj | ConvertTo-FlatObject

    # Output:
    # Name     : Server01
    # Tags     : web; prod; east
    # Network  : {"IP":"10.0.0.1","VLAN":100}
    # Uptime   : 99.9

.EXAMPLE
    # Flatten a collection of objects and export to CSV
    Get-ADComputer -Filter * -Properties * |
        ConvertTo-FlatObject |
        Export-Csv -Path "C:\Reports\computers.csv" -NoTypeInformation

    # Without flattening, multi-value AD attributes would appear as "System.String[]" in the CSV.

.EXAMPLE
    # Flatten objects from a JSON import
    $data = Get-Content ".\devices.json" | ConvertFrom-Json
    $data | ConvertTo-FlatObject | Format-Table -AutoSize

.NOTES
    Name        : ConvertTo-FlatObject
    Author      : Vibhu Bhatnagar
    Version     : 1.0.0

    - Nested object serialization uses ConvertTo-Json with -Depth 3. Objects deeper than
      3 levels will be truncated by ConvertTo-Json with no error from this function.
    - This function does not recurse into nested objects; it serializes them as JSON strings.
      If you need full recursive flattening with dotted-path property names
      (e.g., "Network.IP"), that requires a separate recursive implementation.
    - PSObject.Properties reflection works on any object type, including deserialized
      CIM/WMI instances, Active Directory objects, and XML elements.
    - Compatible with PowerShell 5.1 and PowerShell 7+.

.LINK
    ConvertTo-Json
    Export-Csv
    Select-Object
#>
function ConvertTo-FlatObject {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param([Parameter(ValueFromPipeline)][object]$InputObject)
    process {
        if ($null -eq $InputObject) { return }
        $props = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $val = $prop.Value
            $props[$prop.Name] = switch ($true) {
                ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                    ($val | ForEach-Object { $_ }) -join '; '; break
                }
                ($val -is [hashtable] -or $val -is [PSCustomObject]) {
                    $val | ConvertTo-Json -Compress -Depth 3; break
                }
                default { $val }
            }
        }
        [PSCustomObject]$props
    }
}
#endregion
