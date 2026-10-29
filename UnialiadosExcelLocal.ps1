param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Read', 'Upsert')]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$SheetName,

    [string]$PayloadPath = '',
    [string]$KeyColumn = '',
    [string]$BackupDirectory = '',
    [string]$BackupTag = '',
    [int]$LockTimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Write-EncodedJson {
    param([object]$Value)
    $json = $Value | ConvertTo-Json -Depth 20 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    Write-Output ('BASE64:' + [Convert]::ToBase64String($bytes))
}

function Normalize-Key {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}

function Get-SafeFilePart {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
    }
    return ($Value -replace '[^A-Za-z0-9_.-]', '_')
}

function Ensure-Backup {
    param(
        [string]$SourcePath,
        [string]$Directory,
        [string]$Tag
    )

    if ([string]::IsNullOrWhiteSpace($Directory)) { return '' }
    if (-not (Test-Path -LiteralPath $Directory)) {
        New-Item -ItemType Directory -Path $Directory -Force | Out-Null
    }

    $safeTag = Get-SafeFilePart -Value $Tag
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $extension = [System.IO.Path]::GetExtension($SourcePath)
    $backupPath = Join-Path $Directory ($baseName + '_backup_' + $safeTag + $extension)

    if (-not (Test-Path -LiteralPath $backupPath)) {
        Copy-Item -LiteralPath $SourcePath -Destination $backupPath -Force
    }

    Get-ChildItem -LiteralPath $Directory -Filter ($baseName + '_backup_*' + $extension) -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip 30 |
        Remove-Item -Force

    return $backupPath
}

if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "Workbook not found: $WorkbookPath"
}

$mutex = New-Object System.Threading.Mutex($false, 'Local\UnialiadosExcelLocal')
$lockTaken = $false
$excel = $null
$workbook = $null
$worksheet = $null
$usedRange = $null

try {
    try {
        $lockTaken = $mutex.WaitOne([TimeSpan]::FromSeconds($LockTimeoutSeconds))
    } catch [System.Threading.AbandonedMutexException] {
        $lockTaken = $true
    }

    if (-not $lockTaken) {
        throw "Timed out waiting for the local Excel lock after $LockTimeoutSeconds seconds."
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3

    $readOnly = $Action -eq 'Read'
    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $readOnly)

    if ($Action -eq 'Upsert' -and $workbook.ReadOnly) {
        throw 'The workbook is locked or open in Excel. Close it before the scheduled automation runs.'
    }

    $worksheet = $workbook.Worksheets.Item($SheetName)
    $usedRange = $worksheet.UsedRange
    $values = $usedRange.Value2
    $rowCount = [int]$usedRange.Rows.Count
    $columnCount = [int]$usedRange.Columns.Count

    $headers = New-Object System.Collections.Generic.List[string]
    $headerIndexes = @{}
    for ($column = 1; $column -le $columnCount; $column++) {
        $header = ([string]$values[1, $column]).Trim()
        if ([string]::IsNullOrWhiteSpace($header)) { continue }
        $headers.Add($header)
        $headerIndexes[$header] = $column
    }

    if ($Action -eq 'Read') {
        $rows = New-Object System.Collections.Generic.List[object]
        for ($row = 2; $row -le $rowCount; $row++) {
            $record = [ordered]@{}
            $hasValue = $false
            foreach ($header in $headers) {
                $value = $values[$row, $headerIndexes[$header]]
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $hasValue = $true
                }
                $record[$header] = $value
            }
            if ($hasValue) { $rows.Add([pscustomobject]$record) }
        }

        Write-EncodedJson -Value ($rows.ToArray())
        return
    }

    if ([string]::IsNullOrWhiteSpace($PayloadPath) -or -not (Test-Path -LiteralPath $PayloadPath -PathType Leaf)) {
        throw "Payload file not found: $PayloadPath"
    }
    if ([string]::IsNullOrWhiteSpace($KeyColumn) -or -not $headerIndexes.ContainsKey($KeyColumn)) {
        throw "Key column '$KeyColumn' was not found in worksheet '$SheetName'."
    }

    $backupPath = Ensure-Backup -SourcePath $WorkbookPath -Directory $BackupDirectory -Tag $BackupTag
    $payloadText = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8
    $payload = $payloadText | ConvertFrom-Json
    $payloadItems = if ($payload -is [System.Array]) { @($payload) } else { @($payload) }
    $updated = 0

    foreach ($item in $payloadItems) {
        $properties = @($item.PSObject.Properties)
        $keyProperty = $properties | Where-Object { $_.Name -eq $KeyColumn } | Select-Object -First 1
        if ($null -eq $keyProperty) {
            throw "Payload does not contain key column '$KeyColumn'."
        }

        $targetKey = Normalize-Key -Value $keyProperty.Value
        if ([string]::IsNullOrWhiteSpace($targetKey)) {
            throw "Payload contains an empty '$KeyColumn'."
        }

        $targetRow = 0
        $keyIndex = [int]$headerIndexes[$KeyColumn]
        for ($row = 2; $row -le $rowCount; $row++) {
            if ((Normalize-Key -Value $values[$row, $keyIndex]) -eq $targetKey) {
                $targetRow = $row
                break
            }
        }

        if ($targetRow -eq 0) {
            throw "No row with $KeyColumn=$targetKey exists in worksheet '$SheetName'."
        }

        foreach ($property in $properties) {
            if (-not $headerIndexes.ContainsKey($property.Name)) { continue }
            $cellValue = $property.Value
            if ($null -eq $cellValue) { $cellValue = '' }
            $cell = $null
            try {
                $cell = $worksheet.Cells.Item([int]$targetRow, [int]$headerIndexes[$property.Name])
                if ($cellValue -is [byte] -or
                    $cellValue -is [int16] -or
                    $cellValue -is [int32] -or
                    $cellValue -is [int64] -or
                    $cellValue -is [single] -or
                    $cellValue -is [double] -or
                    $cellValue -is [decimal]) {
                    $cell.Value2 = [double]$cellValue
                } else {
                    $cell.Value2 = [string]$cellValue
                }
            } catch {
                throw "Could not write column '$($property.Name)' for $KeyColumn=$targetKey. $($_.Exception.Message)"
            } finally {
                Release-ComObject -Object $cell
            }
        }
        $updated++
    }

    $workbook.Save()
    Remove-Item -LiteralPath $PayloadPath -Force -ErrorAction SilentlyContinue
    Write-EncodedJson -Value ([ordered]@{
        ok = $true
        updated = $updated
        sheet = $SheetName
        backup = $backupPath
    })
} catch {
    Write-Error ($_.Exception.Message + ' | ' + $_.ScriptStackTrace)
    exit 1
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }

    Release-ComObject -Object $usedRange
    Release-ComObject -Object $worksheet
    Release-ComObject -Object $workbook
    Release-ComObject -Object $excel

    if ($lockTaken) {
        try { $mutex.ReleaseMutex() } catch {}
    }
    $mutex.Dispose()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
