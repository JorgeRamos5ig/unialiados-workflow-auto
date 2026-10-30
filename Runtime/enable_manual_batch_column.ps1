$ErrorActionPreference = 'Stop'
$path = Join-Path (Split-Path -Parent $PSScriptRoot) 'Unialiados_Operacion_Unificada.xlsm'
$backup = Join-Path (Split-Path -Parent $PSScriptRoot) 'Respaldos\Unialiados_Operacion_Unificada.before-manual-batch-20261030.xlsm'
if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Workbook not found: $path" }
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $backup) | Out-Null
Copy-Item -LiteralPath $path -Destination $backup -Force

$excel = $null
$book = $null
$base = $null
$upload = $null
$guide = $null
$table = $null
$newColumn = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path, 0, $false)
    if ($book.ReadOnly) { throw 'Workbook opened read-only; close Excel and retry.' }

    $base = $book.Worksheets.Item('Base de Datos')
    $table = $base.ListObjects.Item('BaseLeads')
    $columnExists = $false
    foreach ($column in $table.ListColumns) {
        if ([string]$column.Name -eq 'Enviar ahora') { $columnExists = $true; break }
    }
    if (-not $columnExists) {
        $newColumn = $table.ListColumns.Add()
        $newColumn.Name = 'Enviar ahora'
    } else {
        $newColumn = $table.ListColumns.Item('Enviar ahora')
    }
    if ($newColumn.DataBodyRange -ne $null) {
        $body = $newColumn.DataBodyRange
        $idRange = $table.ListColumns.Item('ID_Lead').DataBodyRange
        for ($i = 1; $i -le $body.Rows.Count; $i++) {
            if ([string]$idRange.Cells.Item($i, 1).Value2) {
                $body.Cells.Item($i, 1).Value2 = 'No'
            } elseif ($i -eq 1 -and $table.ListRows.Count -eq 1) {
                $body.Cells.Item($i, 1).Value2 = 'No'
            }
        }
        $validation = $body.Validation
        try { $validation.Delete() } catch {}
        $validation.Add(3, 1, 1, 'No,Si')
        $validation.IgnoreBlank = $true
        $validation.InCellDropdown = $true
    }

    # La hoja de carga avisa si un contacto todavia no tiene una campana manual asignada.
    $upload = $book.Worksheets.Item('CARGA MASIVA')
    $upload.Range('N2').Formula = '=IF(A2="","",IF(AND(B2="",C2=""),"EMPRESA SOLA",IF(OR(B2="",C2=""),"REVISAR CONTACTO/CORREO",IF(OR(ISERROR(SEARCH("@",C2)),ISERROR(SEARCH(".",C2))),"EMAIL INVALIDO",IF(OR(H2="",H2="Auto"),"SELECCIONE CAMPANA","LISTO")))))'
    $upload.Range('N2:N1001').FillDown()
    $campaignValidation = $upload.Range('H2:H1001').Validation
    try { $campaignValidation.Delete() } catch {}
    $campaignValidation.Add(3, 1, 1, 'Unialiados,AXA Colpatria,Unisoft')
    $campaignValidation.IgnoreBlank = $true
    $campaignValidation.InCellDropdown = $true

    $guide = $book.Worksheets.Item('GUIA DE CARGA')
    $guide.Range('A18').Value2 = 'IMPORTANTE: cargue y revise los contactos; asigne campana manualmente. En Base de Datos filtre por campana y marque Enviar ahora = Si solo para los leads aprobados. Revise el asunto y el cuerpo de la prueba, guarde y cierre Excel antes de iniciar n8n manualmente. No edite los IDs ni columnas tecnicas.'

    $book.Save()
    Write-Output "Workbook updated. BaseLeads columns=$($table.ListColumns.Count); batch column present=$([string]$table.ListColumns.Item('Enviar ahora').Name -eq 'Enviar ahora'); backup=$backup"
} finally {
    if ($book) { $book.Close($true) }
    if ($excel) { $excel.Quit() }
    foreach ($com in @($newColumn, $table, $guide, $upload, $base, $book, $excel)) {
        if ($com -and [Runtime.InteropServices.Marshal]::IsComObject($com)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($com) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
