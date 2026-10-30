#requires -version 5.1
<#
  IMPORTADOR DE CONTACTOS Y EMPRESAS - UNIALIADOS
  Compatibilidad: Windows + Microsoft Excel de escritorio + Windows PowerShell 5.1.
  Este script NO modifica los nombres de hojas, las columnas ni el codigo VBA.
  Antes de escribir: valida registros, acepta correos compartidos y crea una copia de seguridad.
  Solo evita registrar dos veces el MISMO contacto (empresa + persona + correo).
  No ejecute al mismo tiempo que el flujo de n8n.
#>
param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot 'Unialiados_Operacion_Unificada.xlsm')
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

function Text([object]$Value) {
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}
function Key([object]$Value) {
    $text = (Text $Value).ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $text.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $text = $sb.ToString()
    $text = $text -replace '\b(s\s*\.?\s*a\s*\.?\s*s\s*\.?|ltda\.?|limitada)\s*$', ''
    $text = $text -replace '[^a-z0-9]+', ' '
    return ($text -replace '\s+', ' ').Trim()
}
function Contact-Key([object]$Company, [object]$Person, [object]$Email) {
    # Identifica un contacto por empresa + nombre + correo. El correo por si solo NO es unico.
    return ((Key $Company) + '|' + (Key $Person) + '|' + (Text $Email).ToLowerInvariant())
}
function Release-Com([object]$Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}
function Table-Headers([object]$Table) {
    $headers = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $Table.HeaderRowRange.Columns.Count; $i++) {
        $headers.Add((Text $Table.HeaderRowRange.Cells.Item(1, $i).Value2))
    }
    return ,$headers
}
function Read-Table([object]$Table, [object]$Headers) {
    $rows = New-Object System.Collections.Generic.List[object]
    if ($Table.ListRows.Count -eq 0) { return ,$rows }
    $vals = $Table.DataBodyRange.Value2
    $cols = [int]$Headers.Count
    $count = [int]$Table.ListRows.Count
    for ($r = 0; $r -lt $count; $r++) {
        $record = @{}
        for ($c = 0; $c -lt $cols; $c++) {
            # Excel devuelve escalar solo cuando hay una unica celda (no aplica a estas tablas).
            $record[$Headers[$c]] = $vals[($r + 1), ($c + 1)]
        }
        $rows.Add($record)
    }
    return ,$rows
}
function Get-Integer([object]$Value) {
    $valueString = Text $Value
    $num = 0
    if ([int]::TryParse($valueString, [ref]$num)) { return $num }
    if ($Value -is [double] -or $Value -is [decimal]) { return [int]$Value }
    return 0
}
function Append-Table([object]$Table, [object]$Headers, [object]$Rows) {
    if ($Rows.Count -eq 0) { return }
    $sheet = $Table.Parent
    $n = [int]$Rows.Count
    $cols = [int]$Headers.Count
    $startRow = [int]$Table.Range.Row + [int]$Table.Range.Rows.Count
    $startCol = [int]$Table.Range.Column
    $matrix = [Array]::CreateInstance([object], [int[]]@($n, $cols))
    for ($r = 0; $r -lt $n; $r++) {
        $item = $Rows[$r]
        for ($c = 0; $c -lt $cols; $c++) {
            $name = $Headers[$c]
            $v = if ($item.ContainsKey($name)) { $item[$name] } else { '' }
            if ($null -eq $v) { $v = '' }
            # Excel no debe interpretar como fórmula el contenido pegado por usuarios.
            if ($v -is [string] -and $v -match '^[=+@-]') { $v = "'" + $v }
            $matrix[$r,$c] = $v
        }
    }
    $fullStart = $null; $fullEnd = $null; $newRange = $null
    $start = $null; $end = $null; $writeRange = $null
    try {
        # Expande la tabla ANTES de escribir. Así no dependemos de la expansión
        # automática de ListObject cuando Excel detecta datos bajo la tabla.
        $fullStart = $sheet.Cells.Item([int]$Table.Range.Row, $startCol)
        $fullEnd = $sheet.Cells.Item(($startRow + $n - 1), ($startCol + $cols - 1))
        $newRange = $sheet.Range($fullStart, $fullEnd)
        $Table.Resize($newRange)
        $start = $sheet.Cells.Item($startRow, $startCol)
        $end = $sheet.Cells.Item(($startRow + $n - 1), ($startCol + $cols - 1))
        $writeRange = $sheet.Range($start, $end)
        $writeRange.Value2 = $matrix
    } finally {
        Release-Com $start; Release-Com $end; Release-Com $writeRange
        Release-Com $fullStart; Release-Com $fullEnd; Release-Com $newRange
    }
}
$mutex = $null
$lockTaken = $false
$excel = $null
$book = $null
$backupPath = ''
try {
    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "No se encontro el libro principal: $WorkbookPath" }
    $mutex = New-Object System.Threading.Mutex($false, 'Local\UnialiadosExcelLocal')
    try { $lockTaken = $mutex.WaitOne([TimeSpan]::FromSeconds(15)) }
    catch [System.Threading.AbandonedMutexException] { $lockTaken = $true }
    if (-not $lockTaken) { throw 'n8n u otro proceso esta usando el libro. Espere y reintente.' }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3  # No ejecutar macros durante la importacion; preservarlas.
    $book = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    if ($book.ReadOnly) { throw 'El libro esta abierto o bloqueado. Cierrelo y pause n8n antes de importar.' }
    $loadSheet = $book.Worksheets.Item('CARGA MASIVA')
    $used = $loadSheet.UsedRange
    $valueArray = $used.Value2
    $headerIndex = @{}
    for ($c = 1; $c -le [int]$used.Columns.Count; $c++) {
        $h = Text $valueArray[1,$c]
        if ($h) { $headerIndex[$h] = $c }
    }
    $mustHave = @('Empresa *','Contacto','Correo','Cargo','Telefono')
    # Cabecera del template original usa tilde en Telefono.
    $mustHave[4] = 'Teléfono'
    foreach ($h in $mustHave) { if (-not $headerIndex.ContainsKey($h)) { throw "Falta columna en CARGA MASIVA: $h" } }
    function Get-Field([int]$Row, [string]$Header) {
        if (-not $headerIndex.ContainsKey($Header)) { return '' }
        return (Text $valueArray[$Row, [int]$headerIndex[$Header]])
    }
    $inputRecords = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[string]
    $validCampaigns = @('Unialiados', 'AXA Colpatria', 'Unisoft', 'Auto')
    $validPriorities = @('Alta', 'Media', 'Baja')
    for ($r = 2; $r -le [int]$used.Rows.Count; $r++) {
        $company = Get-Field $r 'Empresa *'
        $person = Get-Field $r 'Contacto'
        $email = (Get-Field $r 'Correo').ToLowerInvariant()
        $cargo = Get-Field $r 'Cargo'
        $phone = Get-Field $r 'Teléfono'
        $segment = Get-Field $r 'Segmento'
        $priority = Get-Field $r 'Prioridad'
        $campaign = Get-Field $r 'Campaña'
        $site = Get-Field $r 'Sitio web'
        $linkedin = Get-Field $r 'LinkedIn contacto'
        $companyLinkedin = Get-Field $r 'LinkedIn empresa'
        $source = Get-Field $r 'Fuente'
        $notes = Get-Field $r 'Observaciones'
        $hasAny = @($company,$person,$email,$cargo,$phone,$segment,$priority,$campaign,$site,$linkedin,$companyLinkedin,$source,$notes) | Where-Object { $_ }
        if ($hasAny.Count -eq 0) { continue }
        if (-not $company) { $errors.Add("Fila $r`: falta Empresa"); continue }
        if ((Key $company).Length -eq 0) { $errors.Add("Fila $r`: empresa invalida"); continue }
        if ([bool]$person -xor [bool]$email) { $errors.Add("Fila $r`: Contacto y Correo deben estar ambos completos o ambos vacios"); continue }
        if ($email -and $email -notmatch '^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$') {
            $errors.Add("Fila $r`: correo invalido o contiene mas de una direccion ($email)"); continue
        }
        if (-not $priority) { $priority = 'Media' }
        if (-not $campaign -and $email) { $errors.Add("Fila $r`: seleccione manualmente Campaña (Unialiados, AXA Colpatria o Unisoft)"); continue }
        if (-not $campaign) { $campaign = 'Unialiados' }
        if ($email -and $campaign -eq 'Auto') { $errors.Add("Fila $r`: seleccione una campana concreta; la asignacion automatica esta deshabilitada"); continue }
        if ($validPriorities -notcontains $priority) { $errors.Add("Fila $r`: prioridad no valida ($priority)"); continue }
        if ($validCampaigns -notcontains $campaign) { $errors.Add("Fila $r`: campana no valida ($campaign)"); continue }
        $inputRecords.Add(@{
            Company=$company; Person=$person; Email=$email; Cargo=$cargo; Phone=$phone;
            Segment=$segment; Priority=$priority; Campaign=$campaign; Site=$site;
            Linkedin=$linkedin; CompanyLinkedin=$companyLinkedin; Source=$source; Notes=$notes
        })
    }
    if ($errors.Count) {
        $detail = ($errors | Select-Object -First 25) -join "`n  - "
        throw "Se encontraron $($errors.Count) errores en CARGA MASIVA. Corrijalos antes de importar:`n  - $detail"
    }
    if ($inputRecords.Count -eq 0) { throw 'No hay empresas para importar. Llene la pestaña CARGA MASIVA.' }

    $companiesTable = $book.Worksheets.Item('Empresas y Archivos').ListObjects.Item('EmpresasArchivos')
    $leadsTable = $book.Worksheets.Item('Base de Datos').ListObjects.Item('BaseLeads')
    $companyHeaders = Table-Headers $companiesTable
    $leadHeaders = Table-Headers $leadsTable
    $companies = Read-Table $companiesTable $companyHeaders
    $leads = Read-Table $leadsTable $leadHeaders
    $companyByKey = @{}
    # Si ya hay dos IDs para una misma empresa, no podemos decidir uno a ciegas.
    $ambiguousCompanyKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $existingContacts = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $companyMaxID = 0; $leadMaxID = 0
    foreach ($row in $companies) {
        $id = Get-Integer $row['ID_Empresa']; if ($id -gt $companyMaxID) { $companyMaxID = $id }
        $key = Key $row['Empresa']
        if ($key) {
            if ($companyByKey.ContainsKey($key) -and $companyByKey[$key] -ne $id) {
                [void]$ambiguousCompanyKeys.Add($key)
            } else { $companyByKey[$key] = $id }
        }
    }
    foreach ($row in $leads) {
        $id = Get-Integer $row['ID_Lead']; if ($id -gt $leadMaxID) { $leadMaxID = $id }
        # Reservar también los IDs de empresa usados por leads huérfanos.
        # De otro modo una nueva empresa podría recibir un ID todavía usado por un lead antiguo.
        $referencedCompanyId = Get-Integer $row['ID_Empresa']
        if ($referencedCompanyId -gt $companyMaxID) { $companyMaxID = $referencedCompanyId }
        $email = Text $row['Correo']
        if ($email) {
            # Permitir mismo correo en empresas/contactos diferentes.
            # Solo detectar identidad exacta que ya esta en Base de Datos.
            [void]$existingContacts.Add((Contact-Key $row['Empresa'] $row['Contacto'] $email))
        }
    }
    foreach ($rec in $inputRecords) {
        $k = Key $rec.Company
        if ($ambiguousCompanyKeys.Contains($k)) {
            throw "La empresa '$($rec.Company)' ya aparece con varios ID_Empresa en Empresas y Archivos. Revise y unifique manualmente sus registros antes de importar más contactos de esa empresa. No se modificaron datos."
        }
    }
    $newCompanies = New-Object System.Collections.Generic.List[object]
    $newLeads = New-Object System.Collections.Generic.List[object]
    $companySkipped = 0; $contactSkipped = 0
    $today = (Get-Date).ToOADate()
    foreach ($rec in $inputRecords) {
        $key = Key $rec.Company
        if (-not $companyByKey.ContainsKey($key)) {
            $companyMaxID++
            $companyByKey[$key] = $companyMaxID
            $masterCampaign = if ($rec.Campaign -eq 'Auto') { 'Unialiados' } else { $rec.Campaign }
            $newCompanies.Add(@{
                'ID_Empresa' = $companyMaxID
                'Empresa' = $rec.Company
                'Segmento' = $rec.Segment
                'Campaña principal' = $masterCampaign
                'Sitio web' = $rec.Site
                'LinkedIn empresa' = $rec.CompanyLinkedin
                'Estado de archivos' = 'Sin archivos'
                'Fecha de actualización' = $today
                'Observaciones' = $rec.Notes
                'Empresa normalizada' = $key
            })
        } else { $companySkipped++ }
        if (-not $rec.Email) { continue }  # Empresa sin contacto: solo maestro de empresas.
        # Permitir el mismo correo en distintos registros; evitar duplicados exactos
        # en el propio lote o contra lo que ya exista en la hoja Base de Datos.
        $contactIdentity = Contact-Key $rec.Company $rec.Person $rec.Email
        if (-not $existingContacts.Add($contactIdentity)) { $contactSkipped++; continue }
        $leadMaxID++
        $newLeads.Add(@{
            'ID_Lead' = $leadMaxID
            'Estado' = 'Nunca Contactado'
            'N° de contactos realizados' = 0
            'Empresa' = $rec.Company
            'Contacto' = $rec.Person
            'Cargo' = $rec.Cargo
            'Correo' = $rec.Email
            'Teléfono' = $rec.Phone
            'LinkedIn' = $rec.Linkedin
            'Segmento' = $rec.Segment
            'Prioridad' = $rec.Priority
            'Etapa' = 'Prospecto'
            'Campaña asignada' = $rec.Campaign
            'Fuente' = $rec.Source
            'Fecha de creación del lead' = $today
            'Excluir (Sí/No)' = 'No'
            'Enviar ahora' = 'No'
            'Observaciones comerciales' = $rec.Notes
            'ID_Empresa' = $companyByKey[$key]
            'Procesando' = 'No'
            'Requiere revision' = 'No'
        })
    }
    if ($newCompanies.Count -eq 0 -and $newLeads.Count -eq 0) {
        Write-Host 'No hay contactos nuevos: los registros exactos ya existian.' -ForegroundColor Yellow
    }

    # El respaldo se crea ANTES de cualquier escritura.
    $backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'Respaldos'
    if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupPath = Join-Path $backupDir ("Unialiados_Operacion_Unificada_antes_carga_$stamp.xlsm")
    Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -ErrorAction Stop

    Append-Table $companiesTable $companyHeaders $newCompanies
    Append-Table $leadsTable $leadHeaders $newLeads

    # Una vez importado correctamente, vacia SOLO los campos editables de la carga.
    # Se mantienen los encabezados, formulas N2:N1001, listas desplegables y estilos.
    $lastStagingRow = [int]$used.Rows.Count
    if ($lastStagingRow -ge 2) {
        $stageRange = $loadSheet.Range('A2:M' + $lastStagingRow)
        $stageRange.ClearContents()
        Release-Com $stageRange
    }
    $book.Save()
    Write-Host "`nIMPORTACION TERMINADA" -ForegroundColor Green
    Write-Host "Empresas nuevas: $($newCompanies.Count)"
    Write-Host "Contactos nuevos: $($newLeads.Count)"
    Write-Host "Contactos identicos omitidos (empresa+persona+correo): $contactSkipped"
    Write-Host "Entradas relacionadas con empresa existente o repetida: $companySkipped"
    Write-Host "Respaldo: $backupPath"
    Write-Host 'La hoja CARGA MASIVA quedo vacia, lista para el siguiente lote.'
    Write-Host 'Se conservaron las macros y las columnas usadas por n8n.'
    Write-Host 'Los correos repetidos en empresas/contactos distintos fueron permitidos.'
} catch {
    Write-Host "`nIMPORTACION CANCELADA: $($_.Exception.Message)" -ForegroundColor Red
    if ($backupPath) { Write-Host "Respaldo conservado: $backupPath" }
    exit 1
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-Com $book; Release-Com $excel
    if ($mutex) { if ($lockTaken) { try { $mutex.ReleaseMutex() } catch {} }; $mutex.Dispose() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
