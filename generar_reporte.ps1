<#
Genera index.html (dashboard estatico) a partir del Excel de ventas.
Uso:
  powershell -ExecutionPolicy Bypass -File generar_reporte.ps1
Cada vez que actualices el archivo de Excel en esta carpeta, vuelve a correr este script
y luego sube (git add / commit / push) el index.html generado.
#>

param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot "2026.08.19_reporte de ventas.xlsx"),
  [string]$SheetName = "VENTAS 2026",
  [string]$OutputPath = (Join-Path $PSScriptRoot "index.html"),
  [datetime]$FechaCorte = (Get-Date),
  [string]$LogoPath = (Join-Path $PSScriptRoot "logo_vemolca.png")
)

$ErrorActionPreference = "Stop"
$culture = [System.Globalization.CultureInfo]::GetCultureInfo("es-ES")

function Fmt0($n) { return [math]::Round([double]$n, 0).ToString("N0", $culture) }
function Fmt1Pct($n) { return ([math]::Round([double]$n, 1)).ToString("N1", $culture) + "%" }
function Pct($part, $total) { if ($total -eq 0) { return 0 } else { return ($part / $total) * 100 } }
function FmtCell($n) {
  $v = [math]::Round([double]$n, 0)
  if ($v -eq 0) { return "<span class='zero'>&mdash;</span>" }
  if ($v -lt 0) { return "<span class='neg'>($([math]::Abs($v).ToString("N0", $culture)))</span>" }
  return $v.ToString("N0", $culture)
}
function Get-NiceMax($v) {
  if ($v -le 0) { return 100.0 }
  $exp = [math]::Floor([math]::Log10($v))
  $base = [math]::Pow(10, $exp)
  $frac = $v / $base
  $niceFrac = 10
  if ($frac -le 1) { $niceFrac = 1 } elseif ($frac -le 2) { $niceFrac = 2 } elseif ($frac -le 5) { $niceFrac = 5 }
  return $niceFrac * $base
}
function New-RankingChartSvg($items, $ariaLabel) {
  $chartW = 900
  $rowH = 30
  $marginL = 96
  $marginR = 64
  $marginT = 10
  $marginB = 26
  $rows = [math]::Max($items.Count, 1)
  $plotH = $rows * $rowH
  $chartH = $marginT + $plotH + $marginB
  $plotW = $chartW - $marginL - $marginR
  $niceMax = Get-NiceMax (($items | Measure-Object Value -Maximum).Maximum)
  $barThick = 20

  $barsSvg = ""
  $i = 0
  foreach ($it in $items) {
    $rowY = $marginT + ($i * $rowH)
    $barCy = $rowY + ($rowH / 2)
    $w = if ($niceMax -gt 0) { $plotW * ($it.Value / $niceMax) } else { 0 }
    if ($w -lt 1 -and $it.Value -gt 0) { $w = 1 }
    $barY = $barCy - ($barThick / 2)
    $isTop = $i -eq 0
    $barClass = if ($isTop) { "bar-rect peak" } else { "bar-rect" }
    $labelSvg = ""
    if ($isTop) {
      $labelSvg = "<text x='$([math]::Round($marginL + $w + 8,1))' y='$([math]::Round($barCy + 4,1))' text-anchor='start' class='chart-toplabel'>$(Fmt0 $it.Value)</text>"
    }
    $docsAttr = if ($it.PSObject.Properties.Name -contains 'Docs') { " data-docs='$($it.Docs)'" } else { "" }
    $barsSvg += "<rect class='bar-hit' tabindex='0' x='0' y='$([math]::Round($rowY,1))' width='$chartW' height='$rowH' data-label='$($it.Label)' data-value='$(Fmt0 $it.Value)'$docsAttr></rect><rect class='$barClass' x='$marginL' y='$([math]::Round($barY,1))' width='$([math]::Round($w,1))' height='$barThick' rx='4'></rect><text x='$($marginL - 8)' y='$([math]::Round($barCy + 4,1))' text-anchor='end' class='chart-xlabel mono'>$($it.Label)</text>$labelSvg`n"
    $i++
  }

  $gridSvg = ""
  for ($t = 0; $t -le 4; $t++) {
    $val = $niceMax * $t / 4
    $gx = $marginL + ($plotW * $t / 4)
    $gridSvg += "<line x1='$([math]::Round($gx,1))' y1='$marginT' x2='$([math]::Round($gx,1))' y2='$([math]::Round($marginT + $plotH,1))' class='chart-grid'></line><text x='$([math]::Round($gx,1))' y='$([math]::Round($marginT + $plotH + 16,1))' text-anchor='middle' class='chart-ylabel'>$(Fmt0 $val)</text>`n"
  }

  return @"
<div class="chart-card">
<svg viewBox="0 0 $chartW $chartH" class="chart-svg" role="img" aria-label="$ariaLabel">
$gridSvg
$barsSvg
</svg>
<div class="chart-tooltip"></div>
</div>
"@
}
function FmtUSD($n) { return "`$ " + ([double]$n).ToString("N2", $culture) }
function Parse-MoneyText($s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return 0.0 }
  $clean = $s -replace '[^\d,.\-]', ''
  if ($clean -eq '' -or $clean -eq '-') { return 0.0 }
  $clean = $clean -replace '\.', ''
  $clean = $clean -replace ',', '.'
  $val = 0.0
  [double]::TryParse($clean, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val) | Out-Null
  return $val
}
function Badge($estado) {
  if ($estado -eq "Facturado y cobrado") { return "<span class='badge badge-green'>Cobrado</span>" }
  if ($estado -eq "Facturado sin cobrar") { return "<span class='badge badge-amber'>Por cobrar</span>" }
  if ($estado -eq "Sin facturar") { return "<span class='badge badge-red'>Sin facturar</span>" }
  return $estado
}

$MESES = @("Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic")

# Entidades HTML para evitar problemas de codificacion con tildes en este script
$e_a = "&aacute;"; $e_e = "&eacute;"; $e_i = "&iacute;"; $e_o = "&oacute;"; $e_u = "&uacute;"
$e_n = "&ntilde;"; $e_ud = "&uuml;"

$logoImgTag = ""
if (Test-Path $LogoPath) {
  $logoBytes = [System.IO.File]::ReadAllBytes($LogoPath)
  $logoB64 = [Convert]::ToBase64String($logoBytes)
  $ext = [System.IO.Path]::GetExtension($LogoPath).TrimStart(".").ToLower()
  if ($ext -eq "jpg") { $ext = "jpeg" }
  $logoImgTag = "<img src=`"data:image/$ext;base64,$logoB64`" alt=`"Vemolca`" class=`"logo`">"
}

if (-not (Test-Path $ExcelPath)) {
  throw "No se encontro el archivo de Excel: $ExcelPath"
}

Write-Host "Leyendo $ExcelPath ..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($ExcelPath, $null, $true)
$ws = $wb.Worksheets.Item($SheetName)

# ---------- Ubicar columnas por nombre de encabezado (fila 1) ----------
# La hoja ha cambiado de estructura antes (p.ej. se agrego "TIPO DE CLIENTE" y corrio
# columnas a la derecha). En vez de indices fijos, buscamos cada columna por su
# encabezado para que el script no falle silenciosamente si vuelve a cambiar.
$headerMap = @{}
for ($c = 1; $c -le 30; $c++) {
  $h = ($ws.Cells.Item(1, $c).Text).Trim().ToUpper()
  if ($h -ne "" -and -not $headerMap.ContainsKey($h)) { $headerMap[$h] = $c }
}
function Get-ColIndex($names, $required) {
  foreach ($n in $names) {
    if ($headerMap.ContainsKey($n)) { return $headerMap[$n] }
  }
  if ($required) { throw "No se encontro en la hoja '$SheetName' la columna esperada ($($names -join ' / ')). Revisa si cambiaron los encabezados de la fila 1." }
  return $null
}

$colTotalFac     = Get-ColIndex @("TOTAL FAC") $true
$colBase         = Get-ColIndex @("BASE IMPONIBLE") $true
$colExento       = Get-ColIndex @("EXENTO") $true
$colIVA          = Get-ColIndex @("IVA") $true
$colIGTF         = Get-ColIndex @("IGTF") $true
$colPO           = Get-ColIndex @("PO") $true
$colCliente      = Get-ColIndex @("CLIENTE") $true
$colTipoCliente  = Get-ColIndex @("TIPO DE CLIENTE") $false
$colDescripcion  = Get-ColIndex @("DESCRIPCION") $true
$colFechaPO      = Get-ColIndex @("FECHA. PO", "FECHA PO") $true
$colFacturado    = Get-ColIndex @("FACTURADO") $true
$colCobranza     = Get-ColIndex @("COBRANZA") $true
$colNroFactura   = Get-ColIndex @("Nº. FACTURA", "N°. FACTURA", "NO. FACTURA") $false
if (-not $colNroFactura) { $colNroFactura = 2 }
$colFechaFactura = Get-ColIndex @("FECHA. FACTURA", "FECHA FACTURA") $false
if (-not $colFechaFactura) { $colFechaFactura = 3 }

$raw = New-Object System.Collections.Generic.List[object]
$firstRow = 2
$r = $firstRow
while ($true) {
  $totalText = $ws.Cells.Item($r, $colTotalFac).Text
  if ([string]::IsNullOrWhiteSpace($totalText)) { break }

  $facturado = ($ws.Cells.Item($r, $colFacturado).Text).Trim().ToUpper()
  $cobranza  = ($ws.Cells.Item($r, $colCobranza).Text).Trim().ToUpper()
  $cliente   = ($ws.Cells.Item($r, $colCliente).Text).Trim().ToUpper()
  $tipoCliente = if ($colTipoCliente) { ($ws.Cells.Item($r, $colTipoCliente).Text).Trim().ToUpper() } else { "" }
  $nroFactura = ($ws.Cells.Item($r, $colNroFactura).Text).Trim()
  $descripcion = ($ws.Cells.Item($r, $colDescripcion).Text).Trim()

  $fechaFacTxt = ($ws.Cells.Item($r, $colFechaFactura).Text).Trim()
  $fechaPOTxt  = ($ws.Cells.Item($r, $colFechaPO).Text).Trim()
  $fechaFac = $null
  if ($fechaFacTxt -ne "") { try { $fechaFac = [datetime]::ParseExact($fechaFacTxt, "dd/MM/yyyy", $culture) } catch {} }
  $fechaPO = $null
  if ($fechaPOTxt -ne "") { try { $fechaPO = [datetime]::ParseExact($fechaPOTxt, "dd/MM/yyyy", $culture) } catch {} }
  $mesRef = if ($fechaFac) { $fechaFac } elseif ($fechaPO) { $fechaPO } else { $null }

  # Clasificacion provisional por columna de texto (se puede sobreescribir mas abajo
  # con la clasificacion oficial de la hoja, si se logra leer)
  $estado = "Sin facturar"
  if ($facturado -eq "FACTURADO" -and $cobranza -eq "COBRADO") { $estado = "Facturado y cobrado" }
  elseif ($facturado -eq "FACTURADO" -and $cobranza -eq "NO COBRADO") { $estado = "Facturado sin cobrar" }

  $raw.Add([PSCustomObject]@{
    WsRow         = $r
    NroFactura    = $nroFactura
    FechaFactura  = $fechaFac
    MesRef        = $mesRef
    TotalFac      = [double]$ws.Cells.Item($r, $colTotalFac).Value2
    BaseImponible = [double]$ws.Cells.Item($r, $colBase).Value2
    Exento        = [double]$ws.Cells.Item($r, $colExento).Value2
    IVA           = [double]$ws.Cells.Item($r, $colIVA).Value2
    IGTF          = [double]$ws.Cells.Item($r, $colIGTF).Value2
    PO            = ($ws.Cells.Item($r, $colPO).Text).Trim()
    Cliente       = $cliente
    TipoCliente   = $tipoCliente
    Descripcion   = $descripcion
    Facturado     = $facturado
    Cobranza      = $cobranza
    Estado        = $estado
  }) | Out-Null

  $r++
}
$lastRow = $r - 1

# ---------- Clasificacion oficial de la hoja (si existe) ----------
# Debajo de la tabla, la hoja suele traer 3 filas de resumen ya validadas por el
# usuario ("FACTURADO Y COBRADO", "FACTURADO Y NO COBRADO", "SIN FACTURAR NI COBRAR"),
# escritas como formulas que suman celdas puntuales de la columna D (TOTAL FAC).
# Si el texto de las columnas FACTURADO/COBRANZA no coincide con esa formula para
# alguna fila (por ejemplo, un dato de cobranza no actualizado), usamos la formula
# oficial como fuente de verdad, ya que es la que el usuario ya reconcilia en Excel.
function Get-RowRefs($formula) {
  $rows = New-Object System.Collections.Generic.HashSet[int]
  if ([string]::IsNullOrWhiteSpace($formula)) { return $rows }
  foreach ($m in [regex]::Matches($formula, '[A-Z]+(\d+)(?::[A-Z]+(\d+))?')) {
    $start = [int]$m.Groups[1].Value
    if ($m.Groups[2].Success) {
      $end = [int]$m.Groups[2].Value
      for ($i = $start; $i -le $end; $i++) { $rows.Add($i) | Out-Null }
    } else {
      $rows.Add($start) | Out-Null
    }
  }
  return $rows
}

$labelCol = 3
$valueCol = 4
$etiquetas = @{}
for ($rr = $lastRow + 1; $rr -le ($lastRow + 20); $rr++) {
  $lbl = ($ws.Cells.Item($rr, $labelCol).Text).Trim().ToUpper()
  if ($lbl -ne "") { $etiquetas[$lbl] = $rr }
}

$rowsCobrado = New-Object System.Collections.Generic.HashSet[int]
$rowsNoCobrado = New-Object System.Collections.Generic.HashSet[int]
$rowsSinFacturar = New-Object System.Collections.Generic.HashSet[int]
$clasificacionOficial = $false

$tieneNoCobrado = $etiquetas.ContainsKey("FACTURADO Y NO COBRADO:") -or $etiquetas.ContainsKey("FACTURADO Y NO COBRADO")
if ($etiquetas.ContainsKey("FACTURADO Y COBRADO") -and $tieneNoCobrado) {
  $keyCobrado = "FACTURADO Y COBRADO"
  $keyNoCobrado = if ($etiquetas.ContainsKey("FACTURADO Y NO COBRADO:")) { "FACTURADO Y NO COBRADO:" } else { "FACTURADO Y NO COBRADO" }
  $keySinFacturar = $etiquetas.Keys | Where-Object { $_ -like "SIN FACTURAR*" } | Select-Object -First 1

  if ($etiquetas.ContainsKey($keyCobrado)) {
    $rowsCobrado = Get-RowRefs ($ws.Cells.Item($etiquetas[$keyCobrado], $valueCol).Formula)
  }
  if ($etiquetas.ContainsKey($keyNoCobrado)) {
    $rowsNoCobrado = Get-RowRefs ($ws.Cells.Item($etiquetas[$keyNoCobrado], $valueCol).Formula)
  }
  if ($keySinFacturar) {
    $rowsSinFacturar = Get-RowRefs ($ws.Cells.Item($etiquetas[$keySinFacturar], $valueCol).Formula)
  }

  $totalClasificado = $rowsCobrado.Count + $rowsNoCobrado.Count + $rowsSinFacturar.Count
  if ($totalClasificado -eq $raw.Count) {
    $clasificacionOficial = $true
    foreach ($item in $raw) {
      if ($rowsCobrado.Contains($item.WsRow)) { $item.Estado = "Facturado y cobrado" }
      elseif ($rowsNoCobrado.Contains($item.WsRow)) { $item.Estado = "Facturado sin cobrar" }
      elseif ($rowsSinFacturar.Contains($item.WsRow)) { $item.Estado = "Sin facturar" }
    }
  }
}

# ---------- Hoja "CC INPROCCA": estado de cuenta detallado de ese cliente ----------
$inproccaRows = New-Object System.Collections.Generic.List[object]
$inproccaHeaders = @("Abonos", "Cr${e_e}dito", "Abonos", "Cr${e_e}dito", "Abonos", "Cr${e_e}dito")
$hojaInprocca = $null
foreach ($hoja in $wb.Worksheets) { if ($hoja.Name -eq "CC INPROCCA") { $hojaInprocca = $hoja } }
if ($hojaInprocca) {
  $wsi = $hojaInprocca
  $rowsI = $wsi.UsedRange.Rows.Count
  $palabraMala = "c" + [char]0xE9 + "dito"
  $palabraBuena = "cr" + [char]0xE9 + "dito"
  for ($hc = 0; $hc -lt 6; $hc++) {
    $txt = ($wsi.Cells.Item(4, 5 + $hc).Text).Trim()
    # Correccion tipografica: el Excel a veces trae "cedito" en vez de "credito"
    $txt = $txt -replace [regex]::Escape($palabraMala), $palabraBuena
    if ($txt -ne "") { $inproccaHeaders[$hc] = $txt }
  }
  for ($ri = 5; $ri -le $rowsI; $ri++) {
    $fecha = ($wsi.Cells.Item($ri, 2).Text).Trim()
    $doc   = ($wsi.Cells.Item($ri, 3).Text).Trim()
    $monto = ($wsi.Cells.Item($ri, 4).Text).Trim()
    $ab1   = ($wsi.Cells.Item($ri, 5).Text).Trim()
    $cr1   = ($wsi.Cells.Item($ri, 6).Text).Trim()
    $ab2   = ($wsi.Cells.Item($ri, 7).Text).Trim()
    $cr2   = ($wsi.Cells.Item($ri, 8).Text).Trim()
    $ab3   = ($wsi.Cells.Item($ri, 9).Text).Trim()
    $cr3   = ($wsi.Cells.Item($ri, 10).Text).Trim()
    if ($fecha -eq "" -and $doc -eq "" -and $monto -eq "") { continue }
    $inproccaRows.Add([PSCustomObject]@{
      Fecha = $fecha; Doc = $doc; Monto = $monto
      Ab1 = $ab1; Cr1 = $cr1; Ab2 = $ab2; Cr2 = $cr2; Ab3 = $ab3; Cr3 = $cr3
    }) | Out-Null
  }
}

# ---------- Hoja "DISPONIBILIDAD": efectivo/bancos disponibles ----------
$dispoRows = New-Object System.Collections.Generic.List[object]
$dispoTitulo = ""
$dispoTasaCOP = ""
$dispoTasaUSDT = ""
$hojaDispo = $null
foreach ($hoja in $wb.Worksheets) { if ($hoja.Name -eq "DISPONIBILIDAD") { $hojaDispo = $hoja } }
if ($hojaDispo) {
  $wsd = $hojaDispo
  $rowsD = $wsd.UsedRange.Rows.Count
  $dispoTasaCOP = ($wsd.Cells.Item(2, 3).Text).Trim()
  $dispoTasaUSDT = ($wsd.Cells.Item(3, 3).Text).Trim()
  $dispoTitulo = ($wsd.Cells.Item(5, 1).Text).Trim()
  for ($rd = 7; $rd -le $rowsD; $rd++) {
    $cuenta = ($wsd.Cells.Item($rd, 1).Text).Trim()
    $origen = ($wsd.Cells.Item($rd, 2).Text).Trim()
    $usd    = ($wsd.Cells.Item($rd, 3).Text).Trim()
    if ($cuenta -eq "" -and $origen -eq "" -and $usd -eq "") { continue }
    $dispoRows.Add([PSCustomObject]@{ Cuenta = $cuenta; Origen = $origen; UsdTexto = $usd; UsdValor = [double]$wsd.Cells.Item($rd, 3).Value2 }) | Out-Null
  }
}

# ---------- Hoja "cuentas por pagar": proveedores y provisiones ----------
$cxpProveedores = New-Object System.Collections.Generic.List[object]
$cxpProvisiones = New-Object System.Collections.Generic.List[object]
$cxpSubtotalProveedores = 0.0
$cxpTotalProvisiones = 0.0
$cxpProvisionesCuadra = $true
$hojaCxP = $null
foreach ($hoja in $wb.Worksheets) { if ($hoja.Name.Trim().ToLower() -eq "cuentas por pagar") { $hojaCxP = $hoja } }
if ($hojaCxP) {
  $wsp = $hojaCxP
  $rowsP = $wsp.UsedRange.Rows.Count
  $rp = 2
  # Bloque de proveedores: hasta la fila "SUB-TOTAL"
  while ($rp -le $rowsP) {
    $etiqueta = ($wsp.Cells.Item($rp, 2).Text).Trim()
    if ($etiqueta -eq "") { $rp++; continue }
    if ($etiqueta -match "(?i)sub-?total") {
      $cxpSubtotalProveedores = [double]($wsp.Cells.Item($rp, 3).Value2)
      $rp++
      break
    }
    $monto = $wsp.Cells.Item($rp, 3).Value2
    $cxpProveedores.Add([PSCustomObject]@{ Proveedor = $etiqueta; Monto = [double]($monto) }) | Out-Null
    $rp++
  }
  # Bloque de provisiones: salta el encabezado de seccion y lee items hasta el total
  # (el total puede venir como fila sin etiqueta con solo el monto, o como fila etiquetada "TOTAL ..."/"TOTAL GENERAL")
  $cxpTotalProvisionesExcel = $null
  $cxpTotalGeneralExcel = $null
  while ($rp -le $rowsP) {
    $etiqueta = ($wsp.Cells.Item($rp, 2).Text).Trim()
    $montoRaw = $wsp.Cells.Item($rp, 3).Value2
    $tieneMonto = ($montoRaw -ne $null -and $montoRaw -ne "")
    if (-not $tieneMonto) { $rp++; continue }
    if ($etiqueta -match "(?i)total\s*general") {
      $cxpTotalGeneralExcel = [double]$montoRaw
      $rp++
      continue
    }
    if ($etiqueta -eq "" -or $etiqueta -match "(?i)total") {
      $cxpTotalProvisionesExcel = [double]$montoRaw
      $rp++
      continue
    }
    $cxpProvisiones.Add([PSCustomObject]@{ Concepto = $etiqueta; Monto = [double]$montoRaw }) | Out-Null
    $rp++
  }
  $cxpTotalProvisiones = ($cxpProvisiones | Measure-Object Monto -Sum).Sum
  if ($cxpTotalProvisionesExcel -ne $null -and [math]::Abs($cxpTotalProvisionesExcel - $cxpTotalProvisiones) -gt 0.5) {
    $cxpProvisionesCuadra = $false
  }
}
$cxpTotalGeneral = $cxpSubtotalProveedores + $cxpTotalProvisiones

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Filas leidas: $($raw.Count)"
Write-Host "Filas leidas de CC INPROCCA: $($inproccaRows.Count)"
Write-Host "Filas leidas de DISPONIBILIDAD: $($dispoRows.Count)"
Write-Host "Filas leidas de CUENTAS POR PAGAR: $($cxpProveedores.Count) proveedores, $($cxpProvisiones.Count) provisiones"
if (-not $cxpProvisionesCuadra) {
  Write-Host "Aviso: el total de provisiones de 'cuentas por pagar' en el Excel no cuadra con la suma de sus items; se uso la suma de los items."
}
if ($clasificacionOficial) {
  Write-Host "Clasificacion de cobranza: usando el resumen oficial de la hoja (coincide con el cuadre de Excel)."
} else {
  Write-Host "Clasificacion de cobranza: no se encontro/no cuadro el resumen oficial de la hoja; se uso la columna COBRANZA fila por fila."
}

# ---------- Totales generales ----------
$totalVentas   = ($raw | Measure-Object TotalFac -Sum).Sum
$totalBase     = ($raw | Measure-Object BaseImponible -Sum).Sum
$totalExento   = ($raw | Measure-Object Exento -Sum).Sum
$totalBaseExe  = $totalBase + $totalExento
$totalIVA      = ($raw | Measure-Object IVA -Sum).Sum
$totalIGTF     = ($raw | Measure-Object IGTF -Sum).Sum
$totalImpuestos = $totalIVA + $totalIGTF
$docsTotal     = $raw.Count
$ticketProm    = if ($docsTotal -gt 0) { $totalVentas / $docsTotal } else { 0 }

$grpCobrado    = $raw | Where-Object { $_.Estado -eq "Facturado y cobrado" }
$grpPorCobrar  = $raw | Where-Object { $_.Estado -eq "Facturado sin cobrar" }
$grpSinFactura = $raw | Where-Object { $_.Estado -eq "Sin facturar" }

$montoCobrado   = ($grpCobrado   | Measure-Object TotalFac -Sum).Sum
$montoPorCobrar = ($grpPorCobrar | Measure-Object TotalFac -Sum).Sum
$montoSinFactura = ($grpSinFactura | Measure-Object TotalFac -Sum).Sum
$docsFacturados = $grpCobrado.Count + $grpPorCobrar.Count
$montoFacturado = $montoCobrado + $montoPorCobrar

# ---------- Evolucion mensual ----------
$mensual = $raw | Where-Object { $_.MesRef -ne $null } |
  Group-Object { $_.MesRef.ToString("yyyy-MM") } |
  Sort-Object Name |
  ForEach-Object {
    $suma = ($_.Group | Measure-Object TotalFac -Sum).Sum
    $ym = $_.Name.Split("-")
    [PSCustomObject]@{
      Label = "$($MESES[[int]$ym[1]-1]) $($ym[0])"
      Docs  = $_.Count
      Monto = $suma
    }
  }
$maxMensual = ($mensual | Measure-Object Monto -Maximum).Maximum

$filasMensual = ($mensual | ForEach-Object {
  $pct = Pct $_.Monto $totalVentas
  "<tr><td class=mono>$($_.Label)</td><td class=n>$($_.Docs)</td><td class=n>$(FmtCell $_.Monto)</td><td class=n>$(Fmt1Pct $pct)</td></tr>"
}) -join "`n"

# ---------- Grafico de barras: ventas por mes ----------
$chartW = 900
$chartH = 260
$marginL = 60
$marginR = 12
$marginT = 30
$marginB = 34
$plotW = $chartW - $marginL - $marginR
$plotH = $chartH - $marginT - $marginB
$niceMax = Get-NiceMax $maxMensual
$mesesCount = [math]::Max($mensual.Count, 1)
$slotW = $plotW / $mesesCount
$barW = [math]::Min(28, $slotW * 0.55)
$mesPico = $mensual | Sort-Object Monto -Descending | Select-Object -First 1

$barsSvg = ""
$i = 0
foreach ($m in $mensual) {
  $cx = $marginL + ($i + 0.5) * $slotW
  $x = $cx - ($barW / 2)
  $h = if ($niceMax -gt 0) { $plotH * ($m.Monto / $niceMax) } else { 0 }
  if ($h -lt 1 -and $m.Monto -gt 0) { $h = 1 }
  $y = $marginT + $plotH - $h
  $hitX = $marginL + ($i * $slotW)
  $isPico = ($mesPico -ne $null) -and ($m.Label -eq $mesPico.Label)
  $barClass = if ($isPico) { "bar-rect peak" } else { "bar-rect" }
  $labelSvg = ""
  if ($isPico) {
    $labelSvg = "<text x='$([math]::Round($cx,1))' y='$([math]::Round($y - 8,1))' text-anchor='middle' class='chart-toplabel'>$(Fmt0 $m.Monto)</text>"
  }
  $mesCorto = $m.Label.Split(' ')[0]
  $barsSvg += "<rect class='bar-hit' tabindex='0' x='$([math]::Round($hitX,1))' y='$marginT' width='$([math]::Round($slotW,1))' height='$plotH' data-label='$($m.Label)' data-value='$(Fmt0 $m.Monto)' data-docs='$($m.Docs)'></rect><rect class='$barClass' x='$([math]::Round($x,1))' y='$([math]::Round($y,1))' width='$([math]::Round($barW,1))' height='$([math]::Round($h,1))' rx='4'></rect>$labelSvg<text x='$([math]::Round($cx,1))' y='$($chartH - 12)' text-anchor='middle' class='chart-xlabel'>$mesCorto</text>`n"
  $i++
}

$gridSvg = ""
for ($t = 0; $t -le 4; $t++) {
  $val = $niceMax * $t / 4
  $gy = $marginT + $plotH - ($plotH * $t / 4)
  $gridSvg += "<line x1='$marginL' y1='$([math]::Round($gy,1))' x2='$($chartW - $marginR)' y2='$([math]::Round($gy,1))' class='chart-grid'></line><text x='$($marginL - 8)' y='$([math]::Round($gy + 3,1))' text-anchor='end' class='chart-ylabel'>$(Fmt0 $val)</text>`n"
}

$monthlyChartSvg = @"
<div class="chart-card">
<svg viewBox="0 0 $chartW $chartH" class="chart-svg" role="img" aria-label="Ventas mensuales en USD">
$gridSvg
$barsSvg
</svg>
<div class="chart-tooltip"></div>
</div>
"@

# ---------- Analisis por cliente ----------
$porCliente = $raw | Group-Object Cliente | ForEach-Object {
  $g = $_.Group
  $venta = ($g | Measure-Object TotalFac -Sum).Sum
  $tipo = ($g | Where-Object { $_.TipoCliente -ne "" } | Select-Object -First 1 -ExpandProperty TipoCliente)
  if (-not $tipo) { $tipo = "" }
  [PSCustomObject]@{
    Cliente     = $_.Name
    TipoCliente = $tipo
    Docs        = $_.Count
    VentaTotal  = $venta
    Cobrado     = ($g | Where-Object { $_.Estado -eq "Facturado y cobrado" } | Measure-Object TotalFac -Sum).Sum
    PorCobrar   = ($g | Where-Object { $_.Estado -eq "Facturado sin cobrar" } | Measure-Object TotalFac -Sum).Sum
    SinFacturar = ($g | Where-Object { $_.Estado -eq "Sin facturar" } | Measure-Object TotalFac -Sum).Sum
  }
} | Sort-Object VentaTotal -Descending

$filasClienteVentas = ($porCliente | ForEach-Object {
  $pct = Pct $_.VentaTotal $totalVentas
  "<tr><td><b>$($_.Cliente)</b></td><td class=n>$($_.Docs)</td><td class=n>$(FmtCell $_.VentaTotal)</td><td class=n>$(Fmt1Pct $pct)</td></tr>"
}) -join "`n"

$carteraCliente = $porCliente | ForEach-Object {
  $pend = $_.PorCobrar + $_.SinFacturar
  [PSCustomObject]@{
    Cliente = $_.Cliente
    TipoCliente = $_.TipoCliente
    Cobrado = $_.Cobrado
    PorCobrar = $_.PorCobrar
    SinFacturar = $_.SinFacturar
    Pendiente = $pend
  }
} | Sort-Object Pendiente -Descending

function New-FilasCarteraYTotal($items) {
  $filas = ($items | ForEach-Object {
    "<tr><td><b>$($_.Cliente)</b></td><td class=n>$(FmtCell $_.Cobrado)</td><td class=n>$(FmtCell $_.PorCobrar)</td><td class=n>$(FmtCell $_.SinFacturar)</td><td class=n>$(FmtCell $_.Pendiente)</td></tr>"
  }) -join "`n"
  $totCobrado = ($items | Measure-Object Cobrado -Sum).Sum
  $totPorCobrar = ($items | Measure-Object PorCobrar -Sum).Sum
  $totSinFacturar = ($items | Measure-Object SinFacturar -Sum).Sum
  $totPendiente = ($items | Measure-Object Pendiente -Sum).Sum
  $tfoot = "<tr><td>TOTAL</td><td class=n>$(FmtCell $totCobrado)</td><td class=n>$(FmtCell $totPorCobrar)</td><td class=n>$(FmtCell $totSinFacturar)</td><td class=n>$(FmtCell $totPendiente)</td></tr>"
  return @{ Filas = $filas; Tfoot = $tfoot }
}

$carteraONG = $carteraCliente | Where-Object { $_.TipoCliente -eq "ONG" }
$txtPetroleo = "PETR" + [char]0xD3 + "LEO"
$carteraPetroleo = $carteraCliente | Where-Object { $_.TipoCliente -eq $txtPetroleo -or $_.TipoCliente -eq "PETROLEO" }
$carteraOtros = $carteraCliente | Where-Object { $_.TipoCliente -ne "ONG" -and $_.TipoCliente -ne $txtPetroleo -and $_.TipoCliente -ne "PETROLEO" }

$carteraONGResult = New-FilasCarteraYTotal $carteraONG
$carteraPetroleoResult = New-FilasCarteraYTotal $carteraPetroleo
$carteraOtrosResult = if ($carteraOtros.Count -gt 0) { New-FilasCarteraYTotal $carteraOtros } else { $null }

$carteraOtrosPanelHtml = ""
if ($carteraOtrosResult) {
  $carteraOtrosPanelHtml = @"
<div class="panel-head"><div class="eyebrow">Cobranza $([char]0xB7) Otros</div><h2>Cartera pendiente $([char]0x2014) Otros</h2><p class="panel-desc">Clientes sin tipo asignado en la columna "TIPO DE CLIENTE" del Excel.</p></div>
<div class="table-scroll">
<table><thead><tr><th>Cliente</th><th class=n>Cobrado</th><th class=n>Por cobrar</th><th class=n>Sin facturar</th><th class=n>Total pendiente</th></tr></thead>
<tbody>
$($carteraOtrosResult.Filas)
</tbody>
<tfoot>$($carteraOtrosResult.Tfoot)</tfoot></table>
</div>
"@
}

# ---------- Grafico: ranking de clientes por venta ----------
$rankItems = $porCliente | ForEach-Object { [PSCustomObject]@{ Label = $_.Cliente; Value = $_.VentaTotal; Docs = $_.Docs } }
$rankChartSvg = New-RankingChartSvg $rankItems "Ranking de clientes por venta"

# ---------- Ventas no facturadas, desglosadas por cliente ----------
$sinFacturaPorCliente = $grpSinFactura | Group-Object Cliente | ForEach-Object {
  [PSCustomObject]@{
    Cliente = $_.Name
    Docs    = $_.Count
    Monto   = ($_.Group | Measure-Object TotalFac -Sum).Sum
  }
} | Sort-Object Monto -Descending

$filasSinFacturaCliente = ($sinFacturaPorCliente | ForEach-Object {
  $pct = Pct $_.Monto $montoSinFactura
  "<tr><td><b>$($_.Cliente)</b></td><td class=n>$($_.Docs)</td><td class=n>$(FmtCell $_.Monto)</td><td class=n>$(Fmt1Pct $pct)</td></tr>"
}) -join "`n"

$sinFacturaChartItems = $sinFacturaPorCliente | ForEach-Object { [PSCustomObject]@{ Label = $_.Cliente; Value = $_.Monto; Docs = $_.Docs } }
$sinFacturaChartSvg = New-RankingChartSvg $sinFacturaChartItems "Ranking de ventas sin facturar por cliente"

$filasSinFacturaDetalle = ($grpSinFactura | Sort-Object @{Expression = { $_.MesRef } } | ForEach-Object {
  $fechaTxt = if ($_.MesRef) { $_.MesRef.ToString("dd/MM/yyyy") } else { "" }
  "<tr><td><b>$($_.Cliente)</b></td><td class=mono>$($_.PO)</td><td>$($_.Descripcion)</td><td class=mono>$fechaTxt</td><td class=n>$(FmtCell $_.TotalFac)</td></tr>"
}) -join "`n"

$top3Pct = Pct (($porCliente | Select-Object -First 3 | Measure-Object VentaTotal -Sum).Sum) $totalVentas
$clienteTop = $porCliente | Select-Object -First 1
$clienteTopPct = Pct $clienteTop.VentaTotal $totalVentas

# ---------- Antiguedad de cuentas por cobrar ----------
$tramos = [ordered]@{ "0-30 dias" = 0.0; "31-60 dias" = 0.0; "61-90 dias" = 0.0; "+90 dias" = 0.0 }
foreach ($row in $grpPorCobrar) {
  if ($row.FechaFactura) {
    $dias = ($FechaCorte - $row.FechaFactura).Days
  } else {
    $dias = 0
  }
  if ($dias -le 30) { $tramos["0-30 dias"] += $row.TotalFac }
  elseif ($dias -le 60) { $tramos["31-60 dias"] += $row.TotalFac }
  elseif ($dias -le 90) { $tramos["61-90 dias"] += $row.TotalFac }
  else { $tramos["+90 dias"] += $row.TotalFac }
}
$filasAging = ($tramos.GetEnumerator() | ForEach-Object {
  $pct = Pct $_.Value $montoPorCobrar
  "<tr><td>$($_.Key)</td><td class=n>$(FmtCell $_.Value)</td><td class=n>$(Fmt1Pct $pct)</td></tr>"
}) -join "`n"

# ---------- Estado de cuenta INPROCCA (hoja "CC INPROCCA") ----------
$fechasOcultarInprocca = @("07/05/2026", "14/07/2026")
$inproccaVisibles = $inproccaRows | Where-Object { $fechasOcultarInprocca -notcontains $_.Fecha }
# Ajuste manual solicitado: en la fila "conciliado", el Abono y el Saldo de credito
# al 07/05/2026 se muestran en $0 en el dashboard (el Excel fuente no se modifica).
foreach ($row in $inproccaVisibles) {
  if ($row.Doc -match "(?i)conciliado") {
    $row.Ab1 = "`$ -"
    $row.Cr1 = "`$ -"
  }
}
$filasInprocca = ($inproccaVisibles | ForEach-Object {
  $esResumen = $_.Doc -match "(?i)deuda|saldo|conciliado"
  $docHtml = if ($esResumen) { "<b>$($_.Doc)</b>" } else { $_.Doc }
  $rowClass = if ($esResumen) { " class='resumen'" } else { "" }
  "<tr$rowClass><td class=ctr>$($_.Fecha)</td><td>$docHtml</td><td class=n>$($_.Monto)</td><td class=ctr>$($_.Ab1)</td><td class=n>$($_.Cr1)</td><td class=ctr>$($_.Ab2)</td><td class=n>$($_.Cr2)</td><td class=ctr>$($_.Ab3)</td><td class=n>$($_.Cr3)</td></tr>"
}) -join "`n"

$inproccaRealRows = $inproccaVisibles | Where-Object { $_.Doc -notmatch "(?i)deuda|saldo|conciliado" }
$inproccaTotMonto = ($inproccaRealRows | ForEach-Object { Parse-MoneyText $_.Monto } | Measure-Object -Sum).Sum
$inproccaTotAb1   = ($inproccaRealRows | ForEach-Object { Parse-MoneyText $_.Ab1 } | Measure-Object -Sum).Sum
$inproccaTotCr1   = ($inproccaRealRows | ForEach-Object { Parse-MoneyText $_.Cr1 } | Measure-Object -Sum).Sum
$inproccaTotAb2   = ($inproccaRealRows | ForEach-Object { Parse-MoneyText $_.Ab2 } | Measure-Object -Sum).Sum
$inproccaTotCr2   = ($inproccaRealRows | ForEach-Object { Parse-MoneyText $_.Cr2 } | Measure-Object -Sum).Sum
$inproccaTotAb3   = ($inproccaRealRows | ForEach-Object { Parse-MoneyText $_.Ab3 } | Measure-Object -Sum).Sum
$inproccaTotCr3   = ($inproccaRealRows | ForEach-Object { Parse-MoneyText $_.Cr3 } | Measure-Object -Sum).Sum

# ---------- Disponibilidad (hoja "DISPONIBILIDAD") ----------
$totalDispoUSD = ($dispoRows | Measure-Object UsdValor -Sum).Sum
$filasDispo = ($dispoRows | ForEach-Object {
  "<tr><td><b>$($_.Cuenta)</b></td><td class=n>$($_.Origen)</td><td class=n>$($_.UsdTexto)</td></tr>"
}) -join "`n"

# ---------- Mes de mayor actividad ----------
$mesTop = $mensual | Sort-Object Monto -Descending | Select-Object -First 1
$mesTopPct = if ($mesTop) { Pct $mesTop.Monto $totalVentas } else { 0 }

# ---------- Top 10 operaciones ----------
$top10 = $raw | Sort-Object TotalFac -Descending | Select-Object -First 10
$filasTop10 = ($top10 | ForEach-Object {
  $doc = if ($_.NroFactura -ne "") { "<span class=mono>$($_.NroFactura)</span>" } else { "<i>por facturar</i>" }
  "<tr><td>$doc</td><td>$($_.Cliente)</td><td>$($_.Descripcion)</td><td class=n>$(FmtCell $_.TotalFac)</td><td>$(Badge $_.Estado)</td></tr>"
}) -join "`n"

# ---------- Conversion a caja ----------
$pctCobrado = Pct $montoCobrado $totalVentas
$pctPorCobrar = Pct $montoPorCobrar $totalVentas
$pctSinFactura = Pct $montoSinFactura $totalVentas
$pctPendienteCaja = $pctPorCobrar + $pctSinFactura
$montoPendienteCaja = $montoPorCobrar + $montoSinFactura
$pctFacturado = Pct $montoFacturado $totalVentas

$fechaCorteTxt = $FechaCorte.ToString("dd/MM/yyyy")
$generadoTs = Get-Date -Format "dd/MM/yyyy HH:mm"

$sinFacturaTabBtnHtml = ""
$sinFacturaPanelHtml = ""
if ($grpSinFactura.Count -gt 0) {
  $sinFacturaTabBtnHtml = "<button class=`"tab-btn`" data-tab=`"tab-sinfacturar`" onclick=`"mostrarTab('tab-sinfacturar', this)`">Ventas Sin Facturar</button>"
  $sinFacturaPanelHtml = @"
<div id="tab-sinfacturar" class="tab-panel">
<div class="panel-head"><div class="eyebrow">Pendiente</div><h2>Ventas Sin Facturar</h2><p class="panel-desc">Operaciones ejecutadas a${e_u}n sin facturar, desglosadas por cliente.</p></div>
$sinFacturaChartSvg
<div class="table-scroll">
<table><thead><tr><th>Cliente</th><th class=n>Docs.</th><th class=n>Monto USD</th><th class=n>% part.</th></tr></thead>
<tbody>
$filasSinFacturaCliente
</tbody>
<tfoot><tr><td>TOTAL</td><td class=n>$($grpSinFactura.Count)</td><td class=n>$(FmtCell $montoSinFactura)</td><td class=n>100,0%</td></tr></tfoot></table>
</div>
<div class="panel-head"><div class="eyebrow">Detalle</div><h2>Detalle por documento</h2></div>
<div class="table-scroll">
<table><thead><tr><th>Cliente</th><th>PO</th><th>Descripci${e_o}n</th><th>Fecha PO</th><th class=n>Monto USD</th></tr></thead>
<tbody>
$filasSinFacturaDetalle
</tbody></table>
</div>
</div>
"@
}

$inproccaTabBtnHtml = ""
$inproccaPanelHtml = ""
if ($inproccaRows.Count -gt 0) {
  $inproccaTabBtnHtml = "<button class=`"tab-btn`" data-tab=`"tab-inprocca`" onclick=`"mostrarTab('tab-inprocca', this)`">INPROCCA</button>"
  $inproccaPanelHtml = @"
<div id="tab-inprocca" class="tab-panel">
<div class="panel-head"><div class="eyebrow">Cliente</div><h2>Estado de cuenta &mdash; INPROCCA</h2></div>
<div class="callout">Detalle del historial de cr${e_e}dito, abonos y saldo de INPROCCA seg${e_u}n la hoja "CC INPROCCA" del Excel. Estos montos son el registro manual de la cuenta y pueden no coincidir exactamente con el resumen agregado de la pesta${e_n}a "Cuentas por cobrar".</div>
<div class="callout">INPROCCA ha realizado un abono de `$25.000 el d${e_i}a 07/05/2026, el d${e_i}a 14/07/2026 realiza un segundo abono de `$25.000 y el 06/08/2026 realiza un tercer abono de `$10.000, quedando un saldo pendiente a la fecha de `$131.594.</div>
<div class="table-scroll">
<table class="wide-table"><colgroup>
<col style="width:8%"><col style="width:14%"><col style="width:12%">
<col style="width:11%"><col style="width:11%"><col style="width:11%"><col style="width:11%"><col style="width:11%"><col style="width:11%">
</colgroup>
<thead>
<tr><th class=ctr>Fecha</th><th>Documento</th><th class=n>Monto</th><th class=ctr>$($inproccaHeaders[0])</th><th class=n>$($inproccaHeaders[1])</th><th class=ctr>$($inproccaHeaders[2])</th><th class=n>$($inproccaHeaders[3])</th><th class=ctr>$($inproccaHeaders[4])</th><th class=n>$($inproccaHeaders[5])</th></tr>
</thead>
<tbody>
$filasInprocca
</tbody>
<tfoot><tr><td colspan=2>TOTAL</td><td class=n>$(FmtUSD $inproccaTotMonto)</td><td class=ctr>`$ -</td><td class=n>`$ -</td><td class=ctr>`$ -</td><td class=n>$(FmtUSD $inproccaTotCr2)</td><td class=ctr>$(FmtUSD $inproccaTotAb3)</td><td class=n>$(FmtUSD $inproccaTotCr3)</td></tr></tfoot></table>
</div>
</div>
"@
}

$cxpTabBtnHtml = ""
$cxpPanelHtml = ""
if ($cxpProveedores.Count -gt 0 -or $cxpProvisiones.Count -gt 0) {
  $filasCxpProveedores = ($cxpProveedores | ForEach-Object {
    "<tr><td>$($_.Proveedor)</td><td class=n>$(FmtCell $_.Monto)</td></tr>"
  }) -join "`n"
  $filasCxpProvisiones = ($cxpProvisiones | ForEach-Object {
    "<tr><td>$($_.Concepto)</td><td class=n>$(FmtCell $_.Monto)</td></tr>"
  }) -join "`n"
  $cxpAvisoHtml = ""
  if (-not $cxpProvisionesCuadra) {
    $cxpAvisoHtml = "<div class=`"callout`">El total de provisiones que trae el Excel no coincide con la suma de sus renglones; aqu${e_i} se muestra la suma de los renglones.</div>"
  }
  $cxpTabBtnHtml = "<button class=`"tab-btn`" data-tab=`"tab-cxp`" onclick=`"mostrarTab('tab-cxp', this)`">Cuentas por pagar</button>"
  $cxpPanelHtml = @"
<div id="tab-cxp" class="tab-panel">
<div class="panel-head"><div class="eyebrow">Obligaciones</div><h2>Cuentas por pagar</h2><p class="panel-desc">Saldos pendientes con proveedores y provisiones para ejecuci${e_o}n de proyectos, seg${e_u}n la hoja "cuentas por pagar" del Excel.</p></div>
$cxpAvisoHtml
<div class="table-scroll">
<table><thead><tr><th>Proveedor</th><th class=n>Monto USD</th></tr></thead>
<tbody>
$filasCxpProveedores
</tbody>
<tfoot><tr><td>SUB-TOTAL</td><td class=n>$(FmtCell $cxpSubtotalProveedores)</td></tr></tfoot></table>
</div>
<div class="panel-head"><div class="eyebrow">Provisiones</div><h2>Provisiones para ejecuci${e_o}n de proyectos</h2></div>
<div class="table-scroll">
<table><thead><tr><th>Concepto</th><th class=n>Monto USD</th></tr></thead>
<tbody>
$filasCxpProvisiones
</tbody>
<tfoot><tr><td>SUB-TOTAL</td><td class=n>$(FmtCell $cxpTotalProvisiones)</td></tr></tfoot></table>
</div>
<div class="table-scroll">
<table><tfoot><tr><td>TOTAL CUENTAS POR PAGAR</td><td class=n>$(FmtCell $cxpTotalGeneral)</td></tr></tfoot></table>
</div>
</div>
"@
}

$dispoTabBtnHtml = ""
$dispoPanelHtml = ""
if ($dispoRows.Count -gt 0) {
  $dispoTabBtnHtml = "<button class=`"tab-btn`" data-tab=`"tab-dispo`" onclick=`"mostrarTab('tab-dispo', this)`">Disponibilidad</button>"
  $dispoPanelHtml = @"
<div id="tab-dispo" class="tab-panel">
<div class="table-scroll">
<table><thead><tr><th>Cuenta</th><th class=n>Importe moneda origen</th><th class=n>Importe USD</th></tr></thead>
<tbody>
$filasDispo
</tbody>
<tfoot><tr><td colspan=2>TOTAL DISPONIBLE</td><td class=n>$(FmtCell $totalDispoUSD)</td></tr></tfoot></table>
</div>
</div>
"@
}

$html = @"
<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ventas $($FechaCorte.Year) &mdash; Informe de Gerencia</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box}
:root{
  --brand-teal:#00877A;
  --brand-teal-dark:#045C53;
  --brand-teal-deep:#0B3D38;
  --slate-600:#4C6360;
  --slate-400:#84A19C;
  --paper:#F3FAF8;
  --card:#FFFFFF;
  --line:#DCEBE7;
  --teal-100:#E1F4F0;
  --amber-600:#B9791F; --amber-100:#FBF0DD;
  --green-600:#256B4C; --green-100:#E3F1E8;
  --red-600:#B3372C;  --red-100:#FBEAE8;
  --font-display:'Space Grotesk',Arial,sans-serif;
  --font-body:'Inter',-apple-system,'Segoe UI',Arial,sans-serif;
  --font-mono:'IBM Plex Mono',ui-monospace,Consolas,monospace;
}
html{background:var(--paper)}
body{font-family:var(--font-body);margin:0;background:var(--paper);color:var(--brand-teal-deep);font-size:14px;line-height:1.5}
.wrap{max-width:1400px;margin:0 auto;padding:0 40px 64px}
.top-sticky{position:sticky;top:0;z-index:30;background:var(--paper)}
header{background:linear-gradient(120deg,#FFFFFF 0%,#EAF7F4 100%);border-bottom:3px solid var(--brand-teal);padding:22px 0 18px}
.header-top{display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:16px}
.brand-block{display:flex;align-items:center;gap:14px}
header .logo{height:30px;display:block}
.eyebrow{font-family:var(--font-mono);font-size:11px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--brand-teal)}
header h1{font-family:var(--font-display);margin:2px 0 0;font-size:23px;font-weight:700;letter-spacing:.2px;color:var(--brand-teal-deep);text-wrap:balance}
.header-sub{margin:10px 0 0;font-size:13px;color:var(--slate-600)}
.header-right{text-align:right;display:flex;flex-direction:column;align-items:flex-end;gap:8px}
.corte-label{font-family:var(--font-mono);font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--slate-600)}
.refresh-row{display:flex;align-items:center;gap:10px}
.refresh-ts{font-size:11px;color:var(--slate-400)}
.refresh-btn{appearance:none;background:var(--card);border:1px solid var(--brand-teal);color:var(--brand-teal-dark);font:inherit;font-weight:600;font-size:12px;padding:6px 13px;border-radius:6px;cursor:pointer;display:inline-flex;align-items:center;gap:6px;transition:background .15s}
.refresh-btn:hover{background:var(--teal-100)}
.refresh-btn:active .refresh-ic{transform:rotate(180deg)}
.refresh-ic{display:inline-block;transition:transform .4s ease}
.tabs{display:flex;gap:22px;background:var(--card);border-bottom:1px solid var(--line);padding:0 2px;overflow-x:auto}
.tab-btn{appearance:none;border:none;background:none;font-family:var(--font-body);font-weight:600;font-size:13.5px;color:var(--slate-600);padding:12px 2px;cursor:pointer;border-bottom:3px solid transparent;margin-bottom:-1px;white-space:nowrap;flex:0 0 auto}
.tab-btn:hover{color:var(--brand-teal-dark)}
.tab-btn.active{color:var(--brand-teal-dark);border-bottom-color:var(--brand-teal);font-weight:700}
.tab-panel{display:none}
.tab-panel.active{display:block}
h2{font-family:var(--font-display);font-size:16px;font-weight:700;margin:0;color:var(--brand-teal-deep);text-wrap:balance}
.panel-head{margin:32px 0 12px}
.panel-desc{margin:3px 0 0;font-size:12.5px;color:var(--slate-600)}
.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin:22px 0}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px}
.kpi.featured{background:var(--teal-100);border-color:var(--brand-teal)}
.kpi .lbl{font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--slate-600);font-weight:600}
.kpi .val{font-family:var(--font-mono);font-size:21px;font-weight:600;margin-top:6px;font-variant-numeric:tabular-nums}
.kpi .val.teal{color:var(--brand-teal-dark)} .kpi .val.green{color:var(--green-600)} .kpi .val.amber{color:var(--amber-600)} .kpi .val.red{color:var(--red-600)} .kpi .val.slate{color:var(--slate-600)}
.kpi .sub{font-size:11.5px;color:var(--slate-400);margin-top:4px}
.table-scroll{overflow-x:auto;border:1px solid var(--line);border-radius:10px}
.table-scroll table{border-radius:0}
table{width:100%;border-collapse:collapse;background:var(--card);font-size:13px}
th{background:var(--teal-100);color:var(--brand-teal-deep);text-align:left;padding:9px 12px;font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;font-weight:700;border-bottom:2px solid var(--brand-teal);white-space:nowrap}
th.n{text-align:right}
td.ctr,th.ctr{text-align:center}
td{padding:8px 12px;border-top:1px solid var(--line)}
td.n{text-align:right}
td.n,.mono{font-family:var(--font-mono);font-variant-numeric:tabular-nums}
tbody tr:nth-child(even){background:#FAFDFC}
tbody tr:hover{background:var(--teal-100)}
tbody tr.resumen{background:#EAF5F2;font-weight:700;color:var(--brand-teal-dark);border-top:2px solid var(--brand-teal);border-bottom:2px solid var(--brand-teal)}
tfoot td{font-weight:700;background:#EAF5F2;color:var(--brand-teal-dark);border-top:2px solid var(--brand-teal);border-bottom:2px solid var(--brand-teal)}
td.bar{width:160px} td.bar div{height:8px;background:var(--brand-teal);border-radius:4px}
.chart-card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px 12px 6px;margin-bottom:18px;position:relative}
.chart-svg{width:100%;height:auto;display:block;overflow:visible}
.chart-grid{stroke:var(--line);stroke-width:1}
.chart-ylabel,.chart-xlabel{font-family:var(--font-mono);font-size:10px;fill:var(--slate-400)}
.chart-xlabel{fill:var(--slate-600)}
.chart-toplabel{font-family:var(--font-mono);font-size:12px;font-weight:600;fill:var(--brand-teal-dark)}
.bar-rect{fill:var(--brand-teal);transition:opacity .1s}
.bar-rect.peak{fill:var(--brand-teal-dark)}
.bar-hit{fill:transparent;cursor:pointer;outline:none}
.bar-hit:hover + .bar-rect,.bar-hit:focus + .bar-rect{opacity:.72}
.chart-tooltip{position:absolute;pointer-events:none;background:var(--brand-teal-deep);color:#fff;font-family:var(--font-body);font-size:12px;padding:7px 11px;border-radius:6px;opacity:0;transition:opacity .1s;transform:translate(-50%,-100%);white-space:nowrap;z-index:5;text-align:left}
.chart-tooltip.show{opacity:1}
.chart-tooltip b{font-family:var(--font-mono)}
.tooltip-title{font-weight:600;margin-bottom:3px}
.tooltip-row{display:flex;align-items:center;gap:6px}
.seg-cobrado{fill:var(--green-600)}
.seg-porcobrar{fill:var(--amber-600)}
.seg-sinfacturar{fill:var(--red-600)}
.chart-legend{display:flex;gap:16px;padding:0 4px 8px;font-size:11.5px;color:var(--slate-600)}
.legend-item{display:inline-flex;align-items:center;gap:6px}
.legend-swatch{width:10px;height:10px;border-radius:2px;display:inline-block;flex:0 0 auto}
.swatch-cobrado{background:var(--green-600)}
.swatch-porcobrar{background:var(--amber-600)}
.swatch-sinfacturar{background:var(--red-600)}
.zero{color:var(--slate-400)}
.neg{color:var(--red-600)}
.badge{display:inline-block;border-radius:20px;padding:3px 10px;font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.03em;white-space:nowrap}
.badge-green{background:var(--green-100);color:var(--green-600)}
.badge-amber{background:var(--amber-100);color:var(--amber-600)}
.badge-red{background:var(--red-100);color:var(--red-600)}
.wide-table th,.wide-table td{white-space:normal}
.note{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px 22px;font-size:13.5px;line-height:1.65}
.note li{margin-bottom:9px}
.callout{background:var(--amber-100);border-left:3px solid var(--amber-600);border-radius:0 8px 8px 0;padding:10px 16px;font-size:12.5px;color:var(--slate-600);margin:10px 0 0}
footer{margin-top:40px;font-size:11px;color:var(--slate-400);text-align:center}
@media (max-width:1100px){.kpis{grid-template-columns:repeat(3,1fr)}}
@media (max-width:700px){
.wrap{padding:0 14px 40px}
header{padding:16px 0 14px}
.header-right{align-items:flex-start;text-align:left;width:100%}
header h1{font-size:19px}
.kpis{grid-template-columns:repeat(2,1fr);gap:9px}
.kpi{padding:12px}
.kpi .val{font-size:17px}
table{font-size:11.5px}
th,td{padding:6px 8px}
td.bar{width:80px}
}
@media print{.top-sticky{position:static} body{background:#fff} .tabs{display:none} .tab-panel{display:block !important}}
</style></head><body><div class="wrap">
<div class="top-sticky">
<header>
<div class="header-top">
<div class="brand-block">
$logoImgTag
<div>
<div class="eyebrow">Reporte de gerencia</div>
<h1>Ventas $($FechaCorte.Year)</h1>
</div>
</div>
<div class="header-right">
<div class="corte-label">Corte al $fechaCorteTxt</div>
<div class="refresh-row">
<span class="refresh-ts">Actualizado $generadoTs</span>
<button class="refresh-btn" onclick="actualizarDashboard(this)" title="Vuelve a cargar la ${e_u}ltima versi${e_o}n publicada de este dashboard">
<span class="refresh-ic">&#8635;</span> Actualizar
</button>
</div>
</div>
</div>
<p class="header-sub">Fuente: hoja "VENTAS $($FechaCorte.Year)" $([char]0xB7) $docsTotal documentos $([char]0xB7) Cifras en USD</p>
</header>
<div class="tabs">
<button class="tab-btn active" data-tab="tab-ventas" onclick="mostrarTab('tab-ventas', this)">Ventas</button>
<button class="tab-btn" data-tab="tab-cobranza" onclick="mostrarTab('tab-cobranza', this)">Cuentas por cobrar</button>
$sinFacturaTabBtnHtml
$inproccaTabBtnHtml
$dispoTabBtnHtml
$cxpTabBtnHtml
</div>
</div>

<div id="tab-ventas" class="tab-panel active">
<div class="kpis">
<div class="kpi featured"><div class="lbl">Ventas totales (con IVA e IGTF)</div><div class="val teal">$(Fmt0 $totalVentas)</div><div class="sub">$docsTotal operaciones $([char]0xB7) ticket promedio $(Fmt0 $ticketProm)</div></div>
<div class="kpi"><div class="lbl">Base imponible + exento</div><div class="val teal">$(Fmt0 $totalBaseExe)</div><div class="sub">Ingreso neto de impuestos</div></div>
<div class="kpi"><div class="lbl">Facturado y cobrado</div><div class="val green">$(Fmt0 $montoCobrado)</div><div class="sub">$(Fmt1Pct $pctCobrado) de las ventas</div></div>
<div class="kpi"><div class="lbl">Facturado por cobrar</div><div class="val amber">$(Fmt0 $montoPorCobrar)</div><div class="sub">$(Fmt1Pct $pctPorCobrar) de las ventas</div></div>
<div class="kpi"><div class="lbl">Pendiente de facturar</div><div class="val red">$(Fmt0 $montoSinFactura)</div><div class="sub">$(Fmt1Pct $pctSinFactura) de las ventas</div></div>
<div class="kpi"><div class="lbl">Impuestos gestionados</div><div class="val slate">$(Fmt0 $totalImpuestos)</div><div class="sub">IVA $(Fmt0 $totalIVA) $([char]0xB7) IGTF $(Fmt0 $totalIGTF)</div></div>
</div>
<div class="panel-head"><div class="eyebrow">S${e_i}ntesis</div><h2>Resumen ejecutivo</h2><p class="panel-desc">Lectura general de ventas, cobranza y concentraci${e_o}n de cartera.</p></div>
<div class="note"><ul>
<li>Las ventas acumuladas del ejercicio $($FechaCorte.Year) alcanzan <b>USD $(Fmt0 $totalVentas)</b> distribuidas en <b>$docsTotal operaciones</b>, de las cuales <b>$docsFacturados</b> ya est${e_a}n facturadas (<b>$(Fmt0 $montoFacturado)</b>, $(Fmt1Pct $pctFacturado)).</li>
<li>La conversi${e_o}n a caja es el punto cr${e_i}tico: solo <b>$(Fmt1Pct $pctCobrado)</b> de la venta est$e_a efectivamente cobrada. Entre cuentas por cobrar ($(Fmt0 $montoPorCobrar)) y ventas sin facturar ($(Fmt0 $montoSinFactura)) hay <b>USD $(Fmt0 $montoPendienteCaja)</b> pendientes de convertir en efectivo ($(Fmt1Pct $pctPendienteCaja) del total).</li>
<li>Concentraci${e_o}n de cartera: <b>$($clienteTop.Cliente)</b> representa $(Fmt1Pct $clienteTopPct) de la facturaci${e_o}n y los tres primeros clientes concentran <b>$(Fmt1Pct $top3Pct)</b>.</li>
<li>El mes de mayor actividad fue <b>$($mesTop.Label)</b> con $(Fmt0 $mesTop.Monto) ($(Fmt1Pct $mesTopPct) del a${e_n}o).</li>
<li>La estructura tributaria muestra <b>$(Fmt0 $totalIVA)</b> de IVA y <b>$(Fmt0 $totalIGTF)</b> de IGTF; el rubro exento asciende a $(Fmt0 $totalExento) ($(Fmt1Pct (Pct $totalExento $totalVentas))).</li>
</ul></div>
<div class="panel-head"><div class="eyebrow">Tendencia</div><h2>Evoluci${e_o}n mensual de las ventas</h2><p class="panel-desc">Documentos y monto vendido por mes del ejercicio $($FechaCorte.Year).</p></div>
<div class="callout">Los documentos a${e_u}n no facturados se asignan al mes de su orden de compra (PO), por no disponer de fecha de factura.</div>
$monthlyChartSvg
<div class="table-scroll">
<table><thead><tr><th>Mes</th><th class=n>Docs.</th><th class=n>Monto USD</th><th class=n>% del a${e_n}o</th></tr></thead>
<tbody>
$filasMensual
</tbody>
<tfoot><tr><td>TOTAL</td><td class=n>$docsTotal</td><td class=n>$(FmtCell $totalVentas)</td><td class=n>100,0%</td></tr></tfoot></table>
</div>
<div class="panel-head"><div class="eyebrow">Cartera</div><h2>Ventas por cliente</h2><p class="panel-desc">Participaci${e_o}n de cada cliente sobre el total facturado.</p></div>
$rankChartSvg
<div class="table-scroll">
<table><thead><tr><th>Cliente</th><th class=n>Docs.</th><th class=n>Venta total</th><th class=n>% part.</th></tr></thead>
<tbody>
$filasClienteVentas
</tbody>
<tfoot><tr><td>TOTAL</td><td class=n>$docsTotal</td><td class=n>$(FmtCell $totalVentas)</td><td class=n>100,0%</td></tr></tfoot></table>
</div>
<div class="panel-head"><div class="eyebrow">Top 10</div><h2>Diez operaciones de mayor monto</h2><p class="panel-desc">Las facturas y ${e_o}rdenes de mayor valor del per${e_i}odo.</p></div>
<div class="table-scroll">
<table><thead><tr><th>Documento</th><th>Cliente</th><th>Concepto</th><th class=n>Monto USD</th><th>Estado</th></tr></thead>
<tbody>
$filasTop10
</tbody></table>
</div>
<div class="panel-head"><div class="eyebrow">Acciones</div><h2>Conclusiones y recomendaciones</h2></div>
<div class="note"><ul>
<li><b>Acelerar la facturaci${e_o}n pendiente.</b> Existen $($grpSinFactura.Count) operaciones a${e_u}n sin facturar por USD $(Fmt0 $montoSinFactura). Emitir estos documentos es la acci${e_o}n de mayor impacto inmediato sobre el flujo de caja.</li>
<li><b>Gesti${e_o}n de cobranza focalizada.</b> Concentrar el esfuerzo en los saldos de mayor antig${e_ud}edad y en los clientes con mayor exposici${e_o}n por cobrar.</li>
<li><b>Diversificaci${e_o}n de cartera.</b> Con $(Fmt1Pct $top3Pct) de la venta en los tres primeros clientes, conviene desarrollar nuevos contratos para reducir la dependencia de pocos clientes.</li>
<li><b>Seguimiento por PO.</b> Varias ${e_o}rdenes se facturan en partes; un control por PO evitar${e_i}a saldos abiertos no facturados.</li>
</ul></div>
</div>

<div id="tab-cobranza" class="tab-panel">
<div class="panel-head"><div class="eyebrow">Estado</div><h2>Situaci${e_o}n de facturaci${e_o}n y cobranza</h2><p class="panel-desc">Documentos agrupados por estado de cobro.</p></div>
<div class="table-scroll">
<table><thead><tr><th>Estado</th><th class=n>Docs.</th><th class=n>Monto USD</th><th class=n>% del total</th></tr></thead>
<tbody>
<tr><td>$(Badge "Facturado y cobrado")</td><td class=n>$($grpCobrado.Count)</td><td class=n>$(FmtCell $montoCobrado)</td><td class=n>$(Fmt1Pct $pctCobrado)</td></tr>
<tr><td>$(Badge "Facturado sin cobrar")</td><td class=n>$($grpPorCobrar.Count)</td><td class=n>$(FmtCell $montoPorCobrar)</td><td class=n>$(Fmt1Pct $pctPorCobrar)</td></tr>
<tr><td>$(Badge "Sin facturar")</td><td class=n>$($grpSinFactura.Count)</td><td class=n>$(FmtCell $montoSinFactura)</td><td class=n>$(Fmt1Pct $pctSinFactura)</td></tr>
</tbody>
<tfoot><tr><td>TOTAL</td><td class=n>$docsTotal</td><td class=n>$(FmtCell $totalVentas)</td><td class=n>100,0%</td></tr></tfoot></table>
</div>
<div class="panel-head"><div class="eyebrow">Cobranza $([char]0xB7) ONG</div><h2>Cartera pendiente $([char]0x2014) ONG</h2><p class="panel-desc">Saldo cobrado, por cobrar y sin facturar de clientes ONG.</p></div>
<div class="table-scroll">
<table><thead><tr><th>Cliente</th><th class=n>Cobrado</th><th class=n>Por cobrar</th><th class=n>Sin facturar</th><th class=n>Total pendiente</th></tr></thead>
<tbody>
$($carteraONGResult.Filas)
</tbody>
<tfoot>$($carteraONGResult.Tfoot)</tfoot></table>
</div>
<div class="panel-head"><div class="eyebrow">Cobranza $([char]0xB7) Petr${e_o}leo</div><h2>Cartera pendiente $([char]0x2014) Petr${e_o}leo</h2><p class="panel-desc">Saldo cobrado, por cobrar y sin facturar de clientes del sector petrolero.</p></div>
<div class="table-scroll">
<table><thead><tr><th>Cliente</th><th class=n>Cobrado</th><th class=n>Por cobrar</th><th class=n>Sin facturar</th><th class=n>Total pendiente</th></tr></thead>
<tbody>
$($carteraPetroleoResult.Filas)
</tbody>
<tfoot>$($carteraPetroleoResult.Tfoot)</tfoot></table>
</div>
$carteraOtrosPanelHtml
<div class="panel-head"><div class="eyebrow">Aging</div><h2>Antig${e_ud}edad de las cuentas por cobrar</h2><p class="panel-desc">Distribuci${e_o}n de la cartera por cobrar, seg${e_u}n d${e_i}as desde la factura.</p></div>
<div class="table-scroll">
<table><thead><tr><th>Tramo (d${e_i}as desde la factura)</th><th class=n>Monto USD</th><th class=n>% de la cartera</th></tr></thead>
<tbody>
$filasAging
</tbody>
<tfoot><tr><td>TOTAL POR COBRAR</td><td class=n>$(FmtCell $montoPorCobrar)</td><td class=n>100,0%</td></tr></tfoot></table>
</div>
</div>

$sinFacturaPanelHtml

$inproccaPanelHtml

$dispoPanelHtml

$cxpPanelHtml

<footer>Informe generado autom${e_a}ticamente a partir de "$([System.IO.Path]::GetFileName($ExcelPath))" $([char]0x2014) hoja $SheetName. $([char]0xB7) $generadoTs</footer>
</div>
<script>
function actualizarDashboard(btn) {
  btn.disabled = true;
  var url = window.location.pathname + '?_=' + Date.now() + window.location.hash;
  window.location.replace(url);
}
document.querySelectorAll('.chart-card').forEach(function (card) {
  var tooltip = card.querySelector('.chart-tooltip');
  if (!tooltip) { return; }
  card.querySelectorAll('.bar-hit').forEach(function (hit) {
    function mostrar() {
      var label = hit.getAttribute('data-label');
      tooltip.textContent = '';
      var titleEl = document.createElement('div');
      titleEl.className = 'tooltip-title';
      titleEl.textContent = label;
      tooltip.appendChild(titleEl);
      if (hit.hasAttribute('data-cobrado')) {
        var filas = [
          ['Cobrado', hit.getAttribute('data-cobrado'), 'swatch-cobrado'],
          ['Por cobrar', hit.getAttribute('data-porcobrar'), 'swatch-porcobrar'],
          ['Sin facturar', hit.getAttribute('data-sinfacturar'), 'swatch-sinfacturar']
        ];
        filas.forEach(function (f) {
          var fila = document.createElement('div');
          fila.className = 'tooltip-row';
          var sw = document.createElement('span');
          sw.className = 'legend-swatch ' + f[2];
          fila.appendChild(sw);
          fila.appendChild(document.createTextNode(f[0] + ': '));
          var strong = document.createElement('b');
          strong.textContent = f[1];
          fila.appendChild(strong);
          tooltip.appendChild(fila);
        });
      } else {
        var value = hit.getAttribute('data-value');
        var docs = hit.getAttribute('data-docs');
        var fila = document.createElement('div');
        var strong = document.createElement('b');
        strong.textContent = value;
        fila.appendChild(strong);
        var resto = ' USD';
        if (docs) { resto += '  ' + String.fromCharCode(183) + '  ' + docs + ' docs.'; }
        fila.appendChild(document.createTextNode(resto));
        tooltip.appendChild(fila);
      }
      tooltip.classList.add('show');
      var cardRect = card.getBoundingClientRect();
      var hitRect = hit.getBoundingClientRect();
      tooltip.style.left = (hitRect.left - cardRect.left + hitRect.width / 2) + 'px';
      tooltip.style.top = (hitRect.top - cardRect.top) + 'px';
    }
    function ocultar() { tooltip.classList.remove('show'); }
    hit.addEventListener('mouseenter', mostrar);
    hit.addEventListener('focus', mostrar);
    hit.addEventListener('mouseleave', ocultar);
    hit.addEventListener('blur', ocultar);
  });
});
function mostrarTab(id, btn) {
  document.querySelectorAll('.tab-panel').forEach(function(p) { p.classList.remove('active'); });
  document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
}
</script>
</body></html>
"@

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutputPath, $html, $utf8NoBom)

Write-Host "Reporte generado en: $OutputPath"
Write-Host "Ventas totales: $(Fmt0 $totalVentas)"
Write-Host "Facturado y cobrado: $(Fmt0 $montoCobrado)"
Write-Host "Facturado sin cobrar: $(Fmt0 $montoPorCobrar)"
Write-Host "Sin facturar: $(Fmt0 $montoSinFactura)"
