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

$raw = New-Object System.Collections.Generic.List[object]
$firstRow = 2
$r = $firstRow
while ($true) {
  $totalText = $ws.Cells.Item($r, 4).Text
  if ([string]::IsNullOrWhiteSpace($totalText)) { break }

  $facturado = ($ws.Cells.Item($r, 15).Text).Trim().ToUpper()
  $cobranza  = ($ws.Cells.Item($r, 16).Text).Trim().ToUpper()
  $cliente   = ($ws.Cells.Item($r, 10).Text).Trim().ToUpper()
  $nroFactura = ($ws.Cells.Item($r, 2).Text).Trim()
  $descripcion = ($ws.Cells.Item($r, 11).Text).Trim()

  $fechaFacTxt = ($ws.Cells.Item($r, 3).Text).Trim()
  $fechaPOTxt  = ($ws.Cells.Item($r, 12).Text).Trim()
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
    TotalFac      = [double]$ws.Cells.Item($r, 4).Value2
    BaseImponible = [double]$ws.Cells.Item($r, 5).Value2
    Exento        = [double]$ws.Cells.Item($r, 6).Value2
    IVA           = [double]$ws.Cells.Item($r, 7).Value2
    IGTF          = [double]$ws.Cells.Item($r, 8).Value2
    PO            = ($ws.Cells.Item($r, 9).Text).Trim()
    Cliente       = $cliente
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
$inproccaFechas = @("", "", "")
$hojaInprocca = $null
foreach ($hoja in $wb.Worksheets) { if ($hoja.Name -eq "CC INPROCCA") { $hojaInprocca = $hoja } }
if ($hojaInprocca) {
  $wsi = $hojaInprocca
  $rowsI = $wsi.UsedRange.Rows.Count
  $inproccaFechas = @(
    ($wsi.Cells.Item(4, 5).Text).Trim(),
    ($wsi.Cells.Item(4, 7).Text).Trim(),
    ($wsi.Cells.Item(4, 9).Text).Trim()
  )
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

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Filas leidas: $($raw.Count)"
Write-Host "Filas leidas de CC INPROCCA: $($inproccaRows.Count)"
Write-Host "Filas leidas de DISPONIBILIDAD: $($dispoRows.Count)"
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
  $ancho = if ($maxMensual -gt 0) { Pct $_.Monto $maxMensual } else { 0 }
  "<tr><td>$($_.Label)</td><td class=n>$($_.Docs)</td><td class=n>$(Fmt0 $_.Monto)</td><td class=n>$(Fmt1Pct $pct)</td><td class=bar><div style='width:$([math]::Round($ancho,1))%'></div></td></tr>"
}) -join "`n"

# ---------- Analisis por cliente ----------
$porCliente = $raw | Group-Object Cliente | ForEach-Object {
  $g = $_.Group
  $venta = ($g | Measure-Object TotalFac -Sum).Sum
  [PSCustomObject]@{
    Cliente     = $_.Name
    Docs        = $_.Count
    VentaTotal  = $venta
    Cobrado     = ($g | Where-Object { $_.Estado -eq "Facturado y cobrado" } | Measure-Object TotalFac -Sum).Sum
    PorCobrar   = ($g | Where-Object { $_.Estado -eq "Facturado sin cobrar" } | Measure-Object TotalFac -Sum).Sum
    SinFacturar = ($g | Where-Object { $_.Estado -eq "Sin facturar" } | Measure-Object TotalFac -Sum).Sum
  }
} | Sort-Object VentaTotal -Descending

$filasClienteVentas = ($porCliente | ForEach-Object {
  $pct = Pct $_.VentaTotal $totalVentas
  "<tr><td><b>$($_.Cliente)</b></td><td class=n>$($_.Docs)</td><td class=n>$(Fmt0 $_.VentaTotal)</td><td class=n>$(Fmt1Pct $pct)</td></tr>"
}) -join "`n"

$carteraCliente = $porCliente | ForEach-Object {
  $pend = $_.PorCobrar + $_.SinFacturar
  [PSCustomObject]@{
    Cliente = $_.Cliente
    Cobrado = $_.Cobrado
    PorCobrar = $_.PorCobrar
    SinFacturar = $_.SinFacturar
    Pendiente = $pend
  }
} | Sort-Object Pendiente -Descending

$filasCarteraCliente = ($carteraCliente | ForEach-Object {
  "<tr><td><b>$($_.Cliente)</b></td><td class=n>$(Fmt0 $_.Cobrado)</td><td class=n>$(Fmt0 $_.PorCobrar)</td><td class=n>$(Fmt0 $_.SinFacturar)</td><td class=n>$(Fmt0 $_.Pendiente)</td></tr>"
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
  "<tr><td>$($_.Key)</td><td class=n>$(Fmt0 $_.Value)</td><td class=n>$(Fmt1Pct $pct)</td></tr>"
}) -join "`n"

# ---------- Estado de cuenta INPROCCA (hoja "CC INPROCCA") ----------
$fechasOcultarInprocca = @("07/05/2026", "14/07/2026")
$filasInprocca = ($inproccaRows | Where-Object { $fechasOcultarInprocca -notcontains $_.Fecha } | ForEach-Object {
  $esResumen = $_.Doc -match "(?i)deuda|saldo|conciliado"
  $docHtml = if ($esResumen) { "<b>$($_.Doc)</b>" } else { $_.Doc }
  $rowClass = if ($esResumen) { " class='resumen'" } else { "" }
  "<tr$rowClass><td>$($_.Fecha)</td><td>$docHtml</td><td class=n>$($_.Monto)</td><td class=n>$($_.Ab1)</td><td class=n>$($_.Cr1)</td><td class=n>$($_.Ab2)</td><td class=n>$($_.Cr2)</td><td class=n>$($_.Ab3)</td><td class=n>$($_.Cr3)</td></tr>"
}) -join "`n"

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
  $doc = if ($_.NroFactura -ne "") { $_.NroFactura } else { "<i>por facturar</i>" }
  "<tr><td>$doc</td><td>$($_.Cliente)</td><td>$($_.Descripcion)</td><td class=n>$(Fmt0 $_.TotalFac)</td><td>$($_.Estado)</td></tr>"
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

$inproccaTabBtnHtml = ""
$inproccaPanelHtml = ""
if ($inproccaRows.Count -gt 0) {
  $inproccaTabBtnHtml = "<button class=`"tab-btn`" data-tab=`"tab-inprocca`" onclick=`"mostrarTab('tab-inprocca', this)`">INPROCCA</button>"
  $inproccaPanelHtml = @"
<div id="tab-inprocca" class="tab-panel">
<h2>Estado de cuenta &mdash; INPROCCA</h2>
<p style="font-size:12.5px;color:#667873">Detalle del historial de cr${e_e}dito, abonos y saldo de INPROCCA seg${e_u}n la hoja "CC INPROCCA" del Excel. Estos montos son el registro manual de la cuenta y pueden no coincidir exactamente con el resumen agregado de la pesta${e_n}a "Cuentas por cobrar".</p>
<div class="table-scroll">
<table><thead>
<tr><th></th><th></th><th></th><th colspan=2>$($inproccaFechas[0])</th><th colspan=2>$($inproccaFechas[1])</th><th colspan=2>$($inproccaFechas[2])</th></tr>
<tr><th>Fecha</th><th>Documento</th><th class=n>Monto</th><th class=n>Abonos</th><th class=n>Cr${e_e}dito</th><th class=n>Abonos</th><th class=n>Cr${e_e}dito</th><th class=n>Abonos</th><th class=n>Cr${e_e}dito</th></tr>
</thead>
<tbody>
$filasInprocca
</tbody></table>
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
<h2>$dispoTitulo</h2>
<p style="font-size:12.5px;color:#667873">Tasa COP: <b>$dispoTasaCOP</b> $([char]0xB7) Tasa USDT: <b>$dispoTasaUSDT</b></p>
<table><thead><tr><th>Cuenta</th><th class=n>Importe moneda origen</th><th class=n>Importe USD</th></tr></thead>
<tbody>
$filasDispo
</tbody>
<tfoot><tr><td colspan=2>TOTAL DISPONIBLE</td><td class=n>$(Fmt0 $totalDispoUSD)</td></tr></tfoot></table>
</div>
"@
}

$html = @"
<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<title>Ventas $($FechaCorte.Year) &mdash; Informe de Gerencia</title>
<style>
*{box-sizing:border-box}
body{font-family:-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;margin:0;background:#f4f6f8;color:#263832}
.wrap{max-width:1080px;margin:0 auto;padding:32px 24px 64px}
header{background:linear-gradient(135deg,#00453e,#008275);color:#fff;padding:34px 28px;border-radius:12px}
.header-top{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;margin-bottom:16px}
header .logo{background:#fff;padding:8px 14px;border-radius:8px;height:28px;display:inline-block}
.refresh-block{display:flex;align-items:center;gap:10px}
.refresh-ts{font-size:11px;color:#fff;opacity:.8;text-align:right;line-height:1.4}
.refresh-btn{appearance:none;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.45);color:#fff;font:inherit;font-weight:600;font-size:12px;padding:7px 14px;border-radius:7px;cursor:pointer;display:inline-flex;align-items:center;gap:7px;transition:background .15s}
.refresh-btn:hover{background:rgba(255,255,255,.26)}
.refresh-btn:active .refresh-ic{transform:rotate(180deg)}
.refresh-ic{display:inline-block;transition:transform .4s ease}
header h1{margin:0 0 6px;font-size:26px;letter-spacing:.3px}
header p{margin:0;opacity:.85;font-size:14px}
h2{font-size:17px;margin:38px 0 12px;padding-bottom:8px;border-bottom:2px solid #d7e6e3;color:#00453e}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-top:18px}
.kpi{background:#fff;border-radius:10px;padding:16px 18px;box-shadow:0 1px 3px rgba(0,50,45,.10);border-left:4px solid #008275}
.kpi .lbl{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:#667873;font-weight:600}
.kpi .val{font-size:24px;font-weight:700;margin-top:4px;color:#00453e}
.kpi .sub{font-size:12px;color:#667873;margin-top:3px}
.kpi.ok{border-left-color:#1e874b} .kpi.warn{border-left-color:#d98324} .kpi.bad{border-left-color:#b4232a}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,50,45,.10);font-size:13px}
th{background:#00453e;color:#fff;text-align:left;padding:10px 12px;font-size:11px;text-transform:uppercase;letter-spacing:.5px}
td{padding:9px 12px;border-top:1px solid #eef2f0}
td.n{text-align:right;font-variant-numeric:tabular-nums}
tbody tr:nth-child(even){background:#f6faf9}
tbody tr.resumen{background:#eaf2f0;font-weight:700}
tfoot td{font-weight:700;background:#eaf2f0;border-top:2px solid #cfe1dc}
td.bar{width:180px} td.bar div{height:10px;background:linear-gradient(90deg,#008275,#5fc0b3);border-radius:5px}
.note{background:#fff;border-radius:10px;padding:18px 22px;box-shadow:0 1px 3px rgba(0,50,45,.10);font-size:14px;line-height:1.65}
.note li{margin-bottom:9px}
footer{margin-top:40px;font-size:11px;color:#8a9a95;text-align:center}
.tabs{display:flex;gap:8px;margin:24px 0 28px;border-bottom:2px solid #d7e6e3}
.tab-btn{appearance:none;border:none;background:none;font:inherit;font-weight:700;font-size:14px;color:#667873;padding:10px 18px;cursor:pointer;border-bottom:3px solid transparent;margin-bottom:-2px}
.tab-btn:hover{color:#008275}
.tab-btn.active{color:#00453e;border-bottom-color:#008275}
.tab-panel{display:none}
.tab-panel.active{display:block}
.table-scroll{overflow-x:auto;border-radius:10px}
.table-scroll table{border-radius:0}
@media print{body{background:#fff} .wrap{padding:0} .tabs{display:none} .tab-panel{display:block !important}}
</style></head><body><div class="wrap">
<header>
<div class="header-top">
$logoImgTag
<div class="refresh-block">
<span class="refresh-ts">${e_u}ltima actualizaci${e_o}n del dashboard<br><b>$generadoTs</b></span>
<button class="refresh-btn" onclick="actualizarDashboard(this)" title="Vuelve a cargar la ${e_u}ltima versi${e_o}n publicada de este dashboard">
<span class="refresh-ic">&#8635;</span> Actualizar
</button>
</div>
</div>
<h1>VENTAS $($FechaCorte.Year) &mdash; INFORME DE GERENCIA</h1>
<p>Fuente: hoja "VENTAS $($FechaCorte.Year)" $([char]0xB7) $docsTotal documentos $([char]0xB7) Datos con corte al $fechaCorteTxt $([char]0xB7) Cifras en USD</p>
</header>
<div class="tabs">
<button class="tab-btn active" data-tab="tab-ventas" onclick="mostrarTab('tab-ventas', this)">Ventas</button>
<button class="tab-btn" data-tab="tab-cobranza" onclick="mostrarTab('tab-cobranza', this)">Cuentas por cobrar</button>
$inproccaTabBtnHtml
$dispoTabBtnHtml
</div>
<div class="kpis">
<div class="kpi"><div class="lbl">Ventas totales (con IVA e IGTF)</div><div class="val">$(Fmt0 $totalVentas)</div><div class="sub">$docsTotal operaciones $([char]0xB7) ticket promedio $(Fmt0 $ticketProm)</div></div>
<div class="kpi"><div class="lbl">Base imponible + exento</div><div class="val">$(Fmt0 $totalBaseExe)</div><div class="sub">Ingreso neto de impuestos</div></div>
<div class="kpi ok"><div class="lbl">Facturado y cobrado</div><div class="val">$(Fmt0 $montoCobrado)</div><div class="sub">$(Fmt1Pct $pctCobrado) de las ventas</div></div>
<div class="kpi warn"><div class="lbl">Facturado por cobrar</div><div class="val">$(Fmt0 $montoPorCobrar)</div><div class="sub">$(Fmt1Pct $pctPorCobrar) de las ventas</div></div>
<div class="kpi bad"><div class="lbl">Pendiente de facturar</div><div class="val">$(Fmt0 $montoSinFactura)</div><div class="sub">$(Fmt1Pct $pctSinFactura) de las ventas</div></div>
<div class="kpi"><div class="lbl">Impuestos gestionados</div><div class="val">$(Fmt0 $totalImpuestos)</div><div class="sub">IVA $(Fmt0 $totalIVA) $([char]0xB7) IGTF $(Fmt0 $totalIGTF)</div></div>
</div>
<h2>Resumen ejecutivo</h2>
<div class="note"><ul>
<li>Las ventas acumuladas del ejercicio $($FechaCorte.Year) alcanzan <b>USD $(Fmt0 $totalVentas)</b> distribuidas en <b>$docsTotal operaciones</b>, de las cuales <b>$docsFacturados</b> ya est${e_a}n facturadas (<b>$(Fmt0 $montoFacturado)</b>, $(Fmt1Pct $pctFacturado)).</li>
<li>La conversi${e_o}n a caja es el punto cr${e_i}tico: solo <b>$(Fmt1Pct $pctCobrado)</b> de la venta est$e_a efectivamente cobrada. Entre cuentas por cobrar ($(Fmt0 $montoPorCobrar)) y ventas sin facturar ($(Fmt0 $montoSinFactura)) hay <b>USD $(Fmt0 $montoPendienteCaja)</b> pendientes de convertir en efectivo ($(Fmt1Pct $pctPendienteCaja) del total).</li>
<li>Concentraci${e_o}n de cartera: <b>$($clienteTop.Cliente)</b> representa $(Fmt1Pct $clienteTopPct) de la facturaci${e_o}n y los tres primeros clientes concentran <b>$(Fmt1Pct $top3Pct)</b>.</li>
<li>El mes de mayor actividad fue <b>$($mesTop.Label)</b> con $(Fmt0 $mesTop.Monto) ($(Fmt1Pct $mesTopPct) del a${e_n}o).</li>
<li>La estructura tributaria muestra <b>$(Fmt0 $totalIVA)</b> de IVA y <b>$(Fmt0 $totalIGTF)</b> de IGTF; el rubro exento asciende a $(Fmt0 $totalExento) ($(Fmt1Pct (Pct $totalExento $totalVentas))).</li>
</ul></div>

<div id="tab-ventas" class="tab-panel active">
<h2>Evoluci${e_o}n mensual de las ventas</h2>
<table><thead><tr><th>Mes</th><th class=n>Docs.</th><th class=n>Monto USD</th><th class=n>% del a${e_n}o</th><th>Peso relativo</th></tr></thead>
<tbody>
$filasMensual
</tbody>
<tfoot><tr><td>TOTAL</td><td class=n>$docsTotal</td><td class=n>$(Fmt0 $totalVentas)</td><td class=n>100,0%</td><td></td></tr></tfoot></table>
<p style="font-size:12.5px;color:#6b7c8d">Los documentos a${e_u}n no facturados se asignan al mes de su orden de compra (PO), por no disponer de fecha de factura.</p>
<h2>Ventas por cliente</h2>
<table><thead><tr><th>Cliente</th><th class=n>Docs.</th><th class=n>Venta total</th><th class=n>% part.</th></tr></thead>
<tbody>
$filasClienteVentas
</tbody>
<tfoot><tr><td>TOTAL</td><td class=n>$docsTotal</td><td class=n>$(Fmt0 $totalVentas)</td><td class=n>100,0%</td></tr></tfoot></table>
<h2>Diez operaciones de mayor monto</h2>
<table><thead><tr><th>Documento</th><th>Cliente</th><th>Concepto</th><th class=n>Monto USD</th><th>Estado</th></tr></thead>
<tbody>
$filasTop10
</tbody></table>
</div>

<div id="tab-cobranza" class="tab-panel">
<h2>Situaci${e_o}n de facturaci${e_o}n y cobranza</h2>
<table><thead><tr><th>Estado</th><th class=n>Docs.</th><th class=n>Monto USD</th><th class=n>% del total</th></tr></thead>
<tbody>
<tr><td>Facturado y cobrado</td><td class=n>$($grpCobrado.Count)</td><td class=n>$(Fmt0 $montoCobrado)</td><td class=n>$(Fmt1Pct $pctCobrado)</td></tr>
<tr><td>Facturado sin cobrar</td><td class=n>$($grpPorCobrar.Count)</td><td class=n>$(Fmt0 $montoPorCobrar)</td><td class=n>$(Fmt1Pct $pctPorCobrar)</td></tr>
<tr><td>Sin facturar</td><td class=n>$($grpSinFactura.Count)</td><td class=n>$(Fmt0 $montoSinFactura)</td><td class=n>$(Fmt1Pct $pctSinFactura)</td></tr>
</tbody>
<tfoot><tr><td>TOTAL</td><td class=n>$docsTotal</td><td class=n>$(Fmt0 $totalVentas)</td><td class=n>100,0%</td></tr></tfoot></table>
<h2>Cartera pendiente por cliente</h2>
<table><thead><tr><th>Cliente</th><th class=n>Cobrado</th><th class=n>Por cobrar</th><th class=n>Sin facturar</th><th class=n>Total pendiente</th></tr></thead>
<tbody>
$filasCarteraCliente
</tbody>
<tfoot><tr><td>TOTAL</td><td class=n>$(Fmt0 $montoCobrado)</td><td class=n>$(Fmt0 $montoPorCobrar)</td><td class=n>$(Fmt0 $montoSinFactura)</td><td class=n>$(Fmt0 $montoPendienteCaja)</td></tr></tfoot></table>
<h2>Antig${e_ud}edad de las cuentas por cobrar</h2>
<table><thead><tr><th>Tramo (d${e_i}as desde la factura)</th><th class=n>Monto USD</th><th class=n>% de la cartera</th></tr></thead>
<tbody>
$filasAging
</tbody>
<tfoot><tr><td>TOTAL POR COBRAR</td><td class=n>$(Fmt0 $montoPorCobrar)</td><td class=n>100,0%</td></tr></tfoot></table>
</div>

$inproccaPanelHtml

$dispoPanelHtml

<h2>Conclusiones y recomendaciones</h2>
<div class="note"><ul>
<li><b>Acelerar la facturaci${e_o}n pendiente.</b> Existen $($grpSinFactura.Count) operaciones a${e_u}n sin facturar por USD $(Fmt0 $montoSinFactura). Emitir estos documentos es la acci${e_o}n de mayor impacto inmediato sobre el flujo de caja.</li>
<li><b>Gesti${e_o}n de cobranza focalizada.</b> Concentrar el esfuerzo en los saldos de mayor antig${e_ud}edad y en los clientes con mayor exposici${e_o}n por cobrar.</li>
<li><b>Diversificaci${e_o}n de cartera.</b> Con $(Fmt1Pct $top3Pct) de la venta en los tres primeros clientes, conviene desarrollar nuevos contratos para reducir la dependencia de pocos clientes.</li>
<li><b>Seguimiento por PO.</b> Varias ${e_o}rdenes se facturan en partes; un control por PO evitar${e_i}a saldos abiertos no facturados.</li>
</ul></div>
<footer>Informe generado autom${e_a}ticamente a partir de "$([System.IO.Path]::GetFileName($ExcelPath))" $([char]0x2014) hoja $SheetName. $([char]0xB7) $generadoTs</footer>
</div>
<script>
function actualizarDashboard(btn) {
  btn.disabled = true;
  var url = window.location.pathname + '?_=' + Date.now() + window.location.hash;
  window.location.replace(url);
}
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
