# Actualiza los valores de "Peru en cifras" en la portada
$ErrorActionPreference = "Stop"
$base = Split-Path -Parent $PSScriptRoot
$domain = $env:WP_DOMAIN
if (-not $domain) { $domain = ([IO.File]::ReadAllText((Join-Path $base "tools\domain.txt"))).Trim() }
$usuario = "hugogabrielac"
$rutaUsuario = Join-Path $base ".secrets\wordpress_usuario.txt"
if (Test-Path $rutaUsuario) {
  $u = ([IO.File]::ReadAllText($rutaUsuario)).Trim()
  if ($u) { $usuario = $u }
}
$pass = $env:WP_APP_PASSWORD
if (-not $pass) { $pass = ([IO.File]::ReadAllText((Join-Path $base ".secrets\wordpress_app_password.txt"))).Trim() }
$basico = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$usuario`:$pass"))
$headers = @{ "Authorization" = "Basic $basico"; "Content-Type" = "application/json; charset=utf-8"; "User-Agent" = "Actualidad Economica (h----------.com)" }
$cultura = [System.Globalization.CultureInfo]::GetCultureInfo("es-ES")

function Serie([string]$codigo, [int]$dias) {
  $hoy = Get-Date
  $ini = $hoy.AddDays(-$dias).ToString("yyyy-MM-dd")
  $fin = $hoy.ToString("yyyy-MM-dd")
  $r = Invoke-RestMethod -Uri ("https://estadisticas.bcrp.gob.pe/estadisticas/series/api/" + $codigo + "/json/" + $ini + "/" + $fin) -TimeoutSec 30
  $vals = @()
  foreach ($p in $r.periods) { if ($p.values[0] -and $p.values[0] -ne "n.d.") { $vals += ,@($p.name, [double]$p.values[0]) } }
  return $vals
}
function Ult($arr, $n) { return $arr[@($arr).Count - $n] }
function Fmt([double]$v, [int]$d) { return $v.ToString("N" + $d, $cultura) }
function Pill([double]$delta, [int]$dec, [string]$unidad, [bool]$bienArriba) {
  if ([Math]::Abs($delta) -lt [Math]::Pow(10, -$dec - 1)) { return '<span class="ae-pill flat">Sin cambios</span>' }
  $arriba = $delta -gt 0
  $clase = "down"
  if ($arriba -eq $bienArriba) { $clase = "up" }
  $flecha = if ($arriba) { [string][char]0x2191 } else { [string][char]0x2193 }
  $signo = if ($arriba) { "+" } else { "" }
  return '<span class="ae-pill ' + $clase + '">' + $flecha + ' ' + $signo + $delta.ToString("N" + $dec, $cultura) + ' ' + $unidad + '</span>'
}
function CambiarValor([string]$c, [string]$id, [string]$valor) {
  return [regex]::Replace($c, '(<div class="ae-mercado-dato[^"]*" id="ae-' + $id + '">)[^<]*(</div>)', ('${1}' + $valor + '${2}'))
}
function CambiarValorFiscal([string]$c, [double]$v) {
  $txt = (Fmt $v 1) + " %"
  $clase = "ae-mercado-dato"
  if ($v -lt 0) { $clase = "ae-mercado-dato neg" }
  return [regex]::Replace($c, '<div class="ae-mercado-dato[^"]*" id="ae-fiscal">[^<]*</div>', ('<div class="' + $clase + '" id="ae-fiscal">' + $txt + '</div>'))
}
function CambiarPill([string]$c, [string]$id, [string]$pill) {
  return [regex]::Replace($c, '(<span id="ae-' + $id + '-var">)[\s\S]*?(</span>)', ('${1}' + $pill + '${2}'))
}

$d = Serie "PD04638PD" 15
$e = Serie "PD04648PD" 15
$rin = Serie "PD04650MD" 20
$inf = Serie "PN01273PM" 220
$tasa = Serie "PD12301MD" 60
$rie = Serie "PD43663XD" 15
$pbi = Serie "PN02507AQ" 600
$des = Serie "PN38063GM" 220
$fis = Serie "PN39524FM" 220

$p = Invoke-RestMethod -Uri ("https://" + $domain + "/wp-json/wp/v2/pages/472?context=edit") -Headers $headers -TimeoutSec 60
$c = [string]$p.content.raw
if (-not $c.Contains('ae-peru-cifras')) { Write-Host 'el bloque Peru en cifras no existe'; exit 1 }

$c = CambiarValor $c "dolar" ("S/ " + (Fmt (Ult $d 1)[1] 3))
$c = CambiarPill $c "dolar" (Pill ((((Ult $d 1)[1] / (Ult $d 2)[1]) - 1) * 100) 2 "%" $false)
$c = CambiarValor $c "euro" ("S/ " + (Fmt (Ult $e 1)[1] 3))
$c = CambiarPill $c "euro" (Pill ((((Ult $e 1)[1] / (Ult $e 2)[1]) - 1) * 100) 2 "%" $false)
$c = CambiarValor $c "rin" ("US$ " + (Fmt (Ult $rin 1)[1] 0))
$c = CambiarPill $c "rin" (Pill ((((Ult $rin 1)[1] / (Ult $rin 2)[1]) - 1) * 100) 1 "%" $true)
$c = CambiarValor $c "inflacion" ((Fmt (Ult $inf 1)[1] 2) + " %")
$c = CambiarPill $c "inflacion" (Pill ((Ult $inf 1)[1] - (Ult $inf 2)[1]) 1 "pp" $false)
$c = CambiarValor $c "tasa" ((Fmt (Ult $tasa 1)[1] 2) + " %")
$c = CambiarPill $c "tasa" (Pill ((Ult $tasa 1)[1] - (Ult $tasa 2)[1]) 2 "pp" $false)
$c = CambiarValor $c "riesgo" ((Fmt (Ult $rie 1)[1] 0) + " pb")
$c = CambiarPill $c "riesgo" (Pill ((Ult $rie 1)[1] - (Ult $rie 2)[1]) 0 "pb" $false)
$c = CambiarValor $c "pbi" ((Fmt (Ult $pbi 1)[1] 1) + " %")
$c = CambiarPill $c "pbi" (Pill ((Ult $pbi 1)[1] - (Ult $pbi 2)[1]) 1 "pp" $true)
$c = CambiarValor $c "desempleo" ((Fmt (Ult $des 1)[1] 1) + " %")
$c = CambiarPill $c "desempleo" (Pill ((Ult $des 1)[1] - (Ult $des 2)[1]) 1 "pp" $false)
$c = CambiarValorFiscal $c (Ult $fis 1)[1]
$c = CambiarPill $c "fiscal" (Pill ((Ult $fis 1)[1] - (Ult $fis 2)[1]) 1 "pp" $true)

$meses = @("enero","febrero","marzo","abril","mayo","junio","julio","agosto","setiembre","octubre","noviembre","diciembre")
$hoy = Get-Date
$fecha = $hoy.Day.ToString() + " de " + $meses[$hoy.Month - 1] + " de " + $hoy.Year
$c = [regex]::Replace($c, '(<span class="fecha" id="ae-cifras-fecha">)[^<]*(</span>)', ('${1}' + $fecha + '${2}'))

$json = @{ content = $c } | ConvertTo-Json -Depth 5 -Compress
Invoke-RestMethod -Uri ("https://" + $domain + "/wp-json/wp/v2/pages/472") -Method Post -Headers $headers -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 90 | Out-Null
Write-Host 'Peru en cifras actualizado'
