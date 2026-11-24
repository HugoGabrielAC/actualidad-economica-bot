# actualizar_cotiz_hora.ps1
# Actualiza SOLO la tira de cotizaciones (Dólar, Oro, Petróleo Brent, Hora internacional)
# y la fecha "Claves del día" de la portada (página WP 472), sin tocar noticias.
$ErrorActionPreference = "Stop"
$base = Split-Path -Parent $PSScriptRoot
$domain = $env:WP_DOMAIN
if (-not $domain) { $domain = ([IO.File]::ReadAllText((Join-Path $base "tools\domain.txt"))).Trim() }
$ua = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/131.0" }
$mesesEs = @("enero","febrero","marzo","abril","mayo","junio","julio","agosto","setiembre","octubre","noviembre","diciembre")
$cultura = [System.Globalization.CultureInfo]::GetCultureInfo("es-ES")
function FechaLarga([datetime]$d) { return $d.Day.ToString() + " de " + $mesesEs[$d.Month - 1] + " de " + $d.Year }
function Fmt([double]$n, [int]$d = 2) { return $n.ToString("N" + $d, $cultura) }
function Esc([string]$s) {
  if (-not $s) { return "" }
  return ($s -replace "&", "&amp;" -replace "<", "&lt;" -replace ">", "&gt;" -replace '"', "&quot;")
}
function Replace-Tarjeta([string]$html, [string]$h3Prefijo, [string]$nuevo) {
  $pos = 0
  while ($true) {
    $k = $html.IndexOf('<div class="ae-tarjeta">', $pos)
    if ($k -lt 0) { return $html }
    $h = $html.IndexOf($h3Prefijo, $k)
    if ($h -ge 0 -and ($h - $k) -lt 80) {
      $p = $k + '<div class="ae-tarjeta">'.Length
      $depth = 1
      while ($depth -gt 0 -and $p -lt $html.Length) {
        $o = $html.IndexOf('<div', $p)
        $cl = $html.IndexOf('</div>', $p)
        if ($cl -lt 0) { return $html }
        if ($o -ge 0 -and $o -lt $cl) { $depth++; $p = $o + 4 }
        else { $depth--; $p = $cl + 6 }
      }
      if ($depth -eq 0) { return $html.Substring(0, $k) + $nuevo + $html.Substring($p) }
      return $html
    }
    $pos = $k + 1
  }
}

# --- Dólar (PEN) ---
$dolar = $null
try {
  $div = Get-Content (Join-Path $base "web\divisas\datos\divisas.json") -Raw -Encoding UTF8 | ConvertFrom-Json
  $p = $div.principales | Where-Object { $_.c -eq "PEN" }
  if ($p) { $dolar = [double]$p.r }
} catch { }
if (-not $dolar) {
  try {
    $er = Invoke-RestMethod "https://open.er-api.com/v6/latest/USD" -TimeoutSec 30 -Headers $ua
    $dolar = [double]$er.rates.PEN
  } catch { }
}

# --- Oro y Petróleo (Yahoo futures, con respaldo del cache) ---
function Get-Futuro([string]$simbolo) {
  try {
    $u = "https://query1.finance.yahoo.com/v8/finance/chart/" + $simbolo + "?range=5d&interval=1d"
    $j = Invoke-RestMethod -Uri $u -TimeoutSec 20 -Headers $ua
    $r = $j.chart.result[0]
    $c = @($r.indicators.quote[0].close | Where-Object { $null -ne $_ })
    if ($c.Count -eq 0) { return $null }
    $v = [double]$c[-1]
    $prev = $null
    if ($c.Count -ge 2) { $prev = [double]$c[-2] }
    $var = $null
    if ($prev -and $prev -ne 0) { $var = ($v - $prev) / $prev * 100 }
    return [pscustomobject]@{ valor = $v; variacion = $var }
  } catch { return $null }
}
function Get-Cache([string]$clave) {
  try {
    $ruta = Join-Path $base "web\portada\datos\cotizaciones.json"
    if (Test-Path $ruta) {
      $prev = Get-Content $ruta -Raw -Encoding UTF8 | ConvertFrom-Json
      $x = $prev | Where-Object { $_.clave -eq $clave } | Select-Object -Last 1
      if ($x) { return [pscustomobject]@{ valor = [double]$x.valor; variacion = $x.variacion } }
    }
  } catch { }
  return $null
}
$oro = Get-Futuro "GC=F"
if (-not $oro) { $oro = Get-Cache "oro" }
$petroleo = Get-Futuro "CL=F"
if (-not $petroleo) { $petroleo = Get-Cache "petroleo" }

function VarTxt($q) {
  if (-not $q -or $null -eq $q.variacion) { return "" }
  $v = [double]$q.variacion
  $signo = "+"
  if ($v -lt 0) { $signo = "" }
  return " (" + $signo + (Fmt $v 1) + " %)"
}

# --- hora internacional ---
$ahora = Get-Date
$tzNy = [TimeZoneInfo]::FindSystemTimeZoneById("Eastern Standard Time")
$tzPk = [TimeZoneInfo]::FindSystemTimeZoneById("China Standard Time")
$tzMd = [TimeZoneInfo]::FindSystemTimeZoneById("Romance Standard Time")
$ny = [TimeZoneInfo]::ConvertTime($ahora, $tzNy).ToString("HH:mm")
$pk = [TimeZoneInfo]::ConvertTime($ahora, $tzPk).ToString("HH:mm")
$md = [TimeZoneInfo]::ConvertTime($ahora, $tzMd).ToString("HH:mm")

# --- strip nuevo ---
$datoDolar = "S/ " + $(if ($dolar) { Fmt $dolar 3 } else { "-" })
$datoOro = "US$ " + $(if ($oro) { Fmt $oro.valor 2 + (VarTxt $oro) } else { "-" })
$datoPet = "US$ " + $(if ($petroleo) { Fmt $petroleo.valor 2 + (VarTxt $petroleo) } else { "-" })
$strip = '<div class="ae-cotiza">' + "`n" +
  '<div><div class="etiqueta">D' + [char]0x00F3 + 'lar</div><div class="dato">' + $datoDolar + '</div><div class="nota">Por USD</div><p style="margin:8px 0 0;"><a class="ae-boton" href="https://' + $domain + '/conversor-de-divisas/">Convertir</a></p></div>' + "`n" +
  '<div><div class="etiqueta">Oro</div><div class="dato">' + $datoOro + '</div><div class="nota">Onza troy</div></div>' + "`n" +
  '<div><div class="etiqueta">Petr' + [char]0x00F3 + 'leo Brent</div><div class="dato">' + $datoPet + '</div><div class="nota">Barril</div></div>' + "`n" +
  '<div><div class="etiqueta">Hora internacional</div><div style="font-size:15px;line-height:1.5;margin-top:4px;">Nueva York ' + $ny + '<br />Pekin ' + $pk + '<br />Madrid ' + $md + '</div><div class="nota">Hora de referencia</div></div>' + "`n" +
  '</div>'

Write-Host ("Dolar: " + $datoDolar + " | Oro: " + $datoOro + " | Petro: " + $datoPet)
Write-Host ("Hora: NY " + $ny + " / Pekin " + $pk + " / Madrid " + $md)

# --- WP: reemplazar el strip y la fecha ---
$usuario = $env:WP_USER
if (-not $usuario) { $usuario = "hugogabrielac" }
if (Test-Path (Join-Path $base ".secrets\wordpress_usuario.txt")) {
  $u = ([IO.File]::ReadAllText((Join-Path $base ".secrets\wordpress_usuario.txt"))).Trim()
  if ($u) { $usuario = $u }
}
$pass = $env:WP_APP_PASSWORD
if (-not $pass) {
  $pass = ([IO.File]::ReadAllText((Join-Path $base ".secrets\wordpress_app_password.txt"))).Trim()
}
$basico = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$usuario`:$pass"))
$headers = @{ "Authorization" = "Basic $basico"; "Content-Type" = "application/json; charset=utf-8"; "User-Agent" = "Actualidad Economica (h----------.com)" }
$re = Invoke-RestMethod -Uri ("https://" + $domain + "/wp-json/wp/v2/pages/472?context=edit") -Headers $headers -TimeoutSec 60
$c = [string]$re.content.rendered

$i = $c.IndexOf('<div class="ae-cotiza">')
if ($i -ge 0) {
  $pos = $i + '<div class="ae-cotiza">'.Length
  $depth = 1
  while ($depth -gt 0 -and $pos -lt $c.Length) {
    $o = $c.IndexOf('<div', $pos)
    $cl = $c.IndexOf('</div>', $pos)
    if ($cl -lt 0) { break }
    if ($o -ge 0 -and $o -lt $cl) { $depth++; $pos = $o + 4 }
    else { $depth--; $pos = $cl + 6 }
  }
  if ($depth -eq 0) {
    $c = $c.Substring(0, $i) + $strip + $c.Substring($pos)
    Write-Host "strip de cotizaciones reemplazado"
  } else { Write-Host "no se pudo delimitar la tira" }
} else { Write-Host "no encontré ae-cotiza" }

$fechaLarga = FechaLarga $ahora
$j = $c.IndexOf('<span class="fecha">')
if ($j -ge 0) {
  $s = $j + '<span class="fecha">'.Length
  $e = $c.IndexOf('</span>', $s)
  $c = $c.Substring(0, $s) + $fechaLarga + $c.Substring($e)
  Write-Host ("fecha actualizada a " + $fechaLarga)
}

# --- pregunta del dia: una sola, rotando por dia del ano (banco amplio) ---
$banca = @()
try {
  $lineas = Get-Content (Join-Path $base "web\bancos\preguntas.txt") -Encoding UTF8
  foreach ($ln in $lineas) {
    if (-not $ln.Trim()) { continue }
    $q = $ln.Split("|")
    if ($q.Count -ge 7) { $banca += ,$q }
  }
} catch { }
if ($banca.Count -gt 0) {
  $dia = (Get-Date).DayOfYear
  $q = $banca[($dia - 1) % $banca.Count]
  $idB = "ae-preg-1"
  $textoExplic = $q[6].Replace("'", " ")
  $js = @'
document.querySelectorAll('.ae-opcion').forEach(function(b){b.addEventListener('click',function(){var bloque=b.getAttribute('data-bloque');var txt=document.getElementById(bloque);document.querySelectorAll('.ae-opcion[data-bloque=' + bloque + ']').forEach(function(x){x.style.background='';x.style.color=''});if(bloque==='__ID__'){if(parseInt(b.getAttribute('data-respuesta'),10)===__C__){b.style.background='#66C96D';b.style.color='#FFFFFF';txt.textContent='__T__';}else{txt.textContent='No es esa. Vuelve a intentarlo.';}}})});
'@
  $js = $js.Replace("__ID__", $idB).Replace("__C__", [string][int]$q[5]).Replace("__T__", $textoExplic)
  $card = '<div class="ae-tarjeta"><h3>Pregunta del d' + [char]0x00ED + 'a</h3>' + "`n" +
    '<p style="font-family:Georgia,''Times New Roman'',serif;font-size:18px;margin:0 0 10px;"><span style="color:#66C96D;font-size:12px;letter-spacing:0.12em;text-transform:uppercase;display:block;margin-bottom:5px;">Pregunta 1</span>' + $q[1] + '</p>' + "`n" +
    '<p><button class="ae-opcion" data-bloque="' + $idB + '" data-respuesta="0">A. ' + $q[2] + '</button><br />' + "`n" +
    '<button class="ae-opcion" data-bloque="' + $idB + '" data-respuesta="1">B. ' + $q[3] + '</button><br />' + "`n" +
    '<button class="ae-opcion" data-bloque="' + $idB + '" data-respuesta="2">C. ' + $q[4] + '</button></p>' + "`n" +
    '<p id="' + $idB + '" style="font-size:14px;color:#555555;font-style:italic;margin:8px 0 0;">Elige una opci' + [char]0x00F3 + 'n y descubre la explicaci' + [char]0x00F3 + 'n.</p>' + "`n" +
    '<p><script>' + $js + '</script></p>' + "`n" +
    '</div>'
  $k = $c.IndexOf('<div class="ae-tarjeta">')
  $found = -1
  while ($k -ge 0) {
    $next = $c.IndexOf('<h3>Pregunta del d', $k)
    if ($next -ge 0 -and ($next - $k) -lt 60) { $found = $k; break }
    $k = $c.IndexOf('<div class="ae-tarjeta">', $k + 1)
  }
  if ($found -ge 0) {
    $close = $c.IndexOf('</div>', $found)
    if ($close -gt $found) {
      $c = $c.Substring(0, $found) + $card + $c.Substring($close + 6)
      Write-Host ("pregunta del dia actualizada: " + $q[1].Substring(0, [Math]::Min(50, $q[1].Length)))
    }
  }
}

# --- frase del dia ---
$frases = @()
try {
  foreach ($ln in (Get-Content (Join-Path $base "web\bancos\frases.txt") -Encoding UTF8)) {
    if (-not $ln.Trim()) { continue }
    $x = $ln.Split("|")
    if ($x.Count -ge 4) { $frases += ,$x }
  }
} catch { }
if ($frases.Count -gt 0) {
  $fr = $frases[((Get-Date).DayOfYear - 1) % $frases.Count]
  $cardFrase = '<div class="ae-tarjeta"><h3>Frase del d' + [char]0x00ED + 'a</h3>' + "`n" +
    '<div style="font-family:Georgia,''Times New Roman'',serif;font-style:italic;font-size:22px;line-height:1.35;margin:2px 0 4px;">' + (Esc $fr[0]) + '</div>' + "`n" +
    '<div style="font-weight:700;margin-top:10px;font-size:16px;">' + (Esc $fr[1]) + '</div>' + "`n" +
    '<div style="font-size:13px;color:#777777;font-style:italic;margin:2px 0 14px;">' + (Esc $fr[3]) + '</div>' + "`n" +
    '<div style="font-size:15px;color:#333333;line-height:1.6;border-top:1px solid #E3E3E3;padding-top:14px;">' + (Esc $fr[2]) + '</div>' + "`n" +
    '</div>'
  $c = Replace-Tarjeta $c '<h3>Frase del d' $cardFrase
  Write-Host ("frase actualizada: " + $fr[1])
}

# --- personaje del dia (Wikipedia con cache) ---
$personajes = @()
try {
  foreach ($ln in (Get-Content (Join-Path $base "web\bancos\personajes.txt") -Encoding UTF8)) {
    if (-not $ln.Trim()) { continue }
    $x = $ln.Split("|")
    if ($x.Count -ge 3) { $personajes += ,$x }
  }
} catch { }
if ($personajes.Count -gt 0) {
  $diaP = (Get-Date).DayOfYear
  $pj = $personajes[($diaP - 1) % $personajes.Count]
  $cacheRuta = Join-Path $base "web\portada\datos\personajes.json"
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $cacheRuta) | Out-Null
  $cacheP = @{}
  if (Test-Path $cacheRuta) {
    try {
      $prev = Get-Content $cacheRuta -Raw -Encoding UTF8 | ConvertFrom-Json
      foreach ($x in $prev) { $cacheP[[string]$x.titulo] = $x }
    } catch { }
  }
  $tituloP = [string]$pj[0]
  $infoP = $null
  if ($cacheP.ContainsKey($tituloP)) { $infoP = $cacheP[$tituloP] }
  else {
    try {
      $urlW = "https://es.wikipedia.org/api/rest_v1/page/summary/" + [uri]::EscapeDataString($tituloP)
      $jW = Invoke-RestMethod -Uri $urlW -TimeoutSec 25 -Headers @{ "User-Agent" = "ActualidadEconomica/1.0 (h----------.com)" }
      $infoP = [pscustomobject]@{
        titulo = $tituloP
        nombre = $(if ($jW.title) { $jW.title } else { $tituloP })
        extracto = $(if ($jW.extract) { $jW.extract } else { $pj[2] })
        imagen = $(if ($jW.thumbnail) { $jW.thumbnail.source } else { $null })
        url = $(if ($jW.content_urls) { $jW.content_urls.desktop.page } else { "https://es.wikipedia.org/wiki/" + [uri]::EscapeDataString($tituloP) })
      }
      $cacheP[$tituloP] = $infoP
      @($cacheP.Keys | ForEach-Object { $cacheP[$_] }) | ConvertTo-Json -Depth 4 | Set-Content $cacheRuta -Encoding UTF8
    } catch {
      $infoP = [pscustomobject]@{ titulo = $tituloP; nombre = $tituloP; extracto = $pj[2]; imagen = $null; url = "https://es.wikipedia.org/wiki/" + [uri]::EscapeDataString($tituloP) }
    }
  }
  $cardPj = '<div class="ae-tarjeta"><h3>Personaje del d' + [char]0x00ED + 'a</h3>' + "`n"
  if ($infoP.imagen) { $cardPj += '<img src="' + (Esc $infoP.imagen) + '" alt="' + (Esc $infoP.nombre) + '" style="width:100%;max-width:260px;float:left;margin:0 14px 8px 0;border-radius:2px;" />' + "`n" }
  $cardPj += '<div style="font-family:Georgia,''Times New Roman'',serif;font-size:20px;font-weight:700;line-height:1.25;">' + (Esc $infoP.nombre) + '</div>' + "`n" +
    '<div style="font-style:italic;color:#555555;font-size:15px;margin:2px 0 8px;">' + (Esc $pj[1]) + '</div>' + "`n" +
    '<p style="font-size:14px;color:#333333;margin:0 0 10px;">' + (Esc $pj[2]) + '</p>' + "`n" +
    '<a href="' + (Esc $infoP.url) + '" target="_blank" rel="noopener" style="color:#66C96D;font-weight:600;">Conoce su historia ' + [char]0x2192 + '</a>' + "`n" +
    '</div>'
  $c = Replace-Tarjeta $c '<h3>Personaje del d' $cardPj
  Write-Host ("personaje actualizado: " + $infoP.nombre)
}

$cuerpo = @{ content = $c; status = "publish" }
$json = $cuerpo | ConvertTo-Json -Depth 5 -Compress
$resp = Invoke-RestMethod -Uri ("https://" + $domain + "/wp-json/wp/v2/pages/472") -Method Post -Headers $headers -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 90
Write-Host ("Publicado: {0}" -f $resp.link)
