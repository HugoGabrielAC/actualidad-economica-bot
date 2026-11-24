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
    if ($h -ge 0 -and ($h - $k) -lt 400) {
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
$dolarPrev = $null
try {
  $div = Get-Content (Join-Path $base "web\divisas\datos\divisas.json") -Raw -Encoding UTF8 | ConvertFrom-Json
  $p = $div.principales | Where-Object { $_.c -eq "PEN" }
  if ($p) { $dolar = [double](@($p.r)[0]); if (@($p.p).Count -gt 0) { $dolarPrev = [double](@($p.p)[0]) } }
} catch { }
if (-not $dolar) {
  try {
    $er = Invoke-RestMethod "https://open.er-api.com/v6/latest/USD" -TimeoutSec 30 -Headers $ua
    $dolar = [double]$er.rates.PEN
  } catch { }
}
$varDolar = $null
if ($dolar -and $dolarPrev -and $dolarPrev -ne 0) { $varDolar = ($dolar - $dolarPrev) / $dolarPrev * 100 }
if ($null -eq $varDolar) {
  $estadoRuta = Join-Path $base "web\portada\datos\divisa_dia.json"
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $estadoRuta) | Out-Null
  $hoyIso = (Get-Date).ToString("yyyy-MM-dd")
  try {
    if (Test-Path $estadoRuta) {
      $est = Get-Content $estadoRuta -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($est.fecha -ne $hoyIso -and $est.PEN) {
        $dolarPrev = [double]$est.PEN
        if ($dolar -and $dolarPrev -ne 0) { $varDolar = ($dolar - $dolarPrev) / $dolarPrev * 100 }
      }
    }
  } catch { }
  if ($dolar) {
    $nuevo = @{ fecha = $hoyIso; PEN = $dolar } | ConvertTo-Json
    Set-Content $estadoRuta -Value $nuevo -Encoding UTF8
  }
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

# --- strip nuevo (Mercados de hoy: tarjetas con icono y variación) ---
$datoDolar = "S/ " + $(if ($dolar) { Fmt $dolar 3 } else { "-" })
$datoOro = "US$ " + $(if ($oro) { Fmt $oro.valor 2 } else { "-" })
$datoPet = "US$ " + $(if ($petroleo) { Fmt $petroleo.valor 2 } else { "-" })

function Pill($q) {
  if (-not $q -or $null -eq $q.variacion) { return '' }
  $v = [double]$q.variacion
  $flecha = [char]0x2193
  if ($v -ge 0) { $flecha = [char]0x2191 }
  $signo = ""
  if ($v -ge 0) { $signo = "+" }
  return '<span class="ae-pill">' + $flecha + ' ' + $signo + (Fmt $v 2) + '%</span>'
}
$pillDolar = ''
if ($null -ne $varDolar) {
  $pq = [pscustomobject]@{ variacion = [double]$varDolar }
  $pillDolar = Pill $pq
}
$icoDolar = '<svg class="ae-ico" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="#1B5E20" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 17 9 11 13 15 21 7"/><polyline points="21 7 15 7 21 13"/></svg>'
$icoOro = '<svg class="ae-ico" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="#1B5E20" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18"/><path d="M5 21l5-12h4l5 12"/><path d="M9 9h6"/></svg>'
$icoPet = '<svg class="ae-ico" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="#1B5E20" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="7" y="4" width="10" height="16" rx="2"/><path d="M7 9h10"/><path d="M7 15h10"/></svg>'
$icoHora = '<svg class="ae-ico" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="#1B5E20" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><polyline points="12 7 12 12 15 14"/></svg>'
$strip = '<div class="ae-cotiza">' + "`n" +
  '<div class="ae-mercado"><div class="ae-mercado-cab">' + $icoDolar + '<span class="ae-mercado-nombre">D' + [char]0x00F3 + 'lar</span></div><div class="ae-mercado-dato">' + $datoDolar + '</div><div class="ae-mercado-pie"><span class="ae-mercado-nota">Por USD</span>' + $pillDolar + '</div></div>' + "`n" +
  '<div class="ae-mercado"><div class="ae-mercado-cab">' + $icoOro + '<span class="ae-mercado-nombre">Oro</span></div><div class="ae-mercado-dato">' + $datoOro + '</div><div class="ae-mercado-pie"><span class="ae-mercado-nota">Onza troy</span>' + (Pill $oro) + '</div></div>' + "`n" +
  '<div class="ae-mercado"><div class="ae-mercado-cab">' + $icoPet + '<span class="ae-mercado-nombre">Petr' + [char]0x00F3 + 'leo Brent</span></div><div class="ae-mercado-dato">' + $datoPet + '</div><div class="ae-mercado-pie"><span class="ae-mercado-nota">Barril</span>' + (Pill $petroleo) + '</div></div>' + "`n" +
  '<div class="ae-mercado"><div class="ae-mercado-cab">' + $icoHora + '<span class="ae-mercado-nombre">Hora internacional</span></div><div class="ae-horas"><span>Nueva York</span><b>' + $ny + '</b><span>Pek' + [char]0x00ED + 'n</span><b>' + $pk + '</b><span>Madrid</span><b>' + $md + '</b></div></div>' + "`n" +
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
$icoPreg = '<svg class="ae-ico" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#1B5E20" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M9.2 9.2a2.8 2.8 0 1 1 4.6 2.3c-.8.6-1.2 1.1-1.2 2.1"/><circle cx="12" cy="17.2" r=".4" fill="#1B5E20"/></svg>'
$icoPj = '<svg class="ae-ico" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#1B5E20" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 3.6-6.5 8-6.5s8 2.5 8 6.5"/></svg>'
$icoFrase = '<svg class="ae-ico" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#1B5E20" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 7H5a1 1 0 0 0-1 1v4a1 1 0 0 0 1 1h3v2a3 3 0 0 1-3 3"/><path d="M20 7h-5a1 1 0 0 0-1 1v4a1 1 0 0 0 1 1h3v2a3 3 0 0 1-3 3"/></svg>'
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
  $card = '<div class="ae-tarjeta"><h3>' + $icoPreg + 'Pregunta del d' + [char]0x00ED + 'a</h3>' + "`n" +
    '<p style="font-family:Georgia,''Times New Roman'',serif;font-size:18px;margin:0 0 10px;">' + $q[1] + '</p>' + "`n" +
    '<p><button class="ae-opcion" data-bloque="' + $idB + '" data-respuesta="0">A. ' + $q[2] + '</button><br />' + "`n" +
    '<button class="ae-opcion" data-bloque="' + $idB + '" data-respuesta="1">B. ' + $q[3] + '</button><br />' + "`n" +
    '<button class="ae-opcion" data-bloque="' + $idB + '" data-respuesta="2">C. ' + $q[4] + '</button></p>' + "`n" +
    '<p id="' + $idB + '" style="font-size:14px;color:#555555;font-style:italic;margin:8px 0 0;">Elige una opci' + [char]0x00F3 + 'n y descubre la explicaci' + [char]0x00F3 + 'n.</p>' + "`n" +
    '<p><script>' + $js + '</script></p>' + "`n" +
    '</div>'
  $k = $c.IndexOf('<div class="ae-tarjeta">')
  $found = -1
  while ($k -ge 0) {
    $next = $c.IndexOf('Pregunta del d', $k)
    if ($next -ge 0 -and ($next - $k) -lt 400) { $found = $k; break }
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
  $cardFrase = '<div class="ae-tarjeta"><h3>' + $icoFrase + 'Frase del d' + [char]0x00ED + 'a</h3>' + "`n" +
    '<div style="font-family:Georgia,''Times New Roman'',serif;font-style:italic;font-size:22px;line-height:1.35;margin:2px 0 4px;">' + (Esc $fr[0]) + '</div>' + "`n" +
    '<div style="font-weight:700;margin-top:10px;font-size:16px;">' + (Esc $fr[1]) + '</div>' + "`n" +
    '<div style="font-size:13px;color:#777777;font-style:italic;margin:2px 0 14px;">' + (Esc $fr[3]) + '</div>' + "`n" +
    '<div style="font-size:15px;color:#333333;line-height:1.6;border-top:1px solid #E3E3E3;padding-top:14px;">' + (Esc $fr[2]) + '</div>' + "`n" +
    '</div>'
  $c = Replace-Tarjeta $c 'Frase del d' $cardFrase
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
  $cardPj = '<div class="ae-tarjeta"><h3>' + $icoPj + 'Personaje del d' + [char]0x00ED + 'a</h3>' + "`n"
  if ($infoP.imagen) { $cardPj += '<img src="' + (Esc $infoP.imagen) + '" alt="' + (Esc $infoP.nombre) + '" style="width:100%;max-width:260px;float:left;margin:0 14px 8px 0;border-radius:2px;" />' + "`n" }
  $cardPj += '<div style="font-family:Georgia,''Times New Roman'',serif;font-size:20px;font-weight:700;line-height:1.25;">' + (Esc $infoP.nombre) + '</div>' + "`n" +
    '<div style="font-style:italic;color:#555555;font-size:15px;margin:2px 0 8px;">' + (Esc $pj[1]) + '</div>' + "`n" +
    '<p style="font-size:14px;color:#333333;margin:0 0 10px;">' + (Esc $pj[2]) + '</p>' + "`n" +
    '<a href="' + (Esc $infoP.url) + '" target="_blank" rel="noopener" style="color:#66C96D;font-weight:600;">Conoce su historia ' + [char]0x2192 + '</a>' + "`n" +
    '</div>'
  $c = Replace-Tarjeta $c 'Personaje del d' $cardPj
  Write-Host ("personaje actualizado: " + $infoP.nombre)
}

# --- ultimos videos (YouTube: 3 ultimos, con reproduccion en la pagina) ---
$cid = "UC__797T08iUoP9xHyh3c7bQ"
$vids = @()
try {
  $rss = (Invoke-WebRequest -Uri ("https://www.youtube.com/feeds/videos.xml?channel_id=" + $cid) -Headers $ua -TimeoutSec 60).Content
  $entries = [regex]::Matches($rss, '<entry>([\s\S]*?)</entry>')
  $vi = 0
  foreach ($e in $entries) {
    if ($vi -ge 3) { break }
    $b = $e.Groups[1].Value
    $vid = [regex]::Match($b, '<yt:videoId>([^<]+)<').Groups[1].Value
    $vt = [System.Net.WebUtility]::HtmlDecode([regex]::Match($b, '<media:title>([^<]+)<').Groups[1].Value)
    $dur = ''
    try {
      $vp = (Invoke-WebRequest -Uri ("https://www.youtube.com/watch?v=" + $vid) -Headers $ua -TimeoutSec 60).Content
      $dm = [regex]::Match($vp, '"lengthSeconds":"(\d+)"')
      if ($dm.Success) {
        $s = [int]$dm.Groups[1].Value
        $dur = [math]::Floor($s / 60).ToString() + ':' + ($s % 60).ToString('00')
      }
    } catch { }
    $vids += ,@{ id = $vid; titulo = $vt; dur = $dur }
    $vi++
  }
} catch { }
if ($vids.Count -gt 0) {
  $cards = ''
  foreach ($v in $vids) {
    $cards += '<div class="ae-video" data-id="' + $v.id + '"><div class="ae-video-thumb"><img src="https://i.ytimg.com/vi/' + $v.id + '/hqdefault.jpg" alt="" loading="lazy" /><span class="ae-video-play"><span>' + [char]0x25B6 + '</span></span>' + $(if ($v.dur) { '<span class="ae-video-dur">' + $v.dur + '</span>' } else { '' }) + '</div><div class="ae-video-title">' + (Esc $v.titulo) + '</div></div>' + "`n"
  }
  $videosHtml = '<div class="ae-videos">' + "`n" + '<h2 class="ae-titulo-seccion">' + [char]0x00DA + 'ltimos videos</h2>' + "`n" + '<div class="ae-videos-grid">' + "`n" + $cards + '</div>' + "`n" + '<script>document.querySelectorAll(''.ae-video'').forEach(function(v){v.addEventListener(''click'',function(){var id=v.getAttribute(''data-id'');v.innerHTML=''<iframe src="https://www.youtube-nocookie.com/embed/''+id+''?autoplay=1&rel=0" allow="autoplay; encrypted-media; picture-in-picture; fullscreen" allowfullscreen style="width:100%;aspect-ratio:16/9;border:0;border-radius:12px;"></iframe>'';});});</script>' + "`n" + '</div>'
  $iv = $c.IndexOf('<div class="ae-videos">')
  if ($iv -ge 0) {
    $pos = $iv + '<div class="ae-videos">'.Length
    $depth = 1
    while ($depth -gt 0 -and $pos -lt $c.Length) {
      $o = $c.IndexOf('<div', $pos)
      $cl = $c.IndexOf('</div>', $pos)
      if ($cl -lt 0) { break }
      if ($o -ge 0 -and $o -lt $cl) { $depth++; $pos = $o + 4 }
      else { $depth--; $pos = $cl + 6 }
    }
    if ($depth -eq 0) { $c = $c.Substring(0, $iv) + $videosHtml + $c.Substring($pos) }
  } else {
    $iua = $c.IndexOf('ltimo an')
    if ($iua -ge 0) {
      $h2start = $c.LastIndexOf('<h2', $iua)
      $c = $c.Substring(0, $h2start) + $videosHtml + $c.Substring($h2start)
    }
  }
  Write-Host ("videos actualizados: " + $vids.Count)
}

$cuerpo = @{ content = $c; status = "publish" }
$json = $cuerpo | ConvertTo-Json -Depth 5 -Compress
$resp = Invoke-RestMethod -Uri ("https://" + $domain + "/wp-json/wp/v2/pages/472") -Method Post -Headers $headers -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 90
Write-Host ("Publicado: {0}" -f $resp.link)
