param(
  [string]$Excel = "$PSScriptRoot\dados\relatorio_blip_consolidado.xlsx",
  [string]$Saida = "$PSScriptRoot\index.html",
  [string]$Senha = $env:BLIP_DASH_SENHA   # opcional: sem senha, o painel abre direto
)
# Le o relatorio consolidado da Blip e gera index.html (arquivo unico).
# Uso semanal: substitua dados\relatorio_blip_consolidado.xlsx e rode este script.
# Abas de Fatura / Resultados por acao: lidas da planilha principal se existirem;
# caso contrario, de dados\relatorio_blip_setembro_ceo.xlsx.
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Excel)) { throw "Excel nao encontrado: $Excel" }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$inv = [Globalization.CultureInfo]::InvariantCulture
$tmpDirs = @()

function Open-Book($path) {
  $d = Join-Path $env:TEMP ("blipx_" + [guid]::NewGuid().ToString('N'))
  [IO.Compression.ZipFile]::ExtractToDirectory($path, $d)
  $script:tmpDirs += $d
  $strs = @()
  if (Test-Path "$d\xl\sharedStrings.xml") {
    [xml]$ss = Get-Content "$d\xl\sharedStrings.xml" -Encoding UTF8
    $strs = @($ss.sst.si | ForEach-Object { $_.InnerText })
  }
  [xml]$wb = Get-Content "$d\xl\workbook.xml" -Encoding UTF8
  $names = @($wb.workbook.sheets.sheet | ForEach-Object { $_.name })
  return @{ dir = $d; strs = $strs; names = $names }
}
function Find-Sheet($book, $nome) {
  for ($k = 0; $k -lt $book.names.Count; $k++) { if ($book.names[$k] -like $nome) { return $k + 1 } }
  return 0
}
function Get-Sheet($book, $nome) {
  $i = Find-Sheet $book $nome
  if ($i -le 0) { throw "Aba '$nome' nao encontrada" }
  return Get-SheetIdx $book $i
}
function Get-SheetIdx($book, $i) {
  [xml]$x = Get-Content "$($book.dir)\xl\worksheets\sheet$i.xml" -Encoding UTF8
  $out = @{}
  foreach ($row in $x.worksheet.sheetData.row) {
    $h = @{}
    foreach ($c in $row.c) {
      if ($null -eq $c.v -or "$($c.v)" -eq '') { continue }
      $col = $c.r -replace '\d', ''
      if ($c.t -eq 's') { $h[$col] = $book.strs[[int]$c.v] }
      elseif ($c.t -eq 'str') { $h[$col] = [string]$c.v }
      else { $h[$col] = [double]::Parse($c.v, $inv) }
    }
    $out[[int]$row.r] = $h
  }
  return $out
}
function Val($v) { if ($v -is [double]) { $v } elseif ($null -ne $v) { [string]$v } else { $null } }
function R2($v) { [math]::Round([double]$v, 2) }

try {
  $bk = Open-Book $Excel

  # Diario: base (planilha principal) + atualizacoes do mes em dados\custo_blip_*.xlsx (as mais recentes prevalecem)
  function Read-Diario($sh) {
    $o = @()
    foreach ($r in ($sh.Keys | Sort-Object)) { if ($r -lt 5) { continue }; $h = $sh[$r]
      if ($h.A -is [double] -and $h.F -is [double]) {
        $o += [ordered]@{ d = $h.A; util = $h.B; auth = $h.C; mkt = $h.D; lite = $h.E; total = $h.F; custo = $h.G } } }
    return ,$o
  }
  $dMap = @{}
  foreach ($x in (Read-Diario (Get-Sheet $bk 'Di?rio'))) { $dMap[[int]$x.d] = $x }
  $maxAtual = ($dMap.Keys | Measure-Object -Maximum).Maximum
  $pAtual = Get-Sheet $bk 'Premissas'
  $upds = @()
  foreach ($f in (Get-ChildItem "$PSScriptRoot\dados" -Filter 'custo_blip_*.xlsx' -File -ErrorAction SilentlyContinue)) {
    $ub = Open-Book $f.FullName
    $rows = Read-Diario (Get-Sheet $ub 'Di?rio')
    if ($rows.Count) { $upds += @{ nome = $f.Name; book = $ub; rows = $rows; max = ($rows | ForEach-Object { $_.d } | Measure-Object -Maximum).Maximum } }
  }
  foreach ($u in ($upds | Sort-Object { $_.max })) {
    foreach ($x in $u.rows) { $dMap[[int]$x.d] = $x }
    if ($u.max -ge $maxAtual) { $maxAtual = $u.max; $pAtual = Get-Sheet $u.book 'Premissas'; $fonteAtual = $u.nome }
  }
  $diario = @(); foreach ($k in ($dMap.Keys | Sort-Object)) { $diario += $dMap[$k] }

  # Premissas (contatos, nota e tarifas da atualizacao mais recente)
  $p = $pAtual
  $tarifas = [ordered]@{ util = $p[4].C; lite = $p[5].C; mkt = $p[6].C; auth = $p[7].C }
  $semana1 = $p[9].C
  $contatos = @()
  foreach ($r in 14..16) { $h = $p[$r]; if ($h -and $h.A) {
    $contatos += [ordered]@{ nome = $h.A; disp = $h.B; pct = $h.C; util = $h.D; lite = $h.E; serv = $h.F } } }
  $notaContatos = [string]$p[19].A

  # Tarifas por mes e historico diario (jan/25 em diante) -> hist com custo por categoria
  $tm = Get-Sheet $bk 'Tarifas por m?s'
  $tabTar = @{}
  foreach ($r in $tm.Keys) { if ($r -lt 5) { continue }; $h = $tm[$r]
    if ($h.A -is [double]) { $dd = [DateTime]::FromOADate($h.A); $tabTar[$dd.ToString('yyyyMM')] = @($h.B, $h.D, $h.C, $h.E) } }  # util, mkt, lite, auth
  $hs = Get-Sheet $bk 'Hist?rico*'
  $dias = @{}
  foreach ($r in $hs.Keys) { if ($r -lt 5) { continue }; $h = $hs[$r]
    if ($h.A -is [double]) { $dias[[int]$h.A] = @($h.B, $h.C, $h.D, $h.E) } }   # util, auth, mkt, lite
  foreach ($x in $diario) { $dias[[int]$x.d] = @($x.util, $x.auth, $x.mkt, $x.lite) }
  $hist = New-Object System.Collections.ArrayList
  foreach ($k in ($dias.Keys | Sort-Object)) {
    $v = $dias[$k]; $dd = [DateTime]::FromOADate($k); $key = $dd.ToString('yyyyMM')
    if ($tabTar.ContainsKey($key)) { $t = $tabTar[$key] } else { $t = @($tarifas.util, $tarifas.mkt, $tarifas.lite, $tarifas.auth) }
    $cu = $v[0] * $t[0]; $cm = $v[2] * $t[1]; $cl = $v[3] * $t[2]; $ca = $v[1] * $t[3]
    [void]$hist.Add([double[]]@($k, $v[0], $v[1], $v[2], $v[3], (R2 $cu), (R2 $cl), (R2 $cm), (R2 $ca)))
  }

  # Consolidado (mensal + anual)
  $grp = @{}
  foreach ($r in $hist) { $k = [DateTime]::FromOADate($r[0]).ToString('yyyyMM'); if (-not $grp.ContainsKey($k)) { $grp[$k] = New-Object System.Collections.ArrayList }; [void]$grp[$k].Add($r) }
  $mensal = @()
  foreach ($k in ($grp.Keys | Sort-Object)) {
    $rows = $grp[$k]; $first = [DateTime]::ParseExact($k + '01', 'yyyyMMdd', $inv)
    $nd = [DateTime]::DaysInMonth($first.Year, $first.Month)
    $u = 0; $a = 0; $mk = 0; $l = 0; $cu = 0; $cl = 0; $cm = 0; $ca = 0; $pico = 0; $dp = $null
    foreach ($r in $rows) { $u += $r[1]; $a += $r[2]; $mk += $r[3]; $l += $r[4]; $cu += $r[5]; $cl += $r[6]; $cm += $r[7]; $ca += $r[8]
      $tot = $r[1] + $r[2] + $r[3] + $r[4]; if ($tot -gt $pico) { $pico = $tot; $dp = $r[0] } }
    $total = $u + $a + $mk + $l; $custo = $cu + $cl + $cm + $ca
    $mensal += [ordered]@{ m = $first.ToOADate(); dias = $rows.Count; util = $u; lite = $l; mkt = $mk; auth = $a; total = $total
      custoUtil = (R2 $cu); custoLite = (R2 $cl); custoMkt = (R2 $cm); custoAuth = (R2 $ca); custo = (R2 $custo)
      medio = $(if ($total) { $custo / $total } else { 0 }); pico = $pico; diaPico = $dp; completo = $(if ($rows.Count -eq $nd) { 1 } else { 0 }) }
  }
  # Comparativo anual (meses completos): ano anterior inteiro, ano atual ate o ultimo mes completo e mesmo periodo do ano anterior
  $anual = @()
  $compl = @($mensal | Where-Object { $_.completo -eq 1 })
  if ($compl.Count) {
    $ult = [DateTime]::FromOADate($compl[-1].m); $y = $ult.Year; $py = $y - 1
    $ab = @('jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez')[$ult.Month - 1]
    function Agg($ms) { $c = 0; $d = 0; $n = 0; foreach ($m in $ms) { $c += $m.custo; $d += $m.total; $n++ }
      return [ordered]@{ meses = $n; disp = $d; custo = (R2 $c); medio = $(if ($d) { $c / $d } else { 0 }); mediaMes = $(if ($n) { $c / $n } else { 0 }) } }
    $ano = @($compl | Where-Object { [DateTime]::FromOADate($_.m).Year -eq $y })
    $antI = @($compl | Where-Object { [DateTime]::FromOADate($_.m).Year -eq $py })
    $antP = @($antI | Where-Object { [DateTime]::FromOADate($_.m).Month -le $ult.Month })
    $rI = Agg $antI; $rA = Agg $ano; $rP = Agg $antP
    if ($antI.Count) { $anual += [ordered]@{ periodo = "Ano $py" + $(if ($antI.Count -eq 12) { ' completo' } else { '' }); meses = $rI.meses; disp = $rI.disp; custo = $rI.custo; medio = $rI.medio; mediaMes = $rI.mediaMes } }
    $anual += [ordered]@{ periodo = "Ano $y (jan a $ab)"; meses = $rA.meses; disp = $rA.disp; custo = $rA.custo; medio = $rA.medio; mediaMes = $rA.mediaMes }
    if ($antP.Count) {
      $anual += [ordered]@{ periodo = "Ano $py (jan a $ab)"; meses = $rP.meses; disp = $rP.disp; custo = $rP.custo; medio = $rP.medio; mediaMes = $rP.mediaMes }
      $vr = 'Varia' + [char]0x00E7 + [char]0x00E3 + 'o vs. mesmo per' + [char]0x00ED + 'odo do ano anterior'
      $anual += [ordered]@{ periodo = $vr; meses = $null; disp = ($rA.disp / $rP.disp - 1); custo = ($rA.custo / $rP.custo - 1); medio = ($rA.medio / $rP.medio - 1); mediaMes = ($rA.mediaMes / $rP.mediaMes - 1) }
    }
  }
  # Fatura, Resultados por acao e Resumo, uma aba de cada por mes (planilha principal e, se faltar mes, arquivo do CEO).
  # O mes vem do titulo (A1/A2) da aba, no formato "setembro/2026". Abas: Resultados*, Resumo*, Fatura*.
  $mesRx = [ordered]@{ 'janeiro'=1;'fevereiro'=2;'mar.o'=3;'abril'=4;'maio'=5;'junho'=6;'julho'=7;'agosto'=8;'setembro'=9;'outubro'=10;'novembro'=11;'dezembro'=12 }
  function Get-MesSerial($txt) {
    foreach ($k in $mesRx.Keys) { if ($txt -match "(?i)$k/(\d{4})") { return ([DateTime]::new([int]$matches[1], $mesRx[$k], 1)).ToOADate() } }
    return $null
  }
  $porMes = @{}
  function Get-Slot($m) { if (-not $porMes.ContainsKey($m)) { $script:porMes[$m] = @{ mes = $m } }; return $porMes[$m] }
  $fontes = @($bk)
  $ceo = "$PSScriptRoot\dados\relatorio_blip_setembro_ceo.xlsx"
  if (Test-Path $ceo) { $fontes += Open-Book $ceo }
  foreach ($b in $fontes) {
    for ($i = 1; $i -le $b.names.Count; $i++) {
      $nome = $b.names[$i - 1]
      if ($nome -notlike 'Resultados*' -and $nome -notlike 'Resumo*' -and $nome -notlike 'Fatura*') { continue }
      $s = Get-SheetIdx $b $i
      $m = Get-MesSerial ([string]$s[1].A + ' ' + [string]$s[2].A)
      if (-not $m) { continue }
      $slot = Get-Slot $m
      if ($nome -like 'Fatura*' -and -not $slot.fatura) {
        $slot.faturaNome = [string]$s[1].A; $slot.fatura = @()
        foreach ($r in ($s.Keys | Sort-Object)) { if ($r -lt 4) { continue }; $h = $s[$r]
          $slot.fatura += [ordered]@{ r = $r; A = $h.A; B = (Val $h.B); C = (Val $h.C); D = (Val $h.D); E = (Val $h.E) } }
      }
      elseif ($nome -like 'Resumo*' -and -not $slot.resumo) {
        $slot.titulo = @([string]$s[1].A, [string]$s[2].A); $slot.resumo = @()
        foreach ($r in ($s.Keys | Sort-Object)) { if ($r -lt 4) { continue }; $h = $s[$r]
          $slot.resumo += [ordered]@{ r = $r; A = $h.A; B = (Val $h.B); C = (Val $h.C) } }
      }
      elseif ($nome -like 'Resultados*' -and -not $slot.camps) {
        $slot.camps = @(); $slot.total = $null
        foreach ($r in ($s.Keys | Sort-Object)) { if ($r -lt 5) { continue }; $h = $s[$r]
          if (-not $h.A -or -not ($h.B -is [double])) { continue }
          $o = [ordered]@{ acao = $h.A; camps = $h.B; env = $h.C; ent = $h.D; custo = $h.E; receita = $h.F; roas = $h.G; situacao = [string]$h.H }
          if ($h.A -eq 'TOTAL') { $slot.total = $o } else { $slot.camps += $o } }
      }
    }
  }
  $campMeses = @(); $tarConf = @(); $fatura = @(); $faturaNome = ''
  foreach ($m in ($porMes.Keys | Sort-Object)) { $sl = $porMes[$m]
    if ($sl.fatura) { $tarConf += $m; $fatura = $sl.fatura; $faturaNome = $sl.faturaNome }
    if ($sl.camps) { $campMeses += [ordered]@{ mes = $m; titulo = $sl.titulo; campanhas = $sl.camps; total = $sl.total; resumo = $sl.resumo } } }

  $data = [ordered]@{
    gerado = (Get-Date -Format 'dd/MM/yyyy HH:mm'); arquivo = ((Split-Path $Excel -Leaf) + $(if ($fonteAtual) { ' + ' + $fonteAtual } else { '' }))
    tarifas = $tarifas; semana1 = $semana1; contatos = $contatos; notaContatos = $notaContatos
    diario = $diario; hist = $hist; mensal = $mensal; anual = $anual
    fatura = $fatura; faturaNome = $faturaNome; tarifasConfirmadas = $tarConf
    campMeses = $campMeses
  }
  $json = ConvertTo-Json -InputObject $data -Depth 6 -Compress
  $logo = 'data:image/png;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes("$PSScriptRoot\logo-b.png"))
  $tpl = Get-Content "$PSScriptRoot\dashboard_template.html" -Raw -Encoding UTF8
  if ($Senha) {
    # Com senha: dados criptografados (PBKDF2-SHA1 200k + AES-256-CBC); o HTML nao contem dados em texto claro.
    $salt = New-Object byte[] 16; $iv = New-Object byte[] 16
    $rng = [Security.Cryptography.RNGCryptoServiceProvider]::new(); $rng.GetBytes($salt); $rng.GetBytes($iv)
    $kdf = New-Object Security.Cryptography.Rfc2898DeriveBytes($Senha, $salt, 200000)
    $aes = [Security.Cryptography.Aes]::Create(); $aes.Mode = 'CBC'; $aes.Padding = 'PKCS7'; $aes.Key = $kdf.GetBytes(32); $aes.IV = $iv
    $plain = [Text.Encoding]::UTF8.GetBytes('BLIP1:' + $json)
    $ct = $aes.CreateEncryptor().TransformFinalBlock($plain, 0, $plain.Length)
    $b64 = [Convert]::ToBase64String([byte[]]($salt + $iv + $ct))
    $html = $tpl.Replace('/*__ENC__*/', $b64).Replace('/*__LOGO__*/', $logo)
  } else {
    # Sem senha: dados em texto claro, painel abre direto.
    $html = $tpl.Replace('/*__DATA__*/null', $json).Replace('/*__LOGO__*/', $logo)
  }
  [IO.File]::WriteAllText($Saida, $html, (New-Object Text.UTF8Encoding($true)))
  Write-Host "OK: $Saida  (dias historico=$($hist.Count), meses=$($mensal.Count), meses de campanhas=$($campMeses.Count), fatura linhas=$($fatura.Count))"
}
finally { foreach ($d in $tmpDirs) { Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue } }
