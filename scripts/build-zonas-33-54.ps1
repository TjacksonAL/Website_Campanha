param(
  [Parameter(Mandatory=$true)][string]$MapeamentoXlsx,
  [Parameter(Mandatory=$true)][string]$CandCsv,
  [Parameter(Mandatory=$true)][string]$SecaoCsv,
  [Parameter(Mandatory=$true)][string]$OutJson
)

Add-Type -AssemblyName System.IO.Compression.FileSystem
$enc1252 = [System.Text.Encoding]::GetEncoding(1252)

# ---------- 1. Ler mapeamento zona 33/54 -> bairro ----------
function Get-EntryText($zip, $name) { $e=$zip.GetEntry($name); $sr=New-Object System.IO.StreamReader($e.Open(),[System.Text.Encoding]::UTF8); $t=$sr.ReadToEnd(); $sr.Close(); return $t }
function Col-Index([string]$r){ $l=($r -replace '[0-9]',''); $i=0; foreach($ch in $l.ToCharArray()){$i=$i*26+([int][char]$ch-[int][char]'A'+1)}; return $i }

$zip = [System.IO.Compression.ZipFile]::OpenRead($MapeamentoXlsx)
$sst=[xml](Get-EntryText $zip "xl/sharedStrings.xml")
$nsS=New-Object System.Xml.XmlNamespaceManager($sst.NameTable); $nsS.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
$si=$sst.SelectNodes("//m:sst/m:si",$nsS)
$shared=New-Object string[] $si.Count
for($i=0;$i -lt $si.Count;$i++){$shared[$i]=$si[$i].InnerText}
function Get-Rows($zip, $shared, $f){
  $d=[xml](Get-EntryText $zip $f)
  $ns=New-Object System.Xml.XmlNamespaceManager($d.NameTable)
  $ns.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  $out=New-Object System.Collections.Generic.List[object]
  foreach($rn in $d.SelectNodes("//m:sheetData/m:row",$ns)){
    $cells=@{}
    foreach($c in $rn.SelectNodes("m:c",$ns)){
      $ci=Col-Index $c.GetAttribute("r"); $t=$c.GetAttribute("t"); $v=$c.SelectSingleNode("m:v",$ns)
      $val=""; if($v){$raw=$v.InnerText; if($t -eq "s"){$val=$shared[[int]$raw]}else{$val=$raw}}
      $cells[$ci]=$val
    }
    $row = New-Object string[] 5
    for($k=1;$k -le 5;$k++){ $row[$k-1] = $cells[$k] }
    $out.Add($row)
  }
  return $out
}
$z33 = Get-Rows $zip $shared "xl/worksheets/sheet2.xml"
$z54 = Get-Rows $zip $shared "xl/worksheets/sheet3.xml"
$zip.Dispose()

$secaoToBairro = @{}
foreach ($r in $z33) { if ($r[1] -match '^\d{4}$') { $key = "33|" + [int]$r[1]; $secaoToBairro[$key] = $r[4] } }
foreach ($r in $z54) { if ($r[1] -match '^\d{4}$') { $key = "54|" + [int]$r[1]; $secaoToBairro[$key] = $r[4] } }
Write-Host "Mapeamento secao->bairro carregado: $($secaoToBairro.Count) secoes (zonas 33+54)"

# ---------- 2. Ler candidatos (consulta_cand) ----------
$splitRegex = New-Object System.Text.RegularExpressions.Regex ';(?=(?:[^"]*"[^"]*")*[^"]*$)'
function Split-CsvLine($line) {
  $raw = $splitRegex.Split($line)
  $out = New-Object string[] $raw.Count
  for ($i=0; $i -lt $raw.Count; $i++) {
    $v = $raw[$i]
    if ($v.Length -ge 2 -and $v[0] -eq '"' -and $v[$v.Length-1] -eq '"') { $v = $v.Substring(1, $v.Length-2) }
    $out[$i] = $v
  }
  return $out
}
function Read-Rows($path) {
  $reader = New-Object System.IO.StreamReader($path, $enc1252)
  $header = $reader.ReadLine()
  $cols = Split-CsvLine $header
  $idx = @{}
  for ($i=0; $i -lt $cols.Count; $i++) { $idx[$cols[$i]] = $i }
  $rows = New-Object System.Collections.Generic.List[string[]]
  while (($line = $reader.ReadLine()) -ne $null) {
    if ($line.Length -eq 0) { continue }
    $rows.Add((Split-CsvLine $line))
  }
  $reader.Close()
  return @{ idx=$idx; rows=$rows }
}

Write-Host "Lendo candidatos..."
$cand = Read-Rows $CandCsv
$ci = $cand.idx
$candidatos = @{}
foreach ($r in $cand.rows) {
  if ($r[$ci['DS_CARGO']] -ne 'DEPUTADO ESTADUAL') { continue }
  $sq = $r[$ci['SQ_CANDIDATO']]
  $candidatos[$sq] = @($r[$ci['NM_URNA_CANDIDATO']], $r[$ci['SG_PARTIDO']])
}
Write-Host "candidatos: $($candidatos.Count)"

# ---------- 3. Ler votacao_secao_2022_AL.csv (sem aspas, so ';' simples) ----------
Write-Host "Lendo votacao por secao (arquivo grande, pode demorar)..."
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$reader = New-Object System.IO.StreamReader($SecaoCsv, $enc1252)
$header = $reader.ReadLine()
$cols = $header.Split(';')
$hi = @{}
for ($i=0; $i -lt $cols.Count; $i++) { $hi[$cols[$i]] = $i }

$maceio = -join @([char]0x4D,[char]0x41,[char]0x43,[char]0x45,[char]0x49,[char]0xD3)
Write-Host "Chave municipio usada na comparacao: $maceio"

$votosBairroAgg = @{}   # "bairro|sq" -> votos
$lineCount = 0
$matchCount = 0
while (($line = $reader.ReadLine()) -ne $null) {
  $lineCount++
  if ($line.Length -eq 0) { continue }
  $f = $line.Split(';')
  if ($f[$hi['NM_MUNICIPIO']] -ne $maceio) { continue }
  if ($f[$hi['DS_CARGO']] -ne 'Deputado Estadual') { continue }
  $zona = $f[$hi['NR_ZONA']]
  if ($zona -ne '33' -and $zona -ne '54') { continue }
  $secao = 0
  [int]::TryParse($f[$hi['NR_SECAO']], [ref]$secao) | Out-Null
  $key = "$zona|$secao"
  $bairro = $secaoToBairro[$key]
  if (-not $bairro) { continue }
  $sq = $f[$hi['SQ_CANDIDATO']]
  if (-not $candidatos.ContainsKey($sq)) { continue }
  $votos = 0
  [int]::TryParse($f[$hi['QT_VOTOS']], [ref]$votos) | Out-Null
  if ($votos -le 0) { continue }
  $matchCount++
  $bkey = "$bairro|$sq"
  if ($votosBairroAgg.ContainsKey($bkey)) { $votosBairroAgg[$bkey] += $votos } else { $votosBairroAgg[$bkey] = $votos }
  if ($lineCount % 2000000 -eq 0) { Write-Host "  ... $lineCount linhas lidas ($($sw.Elapsed))" }
}
$reader.Close()
Write-Host "Total linhas lidas: $lineCount | linhas casadas (Maceio zona33/54 DepEst): $matchCount | tempo: $($sw.Elapsed)"

# ---------- 4. Montar votosBairro / vencedorBairro novos ----------
$votosBairroNovo = New-Object System.Collections.Generic.List[object]
foreach ($kv in $votosBairroAgg.GetEnumerator()) {
  $parts = $kv.Key -split '\|', 2
  $bairro = $parts[0]; $sq = $parts[1]
  $info = $candidatos[$sq]
  $votosBairroNovo.Add(@($bairro, $info[0], $info[1], $kv.Value))
}
Write-Host "votosBairro novo (zonas 33+54): $($votosBairroNovo.Count) linhas"

$porBairro = @{}
foreach ($r in $votosBairroNovo) {
  $b = $r[0]
  if (-not $porBairro.ContainsKey($b)) { $porBairro[$b] = New-Object System.Collections.Generic.List[object] }
  $porBairro[$b].Add($r)
}
$vencedorBairroNovo = New-Object System.Collections.Generic.List[object]
foreach ($kv in $porBairro.GetEnumerator()) {
  $bairro = $kv.Key
  $rows = $kv.Value
  $top = $rows | Sort-Object { -$_[3] } | Select-Object -First 1
  $porPartido = @{}
  $total = 0
  foreach ($r in $rows) {
    $total += $r[3]
    if ($porPartido.ContainsKey($r[2])) { $porPartido[$r[2]] += $r[3] } else { $porPartido[$r[2]] = $r[3] }
  }
  $topPartido = $porPartido.GetEnumerator() | Sort-Object { -$_.Value } | Select-Object -First 1
  $vencedorBairroNovo.Add(@($bairro, $top[1], $top[2], $top[3], $topPartido.Key, $topPartido.Value, $total))
}
Write-Host "vencedorBairro novo (zonas 33+54): $($vencedorBairroNovo.Count) bairros"
Write-Host "--- bairros novos e totais ---"
$vencedorBairroNovo | Sort-Object { $_[0] } | ForEach-Object { Write-Host ($_ -join ' | ') }

$result = [ordered]@{
  votosBairroNovo = $votosBairroNovo
  vencedorBairroNovo = $vencedorBairroNovo
}
$result | ConvertTo-Json -Depth 6 -Compress | Out-File -FilePath $OutJson -Encoding utf8
Write-Host "Gravado em $OutJson"
