# Rebuilds candidatos / votosMunicipio / resumoPartido / municipios blocks
# for painel_eleitoral_al2022.html from the two smaller TSE source CSVs.
# (votosBairro / vencedorBairro are NOT covered here - see dados/README.md)

param(
  [Parameter(Mandatory=$true)][string]$CandCsv,
  [Parameter(Mandatory=$true)][string]$MunzonaCsv,
  [Parameter(Mandatory=$true)][string]$OutJson
)

$enc = [System.Text.Encoding]::GetEncoding(1252)
$splitRegex = New-Object System.Text.RegularExpressions.Regex ';(?=(?:[^"]*"[^"]*")*[^"]*$)'

function Split-CsvLine($line) {
  $raw = $splitRegex.Split($line)
  $out = New-Object string[] $raw.Count
  for ($i=0; $i -lt $raw.Count; $i++) {
    $v = $raw[$i]
    if ($v.Length -ge 2 -and $v[0] -eq '"' -and $v[$v.Length-1] -eq '"') {
      $v = $v.Substring(1, $v.Length-2)
    }
    $out[$i] = $v
  }
  return $out
}

function Read-Rows($path) {
  $reader = New-Object System.IO.StreamReader($path, $enc)
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
$partidoNome = @{}
foreach ($r in $cand.rows) {
  if ($r[$ci['DS_CARGO']] -ne 'DEPUTADO ESTADUAL') { continue }
  $sq = $r[$ci['SQ_CANDIDATO']]
  $nomeUrna = $r[$ci['NM_URNA_CANDIDATO']]
  $sigla = $r[$ci['SG_PARTIDO']]
  $situacao = $r[$ci['DS_SIT_TOT_TURNO']]
  $eleito = $situacao.StartsWith('ELEITO')
  $genero = $r[$ci['DS_GENERO']]
  $corRaca = $r[$ci['DS_COR_RACA']]
  $candidatos[$sq] = @($nomeUrna, $sigla, $situacao, $eleito, $genero, $corRaca)
  if (-not $partidoNome.ContainsKey($sigla)) { $partidoNome[$sigla] = $r[$ci['NM_PARTIDO']] }
}
Write-Host "candidatos (DEPUTADO ESTADUAL): $($candidatos.Count)"

Write-Host "Lendo votacao por municipio/zona (pode demorar)..."
$mz = Read-Rows $MunzonaCsv
$mi = $mz.idx

$agg = @{}  # key "MUNICIPIO|SQ" -> votos
foreach ($r in $mz.rows) {
  if (-not $candidatos.ContainsKey($r[$mi['SQ_CANDIDATO']])) { continue }
  $mun = $r[$mi['NM_MUNICIPIO']]
  $sq = $r[$mi['SQ_CANDIDATO']]
  $votos = 0
  [int]::TryParse($r[$mi['QT_VOTOS_NOMINAIS']], [ref]$votos) | Out-Null
  $key = "$mun|$sq"
  if ($agg.ContainsKey($key)) { $agg[$key] += $votos } else { $agg[$key] = $votos }
}

$votosMunicipio = New-Object System.Collections.Generic.List[object]
$municipiosSet = New-Object System.Collections.Generic.HashSet[string]
$votosPorPartido = @{}   # sigla -> total votos
$candComVotoPorPartido = @{}  # sigla -> set of SQ
foreach ($kv in $agg.GetEnumerator()) {
  if ($kv.Value -le 0) { continue }
  $parts = $kv.Key -split '\|', 2
  $mun = $parts[0]; $sq = $parts[1]
  $municipiosSet.Add($mun) | Out-Null
  $votosMunicipio.Add(@($mun, $sq, $kv.Value))
  $sigla = $candidatos[$sq][1]
  if ($votosPorPartido.ContainsKey($sigla)) { $votosPorPartido[$sigla] += $kv.Value } else { $votosPorPartido[$sigla] = $kv.Value }
  if (-not $candComVotoPorPartido.ContainsKey($sigla)) { $candComVotoPorPartido[$sigla] = (New-Object System.Collections.Generic.HashSet[string]) }
  $candComVotoPorPartido[$sigla].Add($sq) | Out-Null
}
Write-Host "votosMunicipio rows: $($votosMunicipio.Count)"
Write-Host "municipios: $($municipiosSet.Count)"

$eleitosPorPartido = @{}
foreach ($kv in $candidatos.GetEnumerator()) {
  if ($kv.Value[3]) {
    $sigla = $kv.Value[1]
    if ($eleitosPorPartido.ContainsKey($sigla)) { $eleitosPorPartido[$sigla] += 1 } else { $eleitosPorPartido[$sigla] = 1 }
  }
}

$resumoPartido = New-Object System.Collections.Generic.List[object]
foreach ($kv in $votosPorPartido.GetEnumerator()) {
  $sigla = $kv.Key
  $nCand = $candComVotoPorPartido[$sigla].Count
  $nEleito = if ($eleitosPorPartido.ContainsKey($sigla)) { $eleitosPorPartido[$sigla] } else { 0 }
  $resumoPartido.Add(@($sigla, $partidoNome[$sigla], $kv.Value, $nCand, $nEleito))
}
$resumoPartido = $resumoPartido | Sort-Object { -$_[2] }

$municipios = $municipiosSet | Sort-Object

$result = [ordered]@{
  candidatos = $candidatos
  votosMunicipio = $votosMunicipio
  resumoPartido = $resumoPartido
  municipios = $municipios
}
$result | ConvertTo-Json -Depth 6 -Compress | Out-File -FilePath $OutJson -Encoding utf8
Write-Host "Gravado em $OutJson"
Write-Host "--- resumoPartido ---"
$resumoPartido | ForEach-Object { Write-Host ($_ -join ' | ') }
$totalEleitos = ($resumoPartido | ForEach-Object { $_[4] } | Measure-Object -Sum).Sum
Write-Host "Soma eleitos: $totalEleitos"
