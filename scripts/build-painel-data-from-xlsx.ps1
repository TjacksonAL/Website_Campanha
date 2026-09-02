param(
  [Parameter(Mandatory=$true)][string]$XlsxPath,
  [Parameter(Mandatory=$true)][string]$OutJson
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-EntryText($zip, $name) {
  $e = $zip.GetEntry($name)
  if (-not $e) { return $null }
  $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
  $t = $sr.ReadToEnd(); $sr.Close(); return $t
}

function Col-Index([string]$cellRef) {
  $letters = ($cellRef -replace '[0-9]','')
  $idx = 0
  foreach ($ch in $letters.ToCharArray()) { $idx = $idx*26 + ([int][char]$ch - [int][char]'A' + 1) }
  return $idx
}

function Read-Sheet($zip, $sheetFile, $sharedStrings) {
  $xmlText = Get-EntryText $zip $sheetFile
  $doc = New-Object System.Xml.XmlDocument
  $doc.LoadXml($xmlText)
  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  $rowsOut = New-Object 'System.Collections.Generic.Dictionary[int,object]'
  foreach ($rowNode in $doc.SelectNodes("//m:sheetData/m:row", $ns)) {
    $r = [int]$rowNode.GetAttribute("r")
    $rowDict = New-Object 'System.Collections.Generic.Dictionary[int,string]'
    foreach ($c in $rowNode.SelectNodes("m:c", $ns)) {
      $ref = $c.GetAttribute("r")
      $ci = Col-Index $ref
      $t = $c.GetAttribute("t")
      $val = $null
      if ($t -eq "inlineStr") {
        $isNode = $c.SelectSingleNode("m:is", $ns)
        if ($isNode) { $val = $isNode.InnerText }
      } else {
        $vNode = $c.SelectSingleNode("m:v", $ns)
        if ($vNode) {
          $raw = $vNode.InnerText
          if ($t -eq "s") { $val = $sharedStrings[[int]$raw] }
          elseif ($t -eq "b") { $val = if ($raw -eq "1") {"TRUE"} else {"FALSE"} }
          else { $val = $raw }
        }
      }
      $rowDict[$ci] = $val
    }
    $rowsOut[$r] = $rowDict
  }
  return $rowsOut
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($XlsxPath)

# shared strings
$sstText = Get-EntryText $zip "xl/sharedStrings.xml"
$sstDoc = New-Object System.Xml.XmlDocument
$sstDoc.LoadXml($sstText)
$nsS = New-Object System.Xml.XmlNamespaceManager($sstDoc.NameTable)
$nsS.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
$siNodes = $sstDoc.SelectNodes("//m:sst/m:si", $nsS)
$shared = New-Object string[] $siNodes.Count
for ($i=0; $i -lt $siNodes.Count; $i++) { $shared[$i] = $siNodes[$i].InnerText }
Write-Host "sharedStrings: $($shared.Count)"

function Get-Row($rows, $r, $maxCol) {
  $out = New-Object string[] $maxCol
  if ($rows.ContainsKey($r)) {
    $d = $rows[$r]
    foreach ($k in $d.Keys) { if ($k -le $maxCol) { $out[$k-1] = $d[$k] } }
  }
  return $out
}

function Table-ToArray($rows, $headerRow, $lastRow, $maxCol) {
  $list = New-Object System.Collections.Generic.List[object]
  for ($r = $headerRow+1; $r -le $lastRow; $r++) {
    $list.Add((Get-Row $rows $r $maxCol))
  }
  return $list
}

# Le o intervalo real (ref="A4:G43") de uma tabela do Excel, em vez de usar numeros fixos no
# codigo, que quebram silenciosamente quando a tabela cresce numa atualizacao futura.
function Get-TableRange($zip, $tableFile) {
  $txt = Get-EntryText $zip $tableFile
  $x = [xml]$txt
  $ref = $x.table.ref
  $parts = $ref -split ':'
  $startRow = [int]($parts[0] -replace '[A-Za-z]','')
  $endRow = [int]($parts[1] -replace '[A-Za-z]','')
  $endCol = Col-Index ($parts[1] -replace '[0-9]','')
  return @{ headerRow = $startRow; lastRow = $endRow; maxCol = $endCol }
}

Write-Host "Lendo Candidatos (sheet5)..."
$rngCand = Get-TableRange $zip "xl/tables/table3.xml"
$rowsCand = Read-Sheet $zip "xl/worksheets/sheet5.xml" $shared
$candRows = Table-ToArray $rowsCand $rngCand.headerRow $rngCand.lastRow $rngCand.maxCol
Write-Host "linhas candidatos: $($candRows.Count)"

Write-Host "Lendo Votos_por_Municipio (sheet6)..."
$rngVM = Get-TableRange $zip "xl/tables/table4.xml"
$rowsVM = Read-Sheet $zip "xl/worksheets/sheet6.xml" $shared
$vmRows = Table-ToArray $rowsVM $rngVM.headerRow $rngVM.lastRow $rngVM.maxCol
Write-Host "linhas votosMunicipio: $($vmRows.Count)"

Write-Host "Lendo Votos_por_Bairro_Maceio (sheet3)..."
$rngVB = Get-TableRange $zip "xl/tables/table2.xml"
$rowsVB = Read-Sheet $zip "xl/worksheets/sheet3.xml" $shared
$vbRows = Table-ToArray $rowsVB $rngVB.headerRow $rngVB.lastRow $rngVB.maxCol
Write-Host "linhas votosBairro: $($vbRows.Count)"

Write-Host "Lendo Vencedor_por_Bairro_Maceio (sheet2)..."
$rngVEB = Get-TableRange $zip "xl/tables/table1.xml"
$rowsVEB = Read-Sheet $zip "xl/worksheets/sheet2.xml" $shared
$vebRows = Table-ToArray $rowsVEB $rngVEB.headerRow $rngVEB.lastRow $rngVEB.maxCol
Write-Host "linhas vencedorBairro: $($vebRows.Count)"

Write-Host "Lendo Resumo_Partidos_Estado (sheet10)..."
$rngRP = Get-TableRange $zip "xl/tables/table8.xml"
$rowsRP = Read-Sheet $zip "xl/worksheets/sheet10.xml" $shared
$rpRows = Table-ToArray $rowsRP $rngRP.headerRow $rngRP.lastRow $rngRP.maxCol
Write-Host "linhas resumoPartido: $($rpRows.Count)"

$zip.Dispose()

# ---- Build JSON structures ----
$candidatos = @{}
foreach ($r in $candRows) {
  if (-not $r[0]) { continue }
  $sq = $r[0].Trim()
  $nomeUrna = $r[3]
  $siglaPartido = $r[4]
  $situacao = $r[6]
  $eleito = ($r[13] -eq "TRUE" -or $r[13] -eq "Sim" -or $r[13] -eq "SIM")
  $genero = $r[7]
  $corRaca = $r[10]
  $candidatos[$sq] = @($nomeUrna, $siglaPartido, $situacao, $eleito, $genero, $corRaca)
}
Write-Host "candidatos montados: $($candidatos.Count)"

# nome_urna|partido -> SQ (para conseguir o mesmo formato [municipio, SQ, votos] usado no painel)
$nomePartidoToSq = @{}
$dupCount = 0
foreach ($kv in $candidatos.GetEnumerator()) {
  $key = "$($kv.Value[0])|$($kv.Value[1])"
  if ($nomePartidoToSq.ContainsKey($key)) { $dupCount++ } else { $nomePartidoToSq[$key] = $kv.Key }
}
Write-Host "colisoes nome+partido: $dupCount"

$votosMunicipio = New-Object System.Collections.Generic.List[object]
$missSq = 0
foreach ($r in $vmRows) {
  if (-not $r[0]) { continue }
  $mun = $r[0]
  $key = "$($r[1])|$($r[2])"
  $sq = $nomePartidoToSq[$key]
  if (-not $sq) { $missSq++; continue }
  $votos = 0
  [int]::TryParse($r[3], [ref]$votos) | Out-Null
  $votosMunicipio.Add(@($mun, $sq, $votos))
}
Write-Host "votosMunicipio montados: $($votosMunicipio.Count)  (sem SQ encontrado: $missSq)"
Write-Host "AMOSTRA votosMunicipio[0]: $($votosMunicipio[0] -join ' | ')"

$votosBairro = New-Object System.Collections.Generic.List[object]
foreach ($r in $vbRows) {
  if (-not $r[0]) { continue }
  $votos = 0
  [int]::TryParse($r[3], [ref]$votos) | Out-Null
  $votosBairro.Add(@($r[0], $r[1], $r[2], $votos))
}
Write-Host "votosBairro montados: $($votosBairro.Count)"
Write-Host "AMOSTRA votosBairro[0]: $($votosBairro[0] -join ' | ')"

$vencedorBairro = New-Object System.Collections.Generic.List[object]
foreach ($r in $vebRows) {
  if (-not $r[0]) { continue }
  $votosCand = 0; [int]::TryParse($r[3], [ref]$votosCand) | Out-Null
  $votosPartido = 0; [int]::TryParse($r[5], [ref]$votosPartido) | Out-Null
  $total = 0; [int]::TryParse($r[6], [ref]$total) | Out-Null
  $vencedorBairro.Add(@($r[0], $r[1], $r[2], $votosCand, $r[4], $votosPartido, $total))
}
Write-Host "vencedorBairro montados: $($vencedorBairro.Count)"
Write-Host "AMOSTRA vencedorBairro[0]: $($vencedorBairro[0] -join ' | ')"

$resumoPartido = New-Object System.Collections.Generic.List[object]
foreach ($r in $rpRows) {
  if (-not $r[0]) { continue }
  $tot = 0; [int]::TryParse($r[2], [ref]$tot) | Out-Null
  $qc = 0; [int]::TryParse($r[3], [ref]$qc) | Out-Null
  $cad = 0; [int]::TryParse($r[4], [ref]$cad) | Out-Null
  $resumoPartido.Add(@($r[0], $r[1], $tot, $qc, $cad))
}
Write-Host "resumoPartido montados: $($resumoPartido.Count)"

$muSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($row in $votosMunicipio) { $muSet.Add($row[0]) | Out-Null }
$municipios = $muSet | Sort-Object
Write-Host "municipios montados: $($municipios.Count)"
Write-Host "AMOSTRA municipios[0..3]: $($municipios[0..3] -join ' | ')"

# diagnostico da colisao nome+partido
if ($dupCount -gt 0) {
  $seen = @{}
  foreach ($kv in $candidatos.GetEnumerator()) {
    $key = "$($kv.Value[0])|$($kv.Value[1])"
    if ($seen.ContainsKey($key)) { Write-Host "COLISAO: chave='$key' SQ1=$($seen[$key]) SQ2=$($kv.Key)" }
    else { $seen[$key] = $kv.Key }
  }
}

$result = [ordered]@{
  candidatos = $candidatos
  votosMunicipio = $votosMunicipio
  votosBairro = $votosBairro
  vencedorBairro = $vencedorBairro
  resumoPartido = $resumoPartido
  municipios = $municipios
}
$result | ConvertTo-Json -Depth 6 -Compress | Out-File -FilePath $OutJson -Encoding utf8
Write-Host "Gravado em $OutJson"
