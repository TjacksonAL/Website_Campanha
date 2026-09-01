# Bases de dados do painel eleitoral (AL 2022)

O painel [painel_eleitoral_al2022.html](../painel_eleitoral_al2022.html) não busca dados externos: todo o conteúdo fica
embutido em `<script id="raw-data" type="application/json">` (perto da linha ~504 do arquivo), como um objeto `DATA` com
os blocos `candidatos`, `votosMunicipio`, `votosBairro`, `vencedorBairro`, `resumoPartido` e `municipios`.

**Escopo do painel: só Deputado Estadual.** Não inclui Governador, Vice, Senador nem Deputado Federal.

## Onde estão as bases originais

Nenhuma base fica neste repositório (são grandes/binárias demais para versionar em git). Estão em:

`C:\Users\Thulio Jack\OneDrive\Eleição\Banco de dados\`

| Arquivo | Conteúdo | Uso |
|---|---|---|
| `base_eleitoral_AL_2022_deputado_estadual_1.xlsx` | **Planilha-fonte completa** (13 abas) já com todo o processamento feito: candidatos, votos por município, votos por bairro em Maceió, vencedor por bairro, resumo por partido etc. | **Fonte principal** — usar `scripts/build-painel-data-from-xlsx.ps1` |
| `consulta_cand_2022_AL.csv` | Portal de Dados Abertos do TSE — cadastro de candidatos (bruto) | Fonte alternativa/atualização (`scripts/build-painel-data.ps1`), cobre só `candidatos`/`votosMunicipio`/`resumoPartido`/`municipios` |
| `votacao_candidato_munzona_2022_AL.csv` | TSE — votação por candidato, por município/zona (bruto) | idem acima |
| `votacao_secao_2022_AL.csv` | TSE — votação por seção, com nome/endereço do local de votação (bruto, 88MB) | Só seria necessário para remontar `votosBairro`/`vencedorBairro` do zero a partir do bruto — hoje não é mais necessário, ver abaixo |

## A planilha xlsx é a fonte de verdade

`base_eleitoral_AL_2022_deputado_estadual_1.xlsx` tem estas abas relevantes (as demais são visualização/auxiliares):

| Aba | Linhas de dado | Vira no JSON |
|---|---|---|
| `Candidatos` | 282 | `DATA.candidatos` |
| `Votos_por_Municipio` | 10.640 | `DATA.votosMunicipio` |
| `Votos_por_Bairro_Maceio` | 6.153 | `DATA.votosBairro` |
| `Vencedor_por_Bairro_Maceio` | 39 | `DATA.vencedorBairro` |
| `Resumo_Partidos_Estado` | 19 | `DATA.resumoPartido` |
| `Votos_por_Local_Votacao`, `Vencedor_por_Municipio`, `Vencedor_por_Local_Votacao`, `Aux_Top5_*` | — | não usadas pelo painel hoje (dados auxiliares/mais granulares) |

**Validado em 2026-09-01**: rodando [`scripts/build-painel-data-from-xlsx.ps1`](../scripts/build-painel-data-from-xlsx.ps1)
contra esse arquivo, o resultado bate **exatamente** com o que está publicado hoje no painel — mesmas contagens
(282/10.640/6.153/39/19/102), mesma soma total de votos em cada bloco (`votosMunicipio`=1.549.561,
`votosBairro`=`vencedorBairro`=272.916) e conteúdo linha a linha idêntico (testado bairro a bairro). Ou seja, essa
planilha é a fonte original usada para montar o painel atual.

Detalhe técnico: a aba `Votos_por_Municipio` guarda o candidato pelo **nome de urna**, não pelo `SQ_Candidato` que o
JS do painel usa como chave (`DATA.candidatos[String(sq)]`). O script reconstrói o SQ fazendo o join
`nome_urna + partido -> SQ` a partir da aba `Candidatos`. Existe 1 colisão conhecida (dois candidatos "MARCOS FERREIRA"
do PV com o mesmo nome de urna e número, SQs `20001612356` e `20001728453`) — isso já existia nos dados originais e
não afeta os totais por partido, só a atribuição individual de voto entre esses dois SQs específicos.

## Fluxo de atualização

Quando o Thulio enviar dados atualizados:

1. **Se vier uma planilha `.xlsx` no mesmo formato** (aba `Candidatos`, `Votos_por_Municipio`, `Votos_por_Bairro_Maceio`,
   `Vencedor_por_Bairro_Maceio`, `Resumo_Partidos_Estado`): rodar
   ```powershell
   & scripts/build-painel-data-from-xlsx.ps1 -XlsxPath "<caminho do xlsx>" -OutJson "<arquivo de saída>"
   ```
   Isso já cobre os 6 blocos, incluindo os dados de bairro.
2. **Se vierem só os CSVs brutos do TSE** (`consulta_cand_2022_AL.csv` + `votacao_candidato_munzona_2022_AL.csv`):
   rodar
   ```powershell
   & scripts/build-painel-data.ps1 -CandCsv "<consulta_cand...csv>" -MunzonaCsv "<votacao_candidato_munzona...csv>" -OutJson "<arquivo de saída>"
   ```
   Isso cobre `candidatos`/`votosMunicipio`/`resumoPartido`/`municipios`, mas **não** `votosBairro`/`vencedorBairro`
   (precisaria reprocessar `votacao_secao_2022_AL.csv` cruzando local de votação → bairro, o que ainda não está
   automatizado).
3. Substituir o conteúdo de `<script id="raw-data">` em `painel_eleitoral_al2022.html` pelo JSON gerado.

Ambos os scripts usam Windows PowerShell 5.1 puro (sem dependências externas) porque este ambiente não tem
Python nem Node instalados. Os CSVs/xlsx do TSE usam `;` como separador e codificação Windows-1252 — os scripts já
tratam isso.
