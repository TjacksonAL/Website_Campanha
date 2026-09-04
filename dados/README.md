# Bases de dados do painel eleitoral (AL 2022)

O painel [painel_eleitoral_al2022.html](../painel_eleitoral_al2022.html) não busca dados externos: todo o conteúdo fica
embutido em `<script id="raw-data" type="application/json">` (perto da linha ~504 do arquivo), como um objeto `DATA` com
os blocos `candidatos`, `votosMunicipio`, `votosBairro`, `vencedorBairro`, `resumoPartido` e `municipios`.

**Escopo do painel: só Deputado Estadual.** Não inclui Governador, Vice, Senador nem Deputado Federal.

**Cobertura de bairros de Maceió: 100% desde 2026-09-02.** Maceió tem exatamente 5 zonas eleitorais (1ª, 2ª, 3ª, 33ª e
54ª — confirmado varrendo `votacao_secao_2022_AL.csv`). Antes só as zonas 1ª/2ª/3ª estavam mapeadas (39 bairros); as
zonas 33ª/54ª (12 bairros: Tabuleiro do Martins, Clima Bom I/II, Cidade Universitária, Benedito Bentes/I/II,
Petrópolis, Antares, Santos Dumont, Santa Lúcia, Santa Amélia) foram adicionadas cruzando
`mapeamento_bairros_zonas_33_54.xlsx` com `votacao_secao_2022_AL.csv` — ver [`scripts/build-zonas-33-54.ps1`](../scripts/build-zonas-33-54.ps1).
Resultado: `votosBairro` foi de 6.153 para 8.284 linhas, `vencedorBairro` de 39 para 51 bairros, sem nenhuma
sobreposição de nome de bairro entre as duas fontes. **A planilha-fonte `base_eleitoral_AL_2022_deputado_estadual_1.xlsx`
foi atualizada com essas mesmas linhas** (abas `Votos_por_Bairro_Maceio` e `Vencedor_por_Bairro_Maceio`, via automação
do Excel), então ela e o painel publicado estão sincronizados. Existe um backup do estado anterior dela em
`base_eleitoral_AL_2022_deputado_estadual_1.BACKUP-2026-09-02.xlsx`, na mesma pasta.

Nota técnica: esse arquivo está em uma pasta OneDrive com AutoSave — alterações feitas por automação (COM) podem ser
sincronizadas para a nuvem mesmo sem um `.Save()` explícito ter sido concluído, então sempre faça uma cópia de
segurança antes de editar esse arquivo por script.

## Onde estão as bases originais

Nenhuma base fica neste repositório (são grandes/binárias demais para versionar em git). Estão em:

`C:\Users\Thulio Jack\OneDrive\Eleição\Banco de dados\`

| Arquivo | Conteúdo | Uso |
|---|---|---|
| `base_eleitoral_AL_2022_deputado_estadual_1.xlsx` | **Planilha-fonte completa** (13 abas) já com todo o processamento feito: candidatos, votos por município, votos por bairro em Maceió (zonas 1ª/2ª/3ª), vencedor por bairro, resumo por partido etc. | Fonte principal — usar `scripts/build-painel-data-from-xlsx.ps1` |
| `consulta_cand_2022_AL.csv` | Portal de Dados Abertos do TSE — cadastro de candidatos (bruto) | Fonte alternativa/atualização (`scripts/build-painel-data.ps1`), cobre só `candidatos`/`votosMunicipio`/`resumoPartido`/`municipios` |
| `votacao_candidato_munzona_2022_AL.csv` | TSE — votação por candidato, por município/zona (bruto) | idem acima |
| `votacao_secao_2022_AL.csv` | TSE — votação por seção, com nome/endereço do local de votação (bruto, 88MB) | Usado para calcular os votos por bairro das zonas 33ª/54ª (ver abaixo) |
| `tre-al-eleicoes-2022-zona-locais-votacao-secao.pdf` | Relatório oficial do TRE-AL — seção → local de votação → endereço, todas as zonas do estado (316 páginas) | Fonte bruta do mapeamento de bairro; documento original em PDF |
| `mapeamento_bairros_zonas_33_54.xlsx` | Extração já processada do PDF acima, só para as zonas 33ª (Tabuleiro do Martins/Clima Bom I/II) e 54ª (Benedito Bentes, Cidade Universitária, Petrópolis, Antares, Santos Dumont, Santa Lúcia, Santa Amélia): Zona, Seção, Local, Endereço, Bairro | Usado (2026-09-02) para completar a cobertura de bairros de Maceió — validado 100% (64/64 combinações local+endereço conferidas contra o PDF bruto) |

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

## Bases de 2026 (candidaturas)

Em `Banco de dados\Dados 2026\` (fora do repo, mesmo padrão das demais) estão os cadastros de candidatos do TSE
para a eleição de 2026 (candidaturas registradas, pleito só em outubro/2026 — ainda não há dados de votação):

| Arquivo | Conteúdo |
|---|---|
| `consulta_cand_2026_AL.csv` | Cadastro de candidatos (mesma estrutura de `consulta_cand_2022_AL.csv`) — 141 candidatos a Deputado Estadual |
| `consulta_cand_complementar_2026_AL.csv` | Dados complementares: situação de julgamento do registro, teto de gastos de campanha, etc. |
| `rede_social_candidato_2026_AL.csv` | Redes sociais declaradas por candidato |

Análise comparativa 2022 vs 2026 (composição partidária, candidatos recorrentes via CPF, perfil demográfico) feita
em 2026-09-04, com foco adicional na candidata Walkiria Ferreira (PL) — ver `Analise_2022_2026_Walkiria_Ferreira.pdf`
na pasta `Eleição` (não versionado no repo, é um relatório para a campanha, não dado-fonte nem código).

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
