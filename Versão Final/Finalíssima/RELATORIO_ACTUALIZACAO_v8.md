# IGDA-BDA — Relatório de actualização v7 → v8

Data: 28 de Setembro de 2026 · Pasta: `Versão Final/Finalíssima/files/`

## 1. Entregas

| Ficheiro | Estado | Conteúdo |
|---|---|---|
| `IGDA_BDA_Construtor_v8.xlsx` | novo | Construtor com 338 candidatos (331 + 7), séries actualizadas, metas 2027 alinhadas com o PDN em cinco categorias, coluna "Origem da Meta 2027", folhas novas `12_Emprego_INE` e `13_Alteracoes_v8`, navegação do painel alargada. Todas as fórmulas preservadas; filtros, formatação condicional e validação alargados às linhas novas; caches dos gráficos regeneradas; recálculo integral forçado na abertura. |
| `IGDA_BDA_Nota_Metodologica_v8.docx` | novo | Nota metodológica actualizada (secção 8.6 sobre a revisão, com a regra de transposição das metas e a fórmula de conversão WGI; 9.2 com as novas limitações; 10.5 cruzamento com o PDN; anexos com vintages e registo de alterações). |
| `IGDA_BDA_Justificacao_Minimos_Maximos_v8.docx` | novo | Justificação linha a linha dos 338 limites; identifica a categoria de cada meta; corrige o erro de redacção da v7 (escalas 0-100 descritas como "0 e 1"). |
| `IGDA_Evolucao_CEX_v10.pptx` | novo | Apresentação CEX com números da v8, gráficos e tabelas actualizados e um diapositivo novo de cruzamento com o PDN. |
| ficheiros v7/v9 | mantidos | Conservados para auditoria. |

## 2. Resultado

O IGDA 2025 passa de 45,7 (v7, 45,73 antes de arredondar) para 45,7 (v8, 45,68). A variação 2015–2025 mantém-se em +3,4 pontos. Alteram-se a Estabilidade Macroeconómica (49,7 → 49,4 em 2025, pela substituição das séries de dívida e de saldo orçamental) e a Segurança Alimentar e Saúde (57,0 → 56,8, pela remoção dos valores 2024–2025 da mortalidade materna sem suporte). A Diversificação não muda (a série do PIB não petrolífero foi verificada sem revisão). A selecção dos 41 indicadores não mudou.

| Dimensão | v7 2025 | v8 2025 | Δ 2015–25 (v8) | Δ 2023–25 (v8) | Escala 2015 = 100 (2025) |
|---|---|---|---|---|---|
| Governança | 33,2 | 33,2 | +5,8 | −0,4 | 121 |
| Macroeconomia | 49,7 | 49,4 | +0,6 | +2,8 | 101 |
| Capital Humano | 59,1 | 59,1 | +8,6 | +0,4 | 117 |
| Inclusão Social | 49,4 | 49,4 | −1,4 | +3,9 | 97 |
| Infraestruturas | 49,9 | 49,9 | +3,7 | +1,0 | 108 |
| Mercado de Trabalho | 56,2 | 56,2 | +1,9 | −0,2 | 104 |
| Saúde e Segurança Alimentar | 57,0 | 56,8 | −2,4 | −1,4 | 96 |
| Diversificação | 24,9 | 24,9 | +4,9 | +1,7 | 124 |
| IGDA-BDA | 45,7 | 45,7 | +3,4 | +1,1 | 108 |

Série IGDA v8, 2015–2025: 42,3 · 42,5 · 42,2 · 42,6 · 42,5 · 39,7 · 43,5 · 44,3 · 44,6 · 45,7 · 45,7 (máximo em 2024, 45,69; 2025 = 45,68, provisório).

## 3. Problemas de qualidade encontrados e corrigidos

1. **Valores em cache desactualizados no ficheiro v7.** O livro foi gravado sem recálculo e mostrava os resultados da v6 (IGDA 2025 = 46,4; Mercado de Trabalho = 66,0), enquanto as fórmulas e os documentos correspondiam à v7 (45,7; 56,2). Um motor de cálculo independente em Python replicou todas as fórmulas e reproduziu exactamente os valores da v6 (diferença máxima 0) e da v7. O v8 força o recálculo na abertura e regenera as caches dos seis gráficos com os resultados v8 (as caches herdadas ainda continham valores da v6).
2. **Dívida pública e saldo orçamental sem rastreabilidade.** As séries "Compilação interna BDA" eram valores arredondados (dívida 2025 = 63; saldo 2017 = +1,5 quando o valor oficial foi −5,7). Substituídas pela série do FMI WEO de Abril de 2026, em precisão total, que compila os dados do MINFIN (fonte histórica declarada pelo FMI; GFSM 2014) — por isso o marcador de fonte nacional é mantido, convenção agora explicitada em 01_Criterios, 10_Fontes e na Nota. O denominador de 2025 do WEO é o PIB nominal das Contas Nacionais Trimestrais do INE (129 255 mil milhões Kz). Níveis não comparáveis com a leitura PDN/MINFIN (perímetro "dívida pública" e PIB pré-rebasing: 66% em 2022 contra 57,4% na série FMI; em 2025, FMI 51,3%, Banco Mundial ≈52%, MINFIN dívida governamental 46,6%). A Nota regista a sensibilidade: sem o bónus nacional, dívida e saldo sairiam do índice e o IGDA 2025 seria 45,9.
3. **Mortalidade materna 2024–2025 sem suporte.** A v7 tinha 185 (2024) e 170 (2025) numa série de estimativas modeladas (MMEIG) que termina em 2023 (183). Os dois valores foram removidos (2024–2025 passam a carry-forward). A estimativa directa do IIMS 2023-24 (170, IC 99–242, sete anos anteriores ao inquérito) não é comparável ponto a ponto com as estimativas modeladas e fica na linha SAU029 como validação cruzada.
4. **Metas 2027 inconsistentes.** 44 metas revistas, em cinco categorias identificadas na nova coluna "Origem da Meta 2027": 7 fixadas directamente nos valores do PDN (base coincide com a série ou o indicador é o do PDN); 19 transpostas à base da série — quando a base 2022 do PDN não coincide com o valor 2022 da série, aplica-se ao valor da série a variação do PDN (aditiva para níveis e proporções, relativa para mortalidade, desemprego e dívida); 9 convertidas de escala (percentis WGI → estimativas pela fórmula meta = ê2022 + [Φ⁻¹(p2027) − Φ⁻¹(p2022)], documentada com os seis pares de percentis e uma linha de sensibilidade; % do PIB não petrolífero → % do PIB); 1 tomada do Balanço anual do MINPLAN como proxy (malária 215); 1 operacional ajustada (escolaridade 9 → 10); 7 removidas porque o conceito do PDN difere da série (IDE bruto vs líquido, capacidade instalada vs produção, alfabetização total vs feminina, dívida externa vs pública, ranking RSF). A coluna Meta 2027 não entra em nenhuma fórmula. Nove dos 41 indicadores têm a meta já cumprida na última observação (sobretudo metas operacionais herdadas) — registado como recomendação.

   | Meta transposta | Base PDN → meta PDN | Valor da série (ano) | Meta v8 |
   |---|---|---|---|
   | HUM002 esperança de vida | 62 → 63 | 64,25 (2022) | 65,2 |
   | HUM006 mortalidade infantil (‰) | 47 → 37 | 33,7 (2022) | 26,5 |
   | HUM003 alfabetização (%) | 76 → 78 | 68,2 (2023) | 70,2 |
   | HUM024 HALE (anos) | 56 → 57 | 54,7 (2021) | 55,7 |
   | HUM025 / SAU04 mortalidade <5 anos (‰) | 69 → 51 | 52 (2024) / 51,9 (2022) | 38,4 |
   | INC006 pobreza, linha nacional (%) | 31 → 28 | 32,3 (2018) | 29,2 |
   | INF001 electrificação (%) | 43 → 49 | 48,5 (2022) | 54,5 |
   | INF002 água potável (%) | 57 → 61 | 66,5 (2022) | 70,5 |
   | INF003 saneamento (%) | 52 → 55 | 50,3 (2022) | 53,3 |
   | LAB002 / LAB018 desemprego OIT (%) | 30 → 25 (definição INE) | 14,1 (2022) | 11,8 / 11,7 |
   | LAB005 formalização (%); LAB004 informal | 22 → 31 | 20,1 (2022) | 29,1; 70,9 |
   | SAU003 / SAU029 mortalidade materna | 199 → 165 | 185 (2022) / 170 (2024) | 153 / 141 |
   | SAU004 despesa em saúde (% PIB) | 3 → 4 | 2,6 (2022) | 3,6 |
   | MAC004 / MAC016 dívida pública (% PIB) | 66 → 60 | 57,4 (2022) | 52,2 |

5. **PIB não petrolífero.** A série da v7 já provinha do Quadro 5 (medidas de volume não ajustadas) do ficheiro trimestral do INE, cuja soma anual reproduz o PIB anual oficial; o IV trimestre de 2025 já estava incluído e não há revisões. A série fica inalterada e a origem passa a estar documentada.
6. **Metadados e grafia.** Nome da dimensão 8 e nomes de indicadores uniformizados para a norma pré-Acordo ("Sector Privado", "protecção", "electricidade", "directa"); linha de pobreza a 2,15 USD com código, origem e nota coerentes (o valor 31,1% de 2018 é o do PIP/UNdata, já usado em versões anteriores); linhas INC028–INC030 renomeadas para as linhas de 4,20 e 8,30 USD (PPC 2021) que os valores extraídos do WDI efectivamente representam; fonte da formalização corrigida de "INSS" para "INE" e valor de 2019 alinhado com o valor anual publicado pelo INE (25,3%); folha 10_Fontes sincronizada com a Base para as 26 linhas alteradas; URLs nas linhas novas.
7. **Estrutura do livro.** Filtro automático (A4:AT342), formatação condicional (AK/AL/AQ) e lista de validação de "Ajuste manual" alargados às sete linhas novas e à coluna AT; navegação do painel com ligações às folhas 11–13; título do gráfico do IGDA em 07_Subindices corrigido (a v7 mostrava um título inválido); folha 12_Emprego_INE com blocos A–D e scores alternativos normalizados com as fronteiras das séries oficiais que substituem.
8. **Erro de redacção na Justificação de mínimos e máximos.** Várias escalas 0-100 estavam descritas como "limites 0 e 1"; corrigido pelo gerador da v8.

## 4. Resposta ao feedback Fable 5.1

| Ponto | O que foi feito | Limite |
|---|---|---|
| 1. Emprego: substituir/complementar as séries OIT com dados do INE | Séries do IEA acrescentadas ao catálogo (originais 13.ª CIET; harmonizadas pela OIT; primeira observação da nova metodologia 19.ª–21.ª CIET). Folha `12_Emprego_INE` com subíndice alternativo 2019–2025 (séries INE normalizadas com as fronteiras das séries oficiais). A diferença face à leitura oficial é de +0,9 p.p. em média em 2019–2024 e salta para +5,4 p.p. em 2025: é o salto que aproxima o efeito de vintage (a série INE harmonizada regista a queda do desemprego de 13,9% para 10,4% em 2025, que a série OIT modelada ainda não incorpora). | O IEA só existe desde 2019: as séries INE não cumprem a cobertura mínima de 80% e não podem substituir as séries OIT no índice 2015–2025 sem quebrar a metodologia. A nova metodologia do INE (IV trim 2025) não tem retropolação. |
| 2. Inclusão: pobreza e protecção social | INC001 fixada na linha de 2,15 USD (referência do PDN, 31,1% em 2018); INC031 nova (3,00 USD PPC 2021, 39,3%); INC032 nova (cobertura do Kwenda: 3,3% dos agregados em 2021 → 14,8% em 2025; meta anual 2025 do MINPLAN como proxy); INC033 nova (segurados INSS: 1,97 M em 2020 → 3,34 M em 2025, quatro observações; meta PDN 4,3 M). | Nenhum atinge 9 anos de observação; ficam documentados e entram automaticamente quando cumprirem a cobertura. Não existe inquérito de despesas posterior ao IDREA 2018-19. |
| 3. Metas 2027 alinhadas com o PDN | Feito (ponto 4 da secção anterior), com regra explícita para as bases que não coincidem (transposição) e categoria de cada meta na coluna "Origem da Meta 2027". | Metas sem correspondência no PDN (por exemplo inflação — MAC003, saldo orçamental — MAC005, peso dos combustíveis nas exportações — DIV010, Índice de Capacidades Produtivas — DIV001) mantêm valores operacionais, identificados como tal. A lista das 13 metas referida no feedback não foi anexada; a v8 responde às inconsistências identificadas a partir do próprio construtor. |
| Alerta sobre despesa em educação | Confirmado com as dotações do OGE: 2,0% do PIB (2024), 1,8% (2025), 1,7% (2026); 6,4–6,9% da despesa total, contra o compromisso do PDN de 11,8%. | A série UIS (executada) só chega a 2023. |
| Variável de política em Capital Humano (escolaridade obrigatória) | Documentada como limitação e recomendação; não substituída para não introduzir uma quebra nesta revisão. | Decisão de desenho a tomar pelo gabinete responsável pelo índice. |

## 5. Recolha de dados (webscraping)

O ambiente desta sessão bloqueia o acesso directo a todos os sítios externos (INE, BNA, MINFIN, MINPLAN, Banco Mundial, FMI, OIT, OMS, UNCTAD, Transparency International e outros), tanto por `curl` como por leitura de páginas. Só a pesquisa web funcionou, com um limite de 200 pesquisas por sessão. Foram lançados sete agentes de recolha em paralelo; três concluíram antes de o limite se esgotar:

| Família | Resultados | Principais dados obtidos |
|---|---|---|
| INE | 58 | Contas Nacionais Anuais Preliminares 2025 (PIB +3,13%; petrolífero −5,23%; PIB nominal 128 302 mil M Kz); IEA Anuário 2025 (desemprego 28,3%; informal 78,8%); IEA I e II trim 2026 (21,3%; 21,5%); IPCN Dez-2025 15,70% e Ago-2026 8,78%; Censo 2024 (36,6 M habitantes; alfabetização 72,6%); IIMS 2023-24. |
| MINFIN / FMI | 103 | Dívida governamental 2025 46,6% (UGD) e FMI 51,3% (rebasing do PIB); défice 2025 4,1% do PIB; OGE 2024–2026 educação e saúde; FMI Art. IV 2026 (crescimento 3,1%, reservas 7,4 meses, projecção 2026 2,3%). |
| MINPLAN / INSS / Kwenda | 81 | Balanço do PDN (1 036 indicadores; 856 sem execução no I trim 2025; 66,9% das prioridades); INSS 3,34 M segurados (2025); Kwenda 1,35 M agregados; metas PDN confirmadas; projecções 2026 (OGE 4,17%/13,7%; BNA 3,5%/13,5%; BM 2,4%/14,9%). |
| BNA, Banco Mundial/WGI, agências ONU, índices internacionais | não executadas | Limite de pesquisas esgotado. As séries respectivas mantêm o vintage da v7 (WDI bulk de 30-06-2026; WGI edição 2025 com dados até 2024; CPI 2025; WEF 2025; UNCTAD; OMS/UNICEF), o mais recente disponível localmente à data de Agosto de 2026, não reverificado nesta revisão. |

Dados locais do repositório usados: INE (Contas Nacionais anuais e trimestrais, IPCN, IEA antiga e nova metodologia, IIMS, Censo 2024, projecções de população), BNA (indicadores externos 1990–2025), FMI WEO (Abril 2026, precisão total), Banco Mundial WDI (Abril 2026), ILOSTAT (estimativas nacionais harmonizadas), UNCTAD, PDN 2023-2027 (texto e 118 metas) e painel MINPLAN 2025.

Para repetir a recolha nas famílias em falta é necessário permitir o acesso de rede aos domínios oficiais na configuração do ambiente (Network access) ou aumentar o limite de pesquisas (`CLAUDE_CODE_MAX_WEB_SEARCHES_PER_SESSION`).

## 6. Verificação adversarial dos entregáveis

Depois da primeira entrega da v8, 50 agentes independentes verificaram os quatro ficheiros em quatro lentes (fidelidade dos dados às fontes locais; integridade do livro Excel; coerência entre documentos e Excel; crítica metodológica), cada achado sujeito a uma tentativa de refutação. Resultado: 46 achados, 42 confirmados e corrigidos nesta versão, 4 refutados (valor 2024 do Kwenda; vintage WDI "Junho 2026"; denominador 2025 do WEO, que é o PIB do INE; máximo de 8 milhões para os segurados INSS). As correcções mais relevantes estão nas secções 2 e 3; o registo completo consta de `scripts_v8/verify_round1.json` e da folha `13_Alteracoes_v8` (85 entradas).

Decisões tomadas nesta revisão que o GEP/gabinete responsável pode querer rever:

- **Transposição das metas** (secção 3, ponto 4): a alternativa era manter os valores absolutos do PDN mesmo quando a série tem outra base, ou deixar as metas em branco. A regra adoptada é a mesma já usada para as metas WGI.
- **Mortalidade materna**: a alternativa era manter o valor do IIMS (170) em 2024 dentro da série modelada; optou-se por não misturar metodologias e registar o IIMS como validação cruzada.
- **Marcador de fonte nacional** da dívida e do saldo (MINFIN via FMI WEO): mantido, porque identifica o produtor primário; a alternativa (retirar o marcador) altera a composição de Macroeconomia (Nota, 9.2).
- **Designação do gabinete autor**: a Nota v7 diz "Gabinete de Planeamento e Controlo" e a apresentação v9 diz "Gabinete de Estratégia e Planeamento (GEP)"; ambas foram mantidas tal como estavam nos originais, por não ser possível confirmar a designação oficial.

## 7. O que não foi possível verificar

- Valores 2025 de fontes internacionais publicados depois de Agosto de 2026 (WGI 2026, WUENIC, SOFI 2026, UNCTAD PCI, IGME, MMEIG, ILO Nov-2026): mantidos da v7 e assinalados como tal na Nota (Secção 9.2 e Anexo B). O PIP do Banco Mundial não foi consultado (sem acesso web); o valor 31,1% de 2018 é herdado de compilações anteriores e coincide com o PDN (31%).
- Recálculo do livro Excel dentro do ambiente: o LibreOffice instalado não tem o módulo Calc. A cadeia de fórmulas foi validada com o motor Python (`igda_engine.py`), que reproduz exactamente a v6 e a v7; o Excel recalcula tudo ao abrir (se necessário, Ctrl+Alt+F9).
- Série anual completa de segurados do INSS (2015–2025) e de agregados pagos pelo Kwenda (2020): não publicadas em excertos acessíveis.
- A lista das 13 metas inconsistentes referida no feedback (mensagem anterior não anexada).

## 8. Reprodutibilidade

Os scripts usados estão versionados em `Versão Final/Finalíssima/scripts_v8/` (ver `README.md` dessa pasta): `igda_engine.py` (motor de cálculo que replica o construtor), `v8_plan.py` (plano de alterações: séries, metas por categoria, regra de transposição, novas linhas, grafia), `build_v8.py`/`build_v8_sheets.py`/`run_build.py` (construção do v8), `make_analysis.py` (drivers, escala comum, metas cumpridas, sensibilidade), `gen_nota.py`, `gen_justificacao.py`, `gen_cex.py` (documentos), e os dados intermédios (`labour_data.json`, `analysis.json`, `v8_results.json`, `sweep_partial.json` com as citações da recolha web, `verify_round1.json` com os achados da verificação). A folha `13_Alteracoes_v8` do construtor contém o registo completo, linha a linha, das alterações, das 44 metas revistas (com categoria) e da conversão WGI.
