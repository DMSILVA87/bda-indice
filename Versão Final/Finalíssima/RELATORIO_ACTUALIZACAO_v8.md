# IGDA-BDA — Relatório de actualização v7 → v8

Data: 28 de Setembro de 2026 · Pasta: `Versão Final/Finalíssima/files/`

## 1. Entregas

| Ficheiro | Estado | Conteúdo |
|---|---|---|
| `IGDA_BDA_Construtor_v8.xlsx` | novo | Construtor com 338 candidatos (331 + 7), séries actualizadas, metas 2027 alinhadas com o PDN em seis categorias, coluna "Origem da Meta 2027", folhas novas `12_Emprego_INE` e `13_Alteracoes_v8`, navegação do painel alargada. Todas as fórmulas preservadas; filtros, formatação condicional e validação alargados às linhas novas; caches dos gráficos regeneradas; recálculo integral forçado na abertura. |
| `IGDA_BDA_Nota_Metodologica_v8.docx` | novo | Nota metodológica actualizada (secção 8.6 sobre a revisão, com a regra de transposição das metas e a fórmula de conversão WGI; 9.2 com as novas limitações e a sensibilidade ao marcador de fonte nacional; 10.5 cruzamento com o PDN; anexos com vintages, siglas e registo de alterações). |
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
2. **Dívida pública e saldo orçamental sem rastreabilidade.** As séries "Compilação interna BDA" eram valores arredondados (dívida 2025 = 63; saldo 2017 = +1,5 quando o valor oficial foi −5,7). Substituídas pela série do FMI WEO de Abril de 2026, em precisão total, que compila os dados do MINFIN (fonte histórica declarada pelo FMI: "Ministry of Finance or Treasury"; GFSM 2014 a partir de 2018 e GFSM 2001 até 2017, quebra assinalada na Nota; perímetro do governo geral = governo central e provincial). O denominador de 2025 do WEO é o PIB nominal das Contas Nacionais Trimestrais do INE (129 255 mil milhões Kz). Níveis não comparáveis com a leitura PDN/MINFIN (perímetro "dívida pública" e PIB pré-rebasing: 66% em 2022 contra 57,4% na série FMI; em 2025, FMI 51,3%, Banco Mundial ≈52%, MINFIN dívida governamental 46,6%).
3. **Marcador de fonte nacional.** A convenção da v7 foi explicitada: o marcador assinala as séries cujos dados de base são produzidos por instituições nacionais, incluindo estimativas de organismos internacionais construídas sobre esses dados (OIT a partir do IEA; IGME e MMEIG a partir dos inquéritos; MINFIN via FMI WEO). Dos 19 indicadores do índice com marcador, 15 são extraídos de fontes internacionais. A Nota quantifica a sensibilidade: sem o bónus na dívida e no saldo, o IGDA 2025 seria 45,9 (dois indicadores substituídos); sem o bónus nos 15 indicadores de canal internacional, seria 46,5 (quatro substituições em Macroeconomia). A convenção da v7 foi mantida e registada como parâmetro da análise de sensibilidade.
4. **Mortalidade materna 2024–2025 sem suporte.** A v7 tinha 185 (2024) e 170 (2025) numa série de estimativas modeladas (MMEIG) que termina em 2023 (183). Os dois valores foram removidos (2024–2025 passam a carry-forward). A estimativa directa do IIMS 2023-24 (170, IC 99–242, sete anos anteriores ao inquérito) não é comparável ponto a ponto com as estimativas modeladas e fica na linha SAU029 como validação cruzada.
5. **Metas 2027 inconsistentes.** 46 metas revistas, em seis categorias identificadas na nova coluna "Origem da Meta 2027": 7 fixadas directamente nos valores do PDN (base coincide com a série, indicador do próprio PDN ou base não confrontável); 21 transpostas à base da série — quando a base 2022 do PDN não coincide com o valor 2022 da série (fonte, definição ou vintage distintos), aplica-se ao valor da série a variação do PDN (aditiva para níveis e proporções, relativa para mortalidade, desemprego e dívida; observação mais próxima quando a série não tem 2022; emprego informal = 100 − formalização; metas em % do PIB não petrolífero convertidas ×0,8 antes de transpor); 7 convertidas de escala (percentis WGI → estimativas pela fórmula meta = ê2022 + [Φ⁻¹(p2027) − Φ⁻¹(p2022)], documentada com os seis pares de percentis e uma linha de sensibilidade; IDE não petrolífero); 2 tomadas do Balanço anual do MINPLAN como proxy e transpostas à série (sarampo 64; malária 219); 1 operacional ajustada (escolaridade 9 → 10); 8 removidas porque o conceito do PDN difere da série ou a meta anterior estava fora das fronteiras (IDE bruto vs líquido, capacidade instalada vs produção, alfabetização total vs feminina, dívida externa vs pública, ranking RSF, desemprego na nova metodologia INE, mortes por malária). A coluna Meta 2027 não entra em nenhuma fórmula. Em seis indicadores do índice (esperança de vida, mortalidade infantil e <5 anos, electrificação, água, desemprego) a última observação já atinge o valor absoluto do PDN mas não a meta transposta, que preserva a variação e não o nível do plano. Nove dos 41 indicadores têm a meta já cumprida na última observação (sobretudo metas operacionais herdadas) — registado como recomendação.

   | Meta transposta | Base → meta (PDN/MINPLAN) | Valor da série (ano) | Meta v8 |
   |---|---|---|---|
   | HUM002 esperança de vida | 62 → 63 | 64,25 (2022) | 65,2 |
   | HUM006 mortalidade infantil (‰) | 47 → 37 | 33,7 (2022) | 26,5 |
   | HUM003 alfabetização (%) | 76 → 78 | 68,2 (2023) | 70,2 |
   | HUM024 HALE (anos) | 56 → 57 | 54,7 (2021) | 55,7 |
   | HUM025 / SAU04 mortalidade <5 anos (‰) | 69 → 51 | 52 (2024) / 51,9 (2022) | 38,4 |
   | INC006 pobreza, linha nacional (%) | 31 → 28 | 32,3 (2018) | 29,3 |
   | INF001 electrificação (%) | 43 → 49 | 48,5 (2022) | 54,5 |
   | INF002 água potável (%) | 57 → 61 | 66,5 (2022) | 70,5 |
   | INF003 saneamento (%) | 52 → 55 | 50,3 (2022) | 53,3 |
   | LAB002 / LAB018 desemprego OIT (%) | 30 → 25 (definição INE) | 14,1 (2022) | 11,8 / 11,7 |
   | LAB005 formalização (%); LAB004 informal | 22 → 31 | 20,1 (2022) | 29,1; 70,9 |
   | SAU003 / SAU029 mortalidade materna | 199 → 165 | 185 (2022) / 170 (2024) | 153 / 141 |
   | SAU004 despesa em saúde (% PIB) | 3 → 4 | 2,6 (2022) | 3,6 |
   | MAC004 / MAC016 dívida pública (% PIB) | 66 → 60 | 57,4 (2022) | 52,2 |
   | DIV015 / MAC013 crédito ao sector privado (% PIB) | 8,7 → 10,0 (convertido de 10,9 → 12,5% do PIB não petrolífero) | 7,0 (2022) | 8,3 |
   | SAU006 cobertura vacinal sarampo (%) | MINSA 66 → 80 (meta anual 2025) | WUENIC 50 (2022) | 64 |
   | SAU020 incidência da malária (por 1 000) | MINSA 255 → 215 (meta anual 2025) | OMS 260 (2022) | 219 |

6. **PIB não petrolífero.** A série da v7 já provinha do Quadro 5 (medidas de volume não ajustadas) do ficheiro trimestral do INE, cuja soma anual reproduz o PIB anual oficial; o IV trimestre de 2025 já estava incluído e não há revisões. A série fica inalterada e a origem passa a estar documentada.
7. **Metadados e grafia.** Nome da dimensão 8, subtemas, nomes de indicadores, unidades, fontes, origens e notas uniformizados para a norma pré-Acordo ("Sector Privado", "protecção", "electricidade", "directa", "extracção", "activa", "efectiva", "actual"); excepção deliberada: "Infraestruturas", nome original da dimensão 5, grafia também usada pelo próprio PDN 2023-2027. Linha de pobreza a 2,15 USD com código, origem e nota coerentes (o valor 31,1% de 2018 é o do PIP/UNdata, já usado em versões anteriores); linhas INC028–INC030 renomeadas para as linhas de 4,20 e 8,30 USD (PPC 2021) que os valores extraídos do WDI representam; as três linhas de pobreza do IDREA 2018-19 passam a ter o mesmo marcador nacional; fonte da formalização corrigida de "INSS" para "INE" e valor de 2019 alinhado com o valor anual publicado pelo INE (25,3%); os valores de 2025 do Anuário IEA identificados como média dos trimestres I–III na metodologia antiga; folha 10_Fontes sincronizada com a Base para as 30 linhas alteradas, com URLs nas linhas novas e 46 URLs fictícios herdados da v7 (padrão WDI com códigos internos) substituídos pelo endereço institucional (BNA, INE, FMI WEO, etc.); citação de página do PDN corrigida para a despesa em educação (p.22).
8. **Estrutura do livro.** Filtro automático (A4:AT342), formatação condicional (AK/AL/AQ) e lista de validação de "Ajuste manual" alargados às sete linhas novas e à coluna AT; navegação do painel com ligações às folhas 11–13 e ligação de retorno em 11_Escala_Comum; título do gráfico do IGDA em 07_Subindices corrigido (a v7 mostrava um título inválido); folha 12_Emprego_INE com blocos A–D e scores alternativos normalizados com as fronteiras das séries oficiais que substituem.
9. **Erro de redacção na Justificação de mínimos e máximos.** Várias escalas 0-100 estavam descritas como "limites 0 e 1"; corrigido pelo gerador da v8, que também deixa de descrever um mínimo positivo como "incidência nula".

## 4. Resposta ao feedback Fable 5.1

| Ponto | O que foi feito | Limite |
|---|---|---|
| 1. Emprego: substituir/complementar as séries OIT com dados do INE | Séries do IEA acrescentadas ao catálogo (originais 13.ª CIET; harmonizadas pela OIT; primeira observação da nova metodologia 19.ª–21.ª CIET). Folha `12_Emprego_INE` com subíndice alternativo 2019–2025 (séries INE normalizadas com as fronteiras das séries oficiais). A diferença face à leitura oficial é de +0,9 p.p. em média em 2019–2024 e salta para +5,4 p.p. em 2025: é o salto que aproxima o efeito de vintage (a série INE harmonizada regista a queda do desemprego de 13,9% para 10,4% em 2025, que a série OIT modelada ainda não incorpora). | O IEA só existe desde 2019: as séries INE não cumprem a cobertura mínima de 80% e não podem substituir as séries OIT no índice 2015–2025 sem quebrar a metodologia. A nova metodologia do INE (IV trim 2025) não tem retropolação. |
| 2. Inclusão: pobreza e protecção social | INC001 fixada na linha de 2,15 USD (referência do PDN, 31,1% em 2018); INC031 nova (3,00 USD PPC 2021, 39,3%); INC032 nova (cobertura do Kwenda: 3,3% dos agregados em 2021 → 14,8% em 2025; meta anual 2025 do MINPLAN como proxy); INC033 nova (segurados INSS: 1,97 M em 2020 → 3,34 M em 2025, quatro observações; meta PDN 4,3 M). | Nenhum atinge 9 anos de observação; ficam documentados e entram automaticamente quando cumprirem a cobertura. Não existe inquérito de despesas posterior ao IDREA 2018-19. |
| 3. Metas 2027 alinhadas com o PDN | Parcialmente feito (ponto 5 da secção anterior): coluna alinhada e documentada, com regra explícita para as bases que não coincidem (transposição) e categoria de cada meta na coluna "Origem da Meta 2027". Fica por rever a coerência das nove metas já cumpridas na última observação. | Metas sem correspondência no PDN (por exemplo inflação — MAC003, saldo orçamental — MAC005, peso dos combustíveis nas exportações — DIV010, Índice de Capacidades Produtivas — DIV001) mantêm valores operacionais, identificados como tal. A lista das 13 metas referida no feedback não foi anexada; a v8 responde às inconsistências identificadas a partir do próprio construtor. |
| Alerta sobre despesa em educação | Confirmado com as dotações do OGE: 2,0% do PIB (2024), 1,8% (2025), 1,7% (2026); 6,4–6,9% da despesa total, contra o compromisso do PDN (p.22) de 11,8%. | A série UIS (executada) só chega a 2023. |
| Variável de política em Capital Humano (escolaridade obrigatória) | Documentada como limitação e recomendação; não substituída para não introduzir uma quebra nesta revisão. | Decisão de desenho a tomar pelo gabinete responsável pelo índice. |

## 5. Recolha de dados (webscraping)

O ambiente desta sessão bloqueia o acesso directo a todos os sítios externos (INE, BNA, MINFIN, MINPLAN, Banco Mundial, FMI, OIT, OMS, UNCTAD, Transparency International e outros), tanto por `curl` como por leitura de páginas. Só a pesquisa web funcionou, com um limite de 200 pesquisas por sessão. Foram lançados sete agentes de recolha em paralelo; três concluíram antes de o limite se esgotar:

| Família | Resultados | Principais dados obtidos |
|---|---|---|
| INE | 58 | Contas Nacionais Anuais Preliminares 2025 (PIB +3,13%; petrolífero −5,23%; PIB nominal 128 302 mil M Kz); IEA Anuário 2025 (desemprego 28,3%; informal 78,8%; trimestres I–III na metodologia antiga); IEA I e II trim 2026 (21,3%; 21,5%); IPCN Dez-2025 15,70% e Ago-2026 8,78%; Censo 2024 (36,6 M habitantes; alfabetização 72,6%); IIMS 2023-24. |
| MINFIN / FMI | 103 | Dívida governamental 2025 46,6% (UGD) e FMI 51,3% (rebasing do PIB); défice 2025 4,1% do PIB; OGE 2024–2026 educação e saúde; FMI Art. IV 2026 (crescimento 3,1%, reservas 7,4 meses, projecção 2026 2,3%). |
| MINPLAN / INSS / Kwenda | 81 | Balanço do PDN (1 036 indicadores; 856 sem execução no I trim 2025; 66,9% das prioridades); INSS 3,34 M segurados (2025); Kwenda 1,35 M agregados; metas PDN confirmadas; vacinação (MINSA 76% em 2025) e malária (298 por mil em 2025); projecções 2026 (OGE 4,17%/13,7%; BNA 3,5%/13,5%; BM 2,4%/14,9%). |
| BNA, Banco Mundial/WGI, agências ONU, índices internacionais | não executadas | Limite de pesquisas esgotado. As séries respectivas mantêm o vintage da v7 (WDI bulk de 30-06-2026; WGI edição 2025 com dados até 2024; CPI 2025; WEF 2025; UNCTAD; OMS/UNICEF), o mais recente disponível localmente à data de Agosto de 2026, não reverificado nesta revisão. |

Dados locais do repositório usados: INE (Contas Nacionais anuais e trimestrais, IPCN, IEA antiga e nova metodologia, IIMS, Censo 2024, projecções de população), BNA (indicadores externos 1990–2025), FMI WEO (Abril 2026, precisão total), Banco Mundial WDI (Abril 2026), ILOSTAT (estimativas nacionais harmonizadas), UNCTAD, PDN 2023-2027 (texto e 118 metas) e painel MINPLAN 2025.

Para repetir a recolha nas famílias em falta é necessário permitir o acesso de rede aos domínios oficiais na configuração do ambiente (Network access) ou aumentar o limite de pesquisas (`CLAUDE_CODE_MAX_WEB_SEARCHES_PER_SESSION`).

## 6. Verificação adversarial dos entregáveis

Duas rondas de verificação independente foram executadas sobre os quatro ficheiros, em quatro lentes (fidelidade dos dados às fontes locais; integridade do livro Excel; coerência entre documentos e Excel; crítica metodológica), com cada achado sujeito a uma tentativa de refutação por um segundo agente.

- **Ronda 1** (50 agentes): 46 achados, 42 confirmados e 4 refutados (valor 2024 do Kwenda; vintage WDI "Junho 2026"; denominador 2025 do WEO, que é o PIB do INE; máximo de 8 milhões para os segurados INSS). Dos 42 confirmados, 41 foram corrigidos e 1 (designação do gabinete autor) foi registado como decisão pendente. Registo em `scripts_v8/verify_round1.json`.
- **Ronda 2** (32 agentes, sobre os ficheiros corrigidos): 28 achados, todos de severidade baixa ou média excepto um (a convenção do marcador de fonte nacional), tratados nesta versão: regra de transposição completada (observação mais próxima, complemento, conversão prévia de % do PIB não petrolífero) e aplicada de forma consistente (pobreza pela linha nacional em modo aditivo; crédito ao sector privado e metas MINPLAN transpostas; desemprego da nova metodologia INE sem meta comparável); convenção do marcador nacional reformulada e sensibilidade alargada aos 15 indicadores de canal internacional; frase do cruzamento com o PDN alinhada com a coluna de veredictos (cinco consistentes, duas parciais, uma inconsistente); grafia estendida à coluna de notas; 46 URLs fictícios substituídos; ligação de retorno em 11_Escala_Comum; siglas em falta; citação de página do PDN (p.22); metadados do WEO ("Ministry of Finance or Treasury"; quebra GFSM 2001/2014); registo de INC001 reclassificado como série; "seis categorias" em vez de "cinco"; estatuto "parcialmente concluído" das metas do PDN uniformizado entre Nota, CEX e relatório. Registo em `scripts_v8/verify_round2.json`.

Decisões tomadas nesta revisão que o gabinete responsável pode querer rever:

- **Transposição das metas** (secção 3, ponto 5): a alternativa era manter os valores absolutos do PDN mesmo quando a série tem outra base, ou deixar as metas em branco. A regra adoptada é a mesma já usada para as metas WGI e torna seis metas mais exigentes do que o valor absoluto do PDN.
- **Mortalidade materna**: a alternativa era manter o valor do IIMS (170) em 2024 dentro da série modelada; optou-se por não misturar metodologias e registar o IIMS como validação cruzada.
- **Marcador de fonte nacional**: mantida a convenção da v7 (incluindo estimativas internacionais construídas sobre dados nacionais); a alternativa literal (marcador só para canais nacionais) alteraria a composição de Macroeconomia (Nota, 9.2).
- **Grafia "Infraestruturas"**: mantida a forma do nome original da dimensão 5 (também a do PDN), apesar de os documentos seguirem a norma pré-Acordo ("infra-estruturas").
- **Designação do gabinete autor**: a Nota v7 diz "Gabinete de Planeamento e Controlo" e a apresentação v9 diz "Gabinete de Estratégia e Planeamento (GEP)"; ambas foram mantidas tal como estavam nos originais, por não ser possível confirmar a designação oficial.

## 7. O que não foi possível verificar

- Valores 2025 de fontes internacionais publicados depois de Agosto de 2026 (WGI 2026, WUENIC, SOFI 2026, UNCTAD PCI, IGME, MMEIG, ILO Nov-2026): mantidos da v7 e assinalados como tal na Nota (Secção 9.2 e Anexo B). O PIP do Banco Mundial não foi consultado (sem acesso web); o valor 31,1% de 2018 é herdado de compilações anteriores e coincide com o PDN (31%).
- Recálculo do livro Excel dentro do ambiente: o LibreOffice instalado não tem o módulo Calc. A cadeia de fórmulas foi validada com o motor Python (`igda_engine.py`), que reproduz exactamente a v6 e a v7; o Excel recalcula tudo ao abrir (se necessário, Ctrl+Alt+F9).
- Série anual completa de segurados do INSS (2015–2025) e de agregados pagos pelo Kwenda (2020): não publicadas em excertos acessíveis.
- A lista das 13 metas inconsistentes referida no feedback (mensagem anterior não anexada).

## 8. Reprodutibilidade

Os scripts usados estão versionados em `Versão Final/Finalíssima/scripts_v8/` (ver `README.md` dessa pasta): `igda_engine.py` (motor de cálculo que replica o construtor), `v8_plan.py` (plano de alterações: séries, metas por categoria, regra de transposição, novas linhas, grafia, URLs), `build_v8.py`/`build_v8_sheets.py`/`run_build.py` (construção do v8), `make_analysis.py` (drivers, escala comum, metas cumpridas, sensibilidade ao marcador nacional), `gen_nota.py`, `gen_justificacao.py`, `gen_cex.py` (documentos), e os dados intermédios (`labour_data.json`, `analysis.json`, `v8_results.json`, `sweep_partial.json` com as citações da recolha web, `verify_round1.json` e `verify_round2.json` com os achados das verificações). A folha `13_Alteracoes_v8` do construtor contém o registo completo, linha a linha, das 90 alterações, das 46 metas revistas (com categoria) e da conversão WGI.
