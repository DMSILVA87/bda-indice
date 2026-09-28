# IGDA-BDA — Relatório de actualização v7 → v8

Data: 28 de Setembro de 2026 · Pasta: `Versão Final/Finalíssima/files/`

## 1. Entregas

| Ficheiro | Estado | Conteúdo |
|---|---|---|
| `IGDA_BDA_Construtor_v8.xlsx` | novo | Construtor com 338 candidatos (331 + 7), séries actualizadas, metas 2027 alinhadas com o PDN, coluna "Origem da Meta 2027", folhas novas `12_Emprego_INE` e `13_Alteracoes_v8`. Todas as fórmulas preservadas; recálculo integral forçado na abertura. |
| `IGDA_BDA_Nota_Metodologica_v8.docx` | novo | Nota metodológica actualizada (secção 8.6 sobre a revisão, 10.5 cruzamento com o PDN, anexos com vintages e registo de alterações). |
| `IGDA_BDA_Justificacao_Minimos_Maximos_v8.docx` | novo | Justificação linha a linha dos 338 limites; corrige o erro de redacção da v7 (escalas 0-100 descritas como "0 e 1"). |
| `IGDA_Evolucao_CEX_v10.pptx` | novo | Apresentação CEX com números da v8, gráficos e tabelas actualizados e um diapositivo novo de cruzamento com o PDN. |
| ficheiros v7/v9 | mantidos | Conservados para auditoria. |

## 2. Resultado

O IGDA 2025 passa de 45,7 (v7) para 45,7 (v8, 45,66 antes de arredondar). A variação 2015–2025 mantém-se em +3,4 pontos. A única dimensão com alteração visível é a Estabilidade Macroeconómica (49,7 → 49,3 em 2025), por causa da substituição das séries de dívida e de saldo orçamental. A selecção dos 41 indicadores não mudou.

| Dimensão | v7 2025 | v8 2025 | Δ 2015–25 (v8) | Δ 2023–25 (v8) |
|---|---|---|---|---|
| Governança | 33,2 | 33,2 | +5,8 | −0,4 |
| Macroeconomia | 49,7 | 49,3 | +0,6 | +2,7 |
| Capital Humano | 59,1 | 59,1 | +8,6 | +0,4 |
| Inclusão Social | 49,4 | 49,4 | −1,4 | +3,9 |
| Infraestruturas | 49,9 | 49,9 | +3,7 | +1,0 |
| Mercado de Trabalho | 56,2 | 56,2 | +1,9 | −0,2 |
| Saúde e Segurança Alimentar | 57,0 | 57,0 | −2,2 | −1,2 |
| Diversificação | 24,9 | 24,8 | +4,8 | +1,5 |
| IGDA-BDA | 45,7 | 45,7 | +3,4 | +1,1 |

## 3. Problemas de qualidade encontrados e corrigidos

1. **Valores em cache desactualizados no ficheiro v7.** O livro foi gravado sem recálculo e mostrava os resultados da v6 (IGDA 2025 = 46,4; Mercado de Trabalho = 66,0), enquanto as fórmulas e os documentos correspondiam à v7 (45,7; 56,2). Um motor de cálculo independente em Python replicou todas as fórmulas e reproduziu exactamente os valores da v6 (diferença máxima 0) e da v7. O v8 força o recálculo na abertura.
2. **Dívida pública e saldo orçamental sem rastreabilidade.** As séries "Compilação interna BDA" eram valores arredondados (dívida 2025 = 63; saldo 2017 = +1,5 quando o valor oficial foi −5,7). Substituídas pelo FMI WEO de Abril de 2026, coerente com o PIB rebaseado pelo INE (dívida 2025 = 51,3%; MINFIN/UGD 46,6% na definição de dívida governamental; Banco Mundial ≈ 52%).
3. **Mortalidade materna 2025 sem suporte.** A v7 tinha 185 (2024) e 170 (2025). O IIMS 2023-24 (Quadro 17.4) estima 170 por 100 mil (IC 99–242) para os 7 anos anteriores ao inquérito; o valor passa para 2024 e 2025 fica por carry-forward.
4. **Metas 2027 inconsistentes.** 40 metas alteradas para as do PDN 2023-2027 (por exemplo, mortalidade materna 300 → 165, esperança de vida 65 → 63, electrificação 60 → 49, água 70 → 61, desemprego 20 → 25, IPC 35 → 34, crédito ao sector privado 35 → 10). As metas em percentis WGI foram convertidas para a escala de estimativas por deslocamento equivalente na normal padrão a partir de 2022; as metas em % do PIB não petrolífero foram convertidas para % do PIB. A nova coluna "Origem da Meta 2027" documenta cada caso; onde o PDN não fixa meta comparável, a meta operacional foi mantida e identificada.
5. **PIB não petrolífero.** Série recalculada a partir das Contas Nacionais Trimestrais mais recentes do INE (IV trimestre de 2025, com revisões até 0,2 p.p.).
6. **Erro de redacção na Justificação de mínimos e máximos.** Várias escalas 0-100 estavam descritas como "limites 0 e 1"; corrigido pelo gerador da v8.

## 4. Resposta ao feedback Fable 5.1

| Ponto | O que foi feito | Limite |
|---|---|---|
| 1. Emprego: substituir/complementar as séries OIT com dados do INE | Séries do IEA acrescentadas ao catálogo (originais 13.ª CIET; harmonizadas pela OIT; primeira observação da nova metodologia 19.ª–21.ª CIET). Folha `12_Emprego_INE` com subíndice alternativo 2019–2025 que isola o efeito de vintage. A série INE harmonizada regista a queda do desemprego de 13,9% (2024) para 10,4% (2025); a série OIT modelada usada no índice ainda não a incorpora. | O IEA só existe desde 2019: as séries INE não cumprem a cobertura mínima de 80% e não podem substituir as séries OIT no índice 2015–2025 sem quebrar a metodologia. A nova metodologia do INE (IV trim 2025) não tem retropolação. |
| 2. Inclusão: pobreza e protecção social | INC001 fixada na linha de 2,15 USD (referência do PDN, 31,1% em 2018); INC031 nova (3,00 USD PPC 2021, 39,3%); INC032 nova (cobertura do Kwenda: 3,3% dos agregados em 2021 → 14,8% em 2025); INC033 nova (segurados INSS: 1,97 M em 2020 → 3,34 M em 2025; meta PDN 4,3 M). | Nenhum atinge 9 anos de observação; ficam documentados e entram automaticamente quando cumprirem a cobertura. Não existe inquérito de despesas posterior ao IDREA 2018-19. |
| 3. Metas 2027 alinhadas com o PDN | Feito (ponto 3 da secção anterior). | Metas sem correspondência no PDN (inflação, saldo, IPC de exportações, PCI) mantêm valores operacionais identificados. |
| Alerta sobre despesa em educação | Confirmado com as dotações do OGE: 2,0% do PIB (2024), 1,8% (2025), 1,7% (2026); 6,4–6,9% da despesa total, contra o compromisso do PDN de 11,8%. | A série UIS (executada) só chega a 2023. |
| Variável de política em Capital Humano (escolaridade obrigatória) | Documentada como limitação e recomendação; não substituída para não introduzir uma quebra nesta revisão. | Decisão de desenho a tomar pelo GEP. |

## 5. Recolha de dados (webscraping)

O ambiente desta sessão bloqueia o acesso directo a todos os sítios externos (INE, BNA, MINFIN, MINPLAN, Banco Mundial, FMI, OIT, OMS, UNCTAD, Transparency International e outros), tanto por `curl` como por leitura de páginas. Só a pesquisa web funcionou, com um limite de 200 pesquisas por sessão. Foram lançados sete agentes de recolha em paralelo; três concluíram antes de o limite se esgotar:

| Família | Resultados | Principais dados obtidos |
|---|---|---|
| INE | 58 | Contas Nacionais Anuais Preliminares 2025 (PIB +3,13%; petrolífero −5,23%); IEA Anuário 2025 (desemprego 28,3%; informal 78,8%); IEA I e II trim 2026 (21,3%; 21,5%); IPCN Dez-2025 15,70% e Ago-2026 8,78%; Censo 2024 (36,6 M habitantes; alfabetização 72,6%); IIMS 2023-24. |
| MINFIN / FMI | 103 | Dívida governamental 2025 46,6% (UGD) e FMI 51,3% (rebasing do PIB); défice 2025 4,1% do PIB; OGE 2024–2026 educação e saúde; FMI Art. IV 2026 (crescimento 3,1%, reservas 7,4 meses, projecção 2026 2,3%). |
| MINPLAN / INSS / Kwenda | 81 | Balanço do PDN (1 036 indicadores; 856 sem execução no I trim 2025; 66,9% das prioridades); INSS 3,34 M segurados (2025); Kwenda 1,3 M agregados; metas PDN confirmadas; projecções 2026 (OGE 4,17%/13,7%; BNA 3,5%/13,5%; BM 2,4%/14,9%). |
| BNA, Banco Mundial/WGI, agências ONU, índices internacionais | não executadas | Limite de pesquisas esgotado. As séries respectivas mantêm o vintage da v7 (WDI Junho/Julho 2026, WGI 2025, CPI 2025, WEF 2025, UNCTAD, OMS/UNICEF), o mais recente publicado à data de Agosto de 2026. |

Dados locais do repositório usados: INE (Contas Nacionais anuais e trimestrais, IPCN, IEA antiga e nova metodologia, IIMS, Censo 2024, projecções de população), BNA (indicadores externos 1990–2025), FMI WEO (Abril 2026), Banco Mundial WDI (Abril 2026), ILOSTAT (estimativas nacionais harmonizadas), UNCTAD, PDN 2023-2027 (texto e 118 metas) e painel MINPLAN 2025.

Para repetir a recolha nas famílias em falta é necessário permitir o acesso de rede aos domínios oficiais na configuração do ambiente (Network access) ou aumentar o limite de pesquisas (`CLAUDE_CODE_MAX_WEB_SEARCHES_PER_SESSION`).

## 6. O que não foi possível verificar

- Valores 2025 de fontes internacionais publicados depois de Agosto de 2026 (WGI, WUENIC, SOFI 2026, UNCTAD PCI, IGME, ILO Nov-2026): mantidos da v7 e assinalados como tal na Nota (Secção 9.2 e Anexo B).
- Recálculo do livro Excel dentro do ambiente: o LibreOffice instalado não tem o módulo Calc. A cadeia de fórmulas foi validada com o motor Python (`igda_engine.py`), que reproduz exactamente a v6 e a v7; o Excel recalcula tudo ao abrir (se necessário, Ctrl+Alt+F9).
- Série anual completa de segurados do INSS (2015–2025) e de agregados pagos pelo Kwenda (2020): não publicadas em excertos acessíveis.

## 7. Reprodutibilidade

Os scripts usados estão versionados em `Versão Final/Finalíssima/scripts_v8/` (ver `README.md` dessa pasta): `igda_engine.py` (motor de cálculo que replica o construtor), `v8_plan.py` (plano de alterações, fonte única do registo), `build_v8.py`/`build_v8_sheets.py`/`run_build.py` (construção do v8), `gen_nota.py`, `gen_justificacao.py`, `gen_cex.py` (documentos), e os dados intermédios (`labour_data.json`, `analysis.json`, `v8_results.json`, `sweep_partial.json` com as citações da recolha web). A folha `13_Alteracoes_v8` do construtor contém o registo completo, linha a linha, das 66 alterações e das 40 metas revistas.
