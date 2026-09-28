"""
Plano de alterações v7 -> v8 do IGDA_BDA_Construtor.
Tudo o que é parametrizável fica aqui (metas, novas linhas, séries actualizadas, notas), para que build_v8.py seja determinístico
e o relatório de alterações seja gerado a partir da mesma fonte.

Convenções:
  METAS[id] = (meta_2027, origem_texto)           -> escreve coluna I e a nova coluna "Origem da Meta 2027"
  SERIES[id] = {"values": {ano: valor|None}, "origem": str, "fonte": str|None, "codigo": str|None, "nota_add": str, "nacional": 0/1|None}
  NEW_ROWS = lista de dicts com os campos da Base_Potencial
"""
from __future__ import annotations

PDN = "PDN 2023-2027 (Diário da República / AUDA-NEPAD)"

# --------------------------------------------------------------------------------------------
# 1. METAS 2027 ALINHADAS COM O PDN (Feedback Fable 5.1, ponto 3; Nota v7, Secção 11, rec. 3)
# --------------------------------------------------------------------------------------------
METAS = {
    # Governança — PDN p.30 fixa metas em PERCENTIS WGI (2022 -> 2027). Conversão para a escala de estimativas (-2,5..2,5):
    # meta = estimativa 2022 + [Φ⁻¹(p2027) − Φ⁻¹(p2022)] (deslocamento equivalente na normal padrão). Ver Nota Metodológica v8, Secção 8.6.
    "GOV004": (-0.38, "PDN p.30: Estabilidade Política percentil 20,8% (2022) → 25,8% (2027); convertido para estimativa WGI a partir de −0,54 (2022)"),
    "GOV02":  (-0.99, "PDN p.30: Estado de Direito percentil 17,3% → 22,0%; convertido a partir de −1,16 (2022)"),
    "GOV03":  (-0.64, "PDN p.30: Eficiência do Governo percentil 13,0% → 17,5%; convertido a partir de −0,83 (2022)"),
    "GOV05":  (-0.69, "PDN p.30: Qualidade Regulatória percentil 27,4% → 33,5%; convertido a partir de −0,86 (2022)"),
    "GOV06":  (-0.39, "PDN p.30: Controlo da Corrupção percentil 27,9% → 35,3%; convertido a partir de −0,60 (2022)"),
    "GOV008": (-0.90, "PDN p.30: Voz e Participação percentil 24,7% → 29,5%; convertido a partir de −1,05 (2022)"),
    "GOV01":  (34, "PDN p.32: classificação no ranking da Transparência Internacional (0-100) 33 (2022) → 34 (2027) → 42 (2050)"),
    "GOV027": (None, "PDN p.32 fixa meta em posição do ranking RSF (125 → 99), não no score; sem meta comparável"),
    # Macroeconomia — PDN "Grandes números" (p.13) e enquadramento macro (p.16)
    "MAC001": (3.0, "PDN p.13: crescimento real anual médio do PIB de 3,0% em 2023-2027 (trajectória média, não meta anual)"),
    "MAC028": (3.0, "PDN p.13: crescimento real anual médio do PIB de 3,0% em 2023-2027"),
    "MAC014": (4.6, "PDN p.13: crescimento médio anual do PIB não petrolífero de 4,6% em 2023-2027"),
    "DIV017": (4.6, "PDN p.13: crescimento médio anual do PIB não petrolífero de 4,6% em 2023-2027"),
    "MAC004": (60, "PDN p.13/16: dívida pública de 66% (2022) para 60% do PIB (2027)"),
    "MAC016": (60, "PDN p.13/16: dívida pública 66% → 60% do PIB (2027)"),
    "MAC029": (60, "PDN p.13/16: dívida pública 66% → 60% do PIB (2027) — aplicado ao stock externo público como referência"),
    "MAC039": (60, "PDN p.13/16: dívida pública 66% → 60% do PIB (2027) — referência operacional"),
    "MAC007": (6, "PDN p.16: manutenção das reservas internacionais acima de 6 meses de importações"),
    "MAC034": (6, "PDN p.16: reservas internacionais acima de 6 meses de importações"),
    "MAC003": (10, "Sem meta numérica no PDN (compromisso qualitativo de desinflação); meta operacional mantida (um dígito)"),
    "MAC027": (10, "Sem meta numérica no PDN; meta operacional mantida"),
    "MAC030": (10, "Sem meta numérica no PDN; meta operacional mantida"),
    "MAC005": (0, "Sem meta numérica no PDN; meta operacional mantida (equilíbrio orçamental)"),
    "MAC006": (7, "PDN p.163: IDE não petrolífero <1% → 9% do PIB não petrolífero (2027); convertido para ≈7% do PIB (peso não petrolífero ≈80%)"),
    "MAC002": (5000, "Sem meta no PDN (o plano projecta crescimento per capita quase nulo até 2027); meta operacional mantida"),
    # Capital Humano — PDN p.55 (educação) e p.74 (saúde)
    "HUM002": (63, "PDN p.74: esperança média de vida 62 (2022) → 63 (2027) → 68 (2050)"),
    "HUM006": (37, "PDN p.74: taxa de mortalidade de menores de 1 ano 47‰ (2022) → 37‰ (2027) → 14‰ (2050)"),
    "HUM003": (78, "PDN p.55: taxa de alfabetização (≥15 anos) 76% (2022) → 78% (2027) → 90% (2050)"),
    "HUM018": (78, "PDN p.55: alfabetização ≥15 anos 78% (2027); aplicado à alfabetização feminina como referência"),
    "HUM028": (85, "PDN p.55: alfabetização ≥15 anos 78% (2027); a masculina já supera — meta operacional mantida"),
    "HUM004": (70, "PDN p.55: taxa líquida de escolarização no ensino primário 64% (2022) → 70% (2027) → 90% (2050)"),
    "HUM024": (57, "PDN p.74: esperança de vida saudável (HALE) 56 (2022) → 57 (2027) → 60 (2050)"),
    "HUM010": (5, "PDN p.55 fixa o peso da educação na despesa pública (10,0% → 11,8%), não em % do PIB; meta operacional mantida"),
    "HUM030": (10, "Sem meta no PDN; variável de política (duração legal) — manutenção dos 10 anos"),
    "HUM025": (51, "PDN p.74: mortalidade de menores de 5 anos 69‰ (2022) → 51‰ (2027) → 17‰ (2050)"),
    "SAU04":  (51, "PDN p.74: mortalidade de menores de 5 anos 69‰ (2022) → 51‰ (2027) → 17‰ (2050)"),
    # Inclusão Social — PDN p.92
    "INC001": (28, "PDN p.92: população abaixo do limiar de pobreza (<2,15 USD/dia PPC 2017) 31% (2022) → 28% (2027) → 18% (2050)"),
    "INC006": (28, "PDN p.92: pobreza 31% → 28% (2027); aplicado à linha nacional como referência"),
    "INC003": (40, "PDN p.92 fixa a meta em número de segurados (2,5 → 4,3 milhões), não em % de cobertura; meta operacional mantida"),
    "INC005": (4500, "Sem meta no PDN; meta operacional mantida"),
    "INC06":  (35, "Sem meta no PDN; meta operacional mantida"),
    "INC04":  (0.7, "Sem meta no PDN; meta operacional mantida"),
    # Infraestruturas — PDN p.104 (energia) e p.125 (água e saneamento)
    "INF001": (49, "PDN p.104: taxa de electrificação 43% (2022) → 49% (2027) → 72% (2050)"),
    "INF002": (61, "PDN p.125: população que utiliza serviços básicos de água potável 57% (2022) → 61% (2027) → 89% (2050)"),
    "INF003": (55, "PDN p.125: população que utiliza serviços básicos de saneamento 52% (2022) → 55% (2027) → 66% (2050)"),
    "INF016": (73, "PDN p.104: produção de energias renováveis 64% → 73% da capacidade instalada (2027)"),
    "INF006": (30, "Sem meta no PDN (metas em investimento e km reabilitados); meta operacional mantida"),
    "INF05":  (35, "Sem meta no PDN; meta operacional mantida"),
    "INF024": (None, "PDN p.117 fixa a meta portuária em toneladas de carga (17,4 → 29,8 Mt), não em TEU; sem meta comparável"),
    # Mercado de Trabalho — PDN p.13 e p.163
    "LAB002": (25, "PDN p.13: taxa de desemprego 30% (2022) → 25% (2027). Nota: a meta refere-se à definição INE/IEA (alargada); a série OIT modelada usa a definição estrita"),
    "LAB018": (25, "PDN p.13: taxa de desemprego 30% → 25% (definição INE); esta série (harmonizada OIT, definição estrita) não é directamente comparável"),
    "LAB005": (31, "PDN p.163: taxa de formalização 22% (2022) → 31% (2027) → 55% (2050)"),
    "LAB003": (None, "Sem meta no PDN para o desemprego jovem"),
    "LAB016": (65, "Sem meta no PDN; meta operacional mantida"),
    "LAB025": (None, "Sem meta no PDN"),
    "LAB004": (69, "PDN p.163: formalização 22% → 31% (2027) implica emprego informal ≈69%; derivada"),
    # Saúde e Segurança Alimentar — PDN p.74; MINPLAN painel anual 2025
    "SAU003": (165, "PDN p.74: taxa de mortalidade materna 199 (2022) → 165 (2027) → 70 (2050) por 100 mil nascimentos"),
    "SAU029": (165, "PDN p.74: mortalidade materna 199 → 165 (2027)"),
    "SAU004": (4, "PDN p.74: gastos correntes com a saúde 3% (2022) → 4% do PIB (2027) → 7% (2050)"),
    "SAU006": (80, "MINPLAN, Balanço PDN 2025: meta anual de vacinação contra o sarampo 80%; PDN sem meta 2027 explícita"),
    "SAU020": (215, "MINPLAN, Balanço PDN 2025: meta anual de incidência da malária 215 por 1 000 habitantes; PDN sem meta 2027 explícita"),
    "SAU002": (15, "Sem meta numérica no PDN (compromisso 'mais segurança alimentar'); meta operacional mantida"),
    "SAU017": (20, "Sem meta no PDN; meta operacional mantida"),
    # Diversificação — PDN p.163 e p.193
    "DIV015": (10, "PDN p.163: stock de crédito ao sector privado 10,9% → 12,5% do PIB não petrolífero (2027); convertido para ≈10% do PIB (peso não petrolífero ≈80%)"),
    "MAC013": (10, "PDN p.163: crédito ao sector privado 12,5% do PIB não petrolífero (2027) ≈ 10% do PIB"),
    "DIV010": (70, "PDN p.193 fixa exportações não petrolíferas em USD (4,0 → 7,3 mil milhões), não em % das exportações; meta operacional mantida"),
    "DIV007": (12, "Sem meta no PDN para a indústria transformadora; meta operacional mantida"),
    "DIV001": (45, "Sem meta no PDN; meta operacional mantida"),
    "MAC010": (10, "PDN p.193: IDE 6 → 14 mil milhões USD/ano (2027) ≈ 10% do PIB; derivada"),
    "MAC033": (10, "PDN p.193: IDE 6 → 14 mil milhões USD/ano (2027) ≈ 10% do PIB; derivada"),
}

# --------------------------------------------------------------------------------------------
# 2. NOVAS LINHAS DE CANDIDATOS (Feedback pontos 1 e 2)
# --------------------------------------------------------------------------------------------
NEW_ROWS = [
    dict(id="LAB028", dim="6. Mercado de Trabalho", subtema="6. Mercado de Trabalho", indicador="Taxa de emprego 15+ (INE/IEA, harmonizada OIT)",
         unidade="%", sentido="+", minimo=0, maximo=100, meta=None, peso=1, fonte="INE/IEA via OIT (ILOSTAT, LFS)", codigo="ILOSTAT EMP_DWAP_SEX_AGE_RT [BA:13951]",
         origem="OIT ILOSTAT — estimativa nacional (Inquérito sobre o Emprego em Angola, INE), harmonizada", nacional=1, grupo="LAB001", prioridade="Alta",
         values={2019: 62.873, 2020: 62.839, 2021: 63.474, 2022: 64.756, 2023: 62.903, 2024: 64.214, 2025: 66.617},
         nota="v8: série INE/IEA harmonizada pela OIT (definição internacional). Inelegível por cobertura (IEA existe desde 2019); usada na leitura complementar 12_Emprego_INE.",
         meta_origem="Sem meta no PDN"),
    dict(id="LAB029", dim="6. Mercado de Trabalho", subtema="6. Mercado de Trabalho", indicador="Desemprego jovem 15-24 (INE/IEA, harmonizada OIT)",
         unidade="%", sentido="-", minimo=0, maximo=50, meta=None, peso=1, fonte="INE/IEA via OIT (ILOSTAT, LFS)", codigo="ILOSTAT UNE_DEAP_SEX_AGE_RT Y15-24 [BA:13951]",
         origem="OIT ILOSTAT — estimativa nacional (IEA/INE), harmonizada", nacional=1, grupo="G_DESEMP_JUVENIL", prioridade="Alta",
         values={2019: 31.107, 2020: 29.081, 2021: 29.789, 2022: 26.585, 2023: 21.469, 2024: 26.203, 2025: 18.562},
         nota="v8: série INE/IEA harmonizada pela OIT. Inelegível por cobertura; usada em 12_Emprego_INE.", meta_origem="Sem meta no PDN"),
    dict(id="LAB030", dim="6. Mercado de Trabalho", subtema="6. Mercado de Trabalho", indicador="Taxa de desemprego 15+ (INE/IEA, 13.ª CIET, definição nacional alargada)",
         unidade="%", sentido="-", minimo=0, maximo=60, meta=25, peso=1, fonte="INE — Inquérito sobre o Emprego em Angola (IEA)", codigo="INE_IEA_TXDESEMP_13CIET",
         origem="INE — séries cronológicas IEA (antiga metodologia, 13.ª CIET); média dos trimestres disponíveis por ano", nacional=1, grupo="G_DESEMPREGO", prioridade="Alta",
         values={2019: 30.21, 2020: 32.18, 2021: 32.32, 2022: 30.16, 2023: 31.85, 2024: 31.47, 2025: 28.37},
         nota="v8: definição alargada (inclui desencorajados) — nível ~2x superior à definição estrita OIT. 2019 = média II-IV trim; 2023 = apenas IV trim publicado; 2025 = média I-III trim (a série termina no III trim 2025 com a mudança de metodologia). Inelegível por cobertura; usada em 12_Emprego_INE.",
         meta_origem="PDN p.13: taxa de desemprego 30% (2022) → 25% (2027), na definição INE"),
    dict(id="LAB031", dim="6. Mercado de Trabalho", subtema="6. Mercado de Trabalho", indicador="Taxa de desemprego 15+ (INE/IEA, 19.ª-21.ª CIET, nova metodologia)",
         unidade="%", sentido="-", minimo=0, maximo=60, meta=25, peso=1, fonte="INE — IEA (nova metodologia)", codigo="INE_IEA_TXDESEMP_19CIET",
         origem="INE — IEA IV trimestre 2025 (primeira publicação com as resoluções da 19.ª, 20.ª e 21.ª CIET)", nacional=1, grupo="G_DESEMPREGO", prioridade="Alta",
         values={2025: 20.12},
         nota="v8: quebra de série — nova metodologia desde o IV trim 2025 (20,1%; emprego 39,6%; informalidade 78,6%; desemprego jovem 43,3%). Série a acumular a partir de 2026.",
         meta_origem="PDN p.13: taxa de desemprego 30% → 25% (2027)"),
    dict(id="INC031", dim="4. Inclusão Social e Protecção", subtema="4. Inclusão Social", indicador="Taxa de pobreza (< 3,00 USD/dia, PPC 2021)",
         unidade="%", sentido="-", minimo=0, maximo=80, meta=None, peso=1, fonte="Banco Mundial — PIP/WDI", codigo="SI.POV.DDAY (linha 3,00 USD PPC 2021)",
         origem="Banco Mundial — WDI (actualização Abr-2026): linha internacional de pobreza revista para 3,00 USD/dia (PPC 2021)", nacional=0, grupo="G_POBREZA", prioridade="Alta",
         values={2018: 39.3},
         nota="v8: novo padrão do Banco Mundial (Junho 2025). A meta do PDN (28%) refere-se à linha antiga de 2,15 USD (PPC 2017), mantida em INC001. Sem série anual: um inquérito (IDREA 2018-19) no período.",
         meta_origem="PDN p.92 fixa a meta na linha 2,15 USD (PPC 2017); sem meta comparável nesta linha"),
]

# --------------------------------------------------------------------------------------------
# 3. SÉRIES ACTUALIZADAS A PARTIR DE FONTES LOCAIS (ficheiros INE/BNA/FMI do repositório)
# --------------------------------------------------------------------------------------------
# PIB não petrolífero: recalculado das Contas Nacionais Trimestrais do INE (medidas de volume encadeadas, com ajuste sazonal),
# ficheiro descarregado em 2026-05-09 (inclui IV trim 2025 e revisões). Soma anual dos 4 trimestres; crescimento = variação da soma.
NONOIL = {2015: -3.719, 2016: 0.792, 2017: 3.862, 2018: 3.119, 2019: 1.263, 2020: -4.809, 2021: 5.471, 2022: 3.368, 2023: 2.539, 2024: 5.313, 2025: 5.247}
GDP_INE_Q = {2015: 0.740, 2016: -0.323, 2017: -0.206, 2018: -0.563, 2019: -1.048, 2020: -5.301, 2021: 1.200, 2022: 3.731, 2023: 1.431, 2024: 4.889, 2025: 3.029}

# FMI WEO (vintage Abril 2026, cache local imf_weo_ago.xml): dívida bruta e saldo do governo geral, % PIB
IMF_DEBT = {2015: 50.39, 2016: 65.69, 2017: 59.61, 2018: 81.64, 2019: 100.80, 2020: 119.84, 2021: 75.50, 2022: 57.43, 2023: 75.74, 2024: 57.11, 2025: 51.31}
IMF_BAL = {2015: -2.58, 2016: -3.93, 2017: -5.68, 2018: 2.00, 2019: -0.20, 2020: -3.05, 2021: 1.35, 2022: 1.78, 2023: -2.49, 2024: -1.19, 2025: -4.07}
IMF_PRIM = {2015: -1.01, 2016: -1.47, 2017: -2.57, 2018: 6.08, 2019: 4.45, 2020: 2.75, 2021: 5.79, 2022: 5.23, 2023: 2.44, 2024: 3.40, 2025: -0.39}

# INE IPCN: inflação média anual (idêntica ao WDI) e homóloga de Dezembro
INE_CPI_AVG = {2015: None, 2016: 30.69, 2017: 29.84, 2018: 19.63, 2019: 17.08, 2020: 22.27, 2021: 25.75, 2022: 21.36, 2023: 13.64, 2024: 28.24, 2025: 20.16}
INE_CPI_DEC = {2015: 12.09, 2016: 41.12, 2017: 23.67, 2018: 18.60, 2019: 16.90, 2020: 25.10, 2021: 27.03, 2022: 13.86, 2023: 20.01, 2024: 27.50, 2025: 15.70}

SERIES = {
    "MAC014": dict(values=NONOIL, origem="INE — Contas Nacionais Trimestrais (medidas de volume encadeadas, c/ ajuste sazonal), ficheiro 2026-05-09; soma anual",
                   nota_add="v8: série recalculada a partir da última publicação trimestral do INE (inclui IV trim 2025 e revisões); diferenças ≤0,2 p.p. face à v7."),
    "DIV017": dict(values=NONOIL, origem="INE — Contas Nacionais Trimestrais (medidas de volume encadeadas, c/ ajuste sazonal), ficheiro 2026-05-09; soma anual",
                   nota_add="v8: série recalculada a partir da última publicação trimestral do INE (inclui IV trim 2025 e revisões)."),
    "MAC004": dict(values=IMF_DEBT, origem="FMI — World Economic Outlook, Abril 2026 (dívida bruta do governo geral, % PIB)", fonte="MINFIN/FMI WEO",
                   nota_add="v8: substitui a série 'Compilação interna BDA' (valores arredondados, sem rastreabilidade) pela série oficial do FMI WEO Abr-2026 (2025: 51,3%). Alinha-se com a leitura do PDN (134% em 2020 → 65% em 2022) e do Banco Mundial (~52% em 2025)."),
    "MAC005": dict(values=IMF_BAL, origem="FMI — World Economic Outlook, Abril 2026 (saldo global do governo geral, % PIB)", fonte="MINFIN/FMI WEO",
                   nota_add="v8: substitui a série 'Compilação interna BDA' (ex.: 2017 = +1,5 quando o saldo oficial foi −5,7) pela série FMI WEO Abr-2026."),
    "MAC016": dict(values=IMF_DEBT, origem="FMI — World Economic Outlook, Abril 2026", nota_add="v8: vintage Abr-2026 (2025: 51,3%)."),
    "MAC019": dict(values=IMF_BAL, origem="FMI — World Economic Outlook, Abril 2026", nota_add="v8: vintage Abr-2026."),
    "MAC020": dict(values=IMF_PRIM, origem="FMI — World Economic Outlook, Abril 2026", nota_add="v8: vintage Abr-2026."),
    "MAC003": dict(values=None, origem="INE — IPC Nacional (média anual; idêntica ao WDI FP.CPI.TOTL.ZG)", fonte="INE",
                   nota_add="v8: validado com o IPCN do INE (2025: média anual 20,16%; homóloga de Dezembro 15,70%)."),
    "MAC027": dict(values=INE_CPI_DEC, origem="INE — IPC Nacional, variação homóloga de Dezembro (ficheiro 2026-05-09)", nota_add="v8: 2025 = 15,70% (Dez/2025)."),
}

# --------------------------------------------------------------------------------------------
# 4. RENOMEAÇÕES / CORRECÇÕES DE METADADOS
# --------------------------------------------------------------------------------------------
RENAME = {
    "INC001": dict(indicador="Taxa de pobreza (< 2,15 USD/dia, PPC 2017)", codigo="SI.POV.DDAY (linha 2,15 USD PPC 2017, arquivo PIP)",
                   nota_add="v8: a linha de 2,15 USD (PPC 2017) é a referência do PDN (31% em 2022; 28% em 2027). Valor 2018 (IDREA 2018-19) = 31,1% (Banco Mundial/PIP). A nova linha de 3,00 USD (PPC 2021) consta em INC031.",
                   values={2018: 31.1}),
}


# --------------------------------------------------------------------------------------------
# 5. ACTUALIZAÇÕES RESULTANTES DA RECOLHA WEB (INE, MINFIN, FMI) — ronda 1
# --------------------------------------------------------------------------------------------
SERIES.update({
    "MAC001": dict(values={2025: 3.13}, origem="INE — Contas Nacionais Anuais Preliminares 2025 (05-05-2026): PIB real +3,13%; 2015-2024 CN anuais definitivas",
                   nota_add="v8: 2025 confirmado pelas Contas Nacionais Anuais Preliminares do INE (+3,13%; sector petrolífero −5,23%). Coincide com o FMI (Art. IV 2026: 3,1%) e com o WDI."),
    "LAB004": dict(values={2025: 78.8}, origem="INE — IEA, Anuário 2025 (Abril 2026), metodologia 13.ª CIET; 2019-2024 séries cronológicas IEA",
                   nota_add="v8: 2025 = 78,8% (Anuário IEA 2025, valor anual; a v7 usava o III trim). Quebra de série a partir do IV trim 2025 (nova metodologia: 78,6%)."),
    "LAB005": dict(values={2025: 21.2}, origem="INE — IEA (taxa de emprego formal = 100 − informal), Anuário 2025",
                   nota_add="v8: 2025 = 21,2% (complemento da taxa de emprego informal anual 78,8%, Anuário IEA 2025)."),
    "SAU003": dict(values={2024: 170, 2025: None}, origem="Banco Mundial — WDI (estimativas modeladas MMEIG) até 2023; 2024 = IIMS 2023-24 (estimativa directa, 7 anos anteriores ao inquérito)",
                   nota_add="v8: correcção — a v7 registava 185 (2024) e 170 (2025) sem suporte documental; o IIMS 2023-24 (Quadro 17.4) estima a RMM em 170 por 100 mil (IC 99-242) para os 7 anos anteriores ao inquérito, colocada em 2024; 2025 passa a carry-forward."),
    "HUM010": dict(values=None, origem="Banco Mundial/UNESCO — WDI (despesa executada, UIS) até 2023",
                   nota_add="v8: sem observação UIS para 2024-2025. As dotações do OGE confirmam a trajectória descendente sinalizada no feedback: 2,0% do PIB (OGE 2024), 1,8% (OGE 2025), 1,7% (OGE 2026), 6,4-6,9% da despesa total (UNICEF, Budget Briefs)."),
    "SAU004": dict(values=None, origem="Banco Mundial/OMS — GHED (despesa corrente em saúde) até 2023",
                   nota_add="v8: sem observação GHED para 2024-2025. Dotações OGE para a saúde: 5,5% do OGE (2024), 5,7-6,3% (2025), 6,32% do OGE e 1,5% do PIB (2026) — UNICEF Budget Briefs."),
    "INC001": dict(values=None, nota_add="v8: sem novo inquérito de despesas desde o IDREA 2018-19; o PDN toma 31% (2022*) como base."),
})
NEW_ROWS[2]["values"][2025] = 28.3   # LAB030: Anuário IEA 2025 (valor anual oficial, metodologia 13.ª CIET)
NEW_ROWS[2]["nota"] = ("v8: definição alargada (inclui desencorajados) — nível ~2x superior à definição estrita OIT. 2019 = média II-IV trim; 2023 = apenas IV trim publicado; "
                       "2025 = 28,3% (Anuário IEA 2025, Abril 2026). A série termina em 2025 com a mudança de metodologia. Inelegível por cobertura; usada em 12_Emprego_INE.")

# Kwenda — beneficiários acumulados (agregados familiares), MINPLAN painel PDN; cobertura = % dos agregados familiares (Censo 2024: 9 110 616 agregados)
KWENDA_AGREGADOS = {2022: 610832, 2023: 951203, 2024: 1070037, 2025: 1350850}
HOUSEHOLDS_2024 = 9110616
NEW_ROWS.append(dict(id="INC032", dim="4. Inclusão Social e Protecção", subtema="4. Inclusão Social", indicador="Cobertura do Kwenda (agregados beneficiários acumulados, % dos agregados familiares)",
    unidade="%", sentido="+", minimo=0, maximo=50, meta=19.8, peso=1, fonte="MINPLAN/MASFAMU — Balanço do PDN (painel anual 2025); INE Censo 2024 (agregados)", codigo="MINPLAN_PDN_KWENDA_COV",
    origem="MINPLAN — painel 'Implementation - Key Indicators by Domain' (beneficiários de transferências sociais monetárias, cumulativo) ÷ agregados familiares do Censo 2024 (9 110 616)", nacional=1, grupo="G_PROTECCAO_SOCIAL", prioridade="Alta",
    values={y: round(100 * v / HOUSEHOLDS_2024, 2) for y, v in KWENDA_AGREGADOS.items()},
    nota="v8: novo indicador de protecção social não contributiva (feedback ponto 2). Denominador fixo (Censo 2024) para isolar a expansão do programa. Meta 2027 = 1,8 milhões de agregados (meta anual MINPLAN 2025) ÷ 9,11 M. Inelegível por cobertura (programa iniciado em 2020); documentado para incorporação futura.",
    meta_origem="MINPLAN, Balanço PDN 2025: meta de 1 800 000 agregados beneficiários; PDN p.92 não fixa meta em % de cobertura"))


# --------------------------------------------------------------------------------------------
# 6. ACTUALIZAÇÕES RESULTANTES DA RECOLHA WEB (MINPLAN / MAPTSS / INSS / FAS-Kwenda) — ronda 1
# --------------------------------------------------------------------------------------------
# INSS — segurados inscritos (stock, milhões). Fontes: INSS/MAPTSS (Boletim Anual da Protecção Social Obrigatória 2025; Forbes África Lusófona 27-01-2025; Expansão 2026)
INSS_SEGURADOS_M = {2020: 1.97, 2022: 2.13, 2024: 3.004238, 2025: 3.341475}
NEW_ROWS.append(dict(id="INC033", dim="4. Inclusão Social e Protecção", subtema="4. Inclusão Social", indicador="Segurados inscritos na protecção social obrigatória (INSS, milhões)",
    unidade="milhões", sentido="+", minimo=0, maximo=8, meta=4.3, peso=1, fonte="INSS/MAPTSS — Boletim Anual da Protecção Social Obrigatória", codigo="MAPTSS_INSS_SEGURADOS",
    origem="INSS/MAPTSS: 2024 = 3 004 238 (Balanço 2024); 2025 = 3 341 475 (Boletim Anual 2025); 2020 = 1,97 M e Fev-2022 = 2,13 M (INSS via Expansão)", nacional=1, grupo="G_PROTECCAO_SOCIAL_CONTRIB", prioridade="Alta",
    values=INSS_SEGURADOS_M,
    nota="v8: indicador de protecção social contributiva (PDN, Programa 22: segurados 2,5 → 4,3 milhões). Stock acumulado de inscritos (inclui inactivos); 2022 refere-se a Fevereiro. Inelegível por cobertura (4 observações); série anual completa a obter junto do INSS.",
    meta_origem="PDN p.92: número de segurados registados na protecção social obrigatória 2,5 (2022) → 4,3 milhões (2027) → 13,6 (2050)"))
# Kwenda — 2021: pouco mais de 300 mil agregados pagos (FAS/IDL via Expansão)
for row in NEW_ROWS:
    if row["id"] == "INC032":
        row["values"][2021] = round(100 * 300000 / HOUSEHOLDS_2024, 2)
        row["nota"] += " 2021 ≈ 300 mil agregados pagos (FAS/IDL); 2022-2025 = painel MINPLAN (610 832; 951 203; 1 070 037; 1 350 850)."

# Contexto para os documentos (não altera o construtor)
CONTEXTO = dict(
    pdn_indicadores_total=1036, pdn_indicadores_sem_execucao_2025T1=856, pdn_prioridades_materializadas_pct=66.9, pdn_projectos_sem_execucao_2024=846,
    inss_novos_segurados_2025=337237, inss_empregos_formais_2025=218669, kwenda_ii_meta=1500000, kwenda_ii_financiamento_musd=400,
    proj_2026={"OGE 2026 (MINFIN)": {"pib": 4.17, "inflacao": 13.7}, "FMI Art. IV 2026": {"pib": 2.3, "inflacao": 12.9}, "BNA": {"pib": 3.5, "inflacao": 13.5}, "Banco Mundial": {"pib": 2.4, "inflacao": 14.9}},
    inflacao_2026={"Mar": 12.42, "Mai": 10.88, "Jun": 10.11, "Jul": 9.33, "Ago": 8.78},
    iea_2026={"I trim": {"desemprego": 21.3, "emprego": 42.3, "jovem": 40.7, "informal": 79.3, "forca_trabalho": 53.8}, "II trim": {"desemprego": 21.5, "emprego": 44.2, "jovem": 40.5, "informal": 80.1, "forca_trabalho": 56.3}},
    divida_ugd_2025=46.59, divida_fmi_2025=51.31, divida_wb_2025=52, saldo_2025_minfin=-4.1, reservas_meses_fmi_2025=7.4,
    censo2024={"populacao": 36604681, "alfabetizacao_15+": 72.6, "electricidade_rede_agregados": 48.6, "agua_canalizada_agregados": 37.0, "agregados": 9110616},
)
