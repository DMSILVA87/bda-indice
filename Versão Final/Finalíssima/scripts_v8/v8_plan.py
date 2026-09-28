"""
Plano de alterações v7 -> v8 do IGDA_BDA_Construtor.
Tudo o que é parametrizável fica aqui (metas, novas linhas, séries actualizadas, notas), para que build_v8.py seja determinístico
e o registo de alterações (13_Alteracoes_v8, Nota v8 Anexo C, relatório) seja gerado a partir da mesma fonte.

Convenções:
  METAS[id] = (meta_2027, origem_texto, categoria)
      categoria ∈ {"pdn", "transposta", "convertida", "minplan", "operacional", "sem_meta"}
        pdn         — valor quantificado do PDN aplicado directamente (base do PDN coincide com a série, ou o indicador é o do próprio PDN)
        transposta  — a base 2022 do PDN não coincide com o valor 2022 da série (fonte/definição distinta): aplica-se ao valor
                      2022 da série a VARIAÇÃO do PDN (aditiva para níveis e proporções; relativa para taxas de mortalidade,
                      de desemprego e rácios de dívida). Especificada em TRANSPOR e calculada em build_v8.py.
        convertida  — mudança de escala/unidade (percentis WGI → estimativas; % do PIB não petrolífero → % do PIB)
        minplan     — meta anual 2025 do Balanço do PDN (MINPLAN) usada como proxy da meta 2027
        operacional — meta operacional da v7 mantida ou ajustada (sem meta comparável no PDN)
        sem_meta    — sem meta (removida ou inexistente), porque o conceito do PDN difere do da série
  SERIES[id] = {"values": {ano: valor|None}|None, "origem", "fonte", "codigo", "url", "nota_add", "nota_sub": [(velho, novo)], "nacional", "tipo"}
  NEW_ROWS = lista de dicts com os campos da Base_Potencial
"""
from __future__ import annotations
from statistics import NormalDist

PDN = "PDN 2023-2027 (Diário da República / AUDA-NEPAD)"
WEO_URL = "https://www.imf.org/en/Publications/WEO/weo-database/2026/april"
PIP_URL = "https://pip.worldbank.org/country-profiles/AGO"

# --------------------------------------------------------------------------------------------
# 1. METAS 2027 ALINHADAS COM O PDN (Feedback Fable 5.1, ponto 3; Nota v7, Secção 11, rec. 3)
# --------------------------------------------------------------------------------------------
# 1.1 Governança — o PDN (p.30) fixa metas em PERCENTIS WGI (2022 -> 2027). Conversão para a escala de estimativas (-2,5..2,5):
#     meta = ê2022 + [Φ⁻¹(p2027) − Φ⁻¹(p2022)]   (deslocamento equivalente na normal padrão, aplicado à estimativa 2022 da série)
#     Os percentile ranks do WGI são posições empíricas e de vintage anterior, pelo que Φ(ê2022) ≠ p2022; o deslocamento em z é
#     invariante a essa diferença de nível. Documentado na Nota v8 (8.6) e em 13_Alteracoes_v8.
WGI_PDN = {  # id: (designação, p2022 %, p2027 %, ê2022 na série)
    "GOV004": ("Estabilidade Política e Ausência de Violência", 20.8, 25.8, -0.5414),
    "GOV02": ("Estado de Direito", 17.3, 22.0, -1.1625),
    "GOV03": ("Eficiência do Governo", 13.0, 17.5, -0.8326),
    "GOV05": ("Qualidade Regulatória", 27.4, 33.5, -0.8619),
    "GOV06": ("Controlo da Corrupção", 27.9, 35.3, -0.6032),
    "GOV008": ("Voz e Participação", 24.7, 29.5, -1.0481),
}
_N = NormalDist()


def wgi_meta(e2022, p2022, p2027):
    return round(e2022 + (_N.inv_cdf(p2027 / 100) - _N.inv_cdf(p2022 / 100)), 2)


def wgi_meta_pp(e2022, p2022, p2027):
    """Sensibilidade: deslocamento em pontos percentuais aplicado ao percentil implícito Φ(ê2022)."""
    p_impl = _N.cdf(e2022) * 100
    return round(_N.inv_cdf(min(99.9, max(0.1, p_impl + (p2027 - p2022))) / 100), 2)


METAS = {}
for _id, (_nome, _p22, _p27, _e22) in WGI_PDN.items():
    METAS[_id] = (wgi_meta(_e22, _p22, _p27),
                  f"PDN p.30: {_nome} percentil WGI {str(_p22).replace('.', ',')}% (2022) → {str(_p27).replace('.', ',')}% (2027); convertido para a escala de estimativas "
                  f"a partir de {str(_e22).replace('.', ',')} (2022): meta = ê2022 + [Φ⁻¹(p2027) − Φ⁻¹(p2022)]", "convertida")

METAS.update({
    "GOV01": (34, "PDN p.32: classificação no ranking da Transparência Internacional (0-100) 33 (2022) → 34 (2027) → 42 (2050); a série regista 33 em 2022 (base coincide)", "pdn"),
    "GOV027": (None, "PDN p.32 fixa meta em posição do ranking RSF (125 → 99), não no score; sem meta comparável", "sem_meta"),
    # Macroeconomia — PDN "Grandes números" (p.13), enquadramento macro (p.16) e Quadro Visão 2027 (p.24)
    "MAC001": (3.0, "PDN p.13: crescimento real anual médio do PIB de 3,0% em 2023-2027 (trajectória média, não meta anual)", "pdn"),
    "MAC028": (3.0, "PDN p.13: crescimento real anual médio do PIB de 3,0% em 2023-2027 (trajectória média)", "pdn"),
    "MAC014": (4.6, "PDN p.13: crescimento médio anual do PIB não petrolífero de 4,6% em 2023-2027 (trajectória média)", "pdn"),
    "DIV017": (4.6, "PDN p.13: crescimento médio anual do PIB não petrolífero de 4,6% em 2023-2027 (trajectória média)", "pdn"),
    "MAC029": (None, "PDN p.24 fixa meta para a dívida pública total (66% → 60% do PIB), não para o stock de dívida externa pública; sem meta comparável (a v7 usava 60 como referência)", "sem_meta"),
    "MAC039": (None, "PDN p.24 fixa meta para a dívida pública total (66% → 60% do PIB), não para a dívida externa total; sem meta comparável (a v7 usava 60 como referência)", "sem_meta"),
    "MAC007": (6, "PDN p.16: manutenção das reservas internacionais acima de 6 meses de importações", "pdn"),
    "MAC034": (6, "PDN p.16: reservas internacionais acima de 6 meses de importações", "pdn"),
    "MAC003": (10, "Sem meta numérica no PDN (compromisso qualitativo de desinflação); meta operacional mantida (um dígito)", "operacional"),
    "MAC027": (10, "Sem meta numérica no PDN; meta operacional mantida", "operacional"),
    "MAC030": (10, "Sem meta numérica no PDN; meta operacional mantida", "operacional"),
    "MAC005": (0, "Sem meta numérica no PDN; meta operacional mantida (equilíbrio orçamental)", "operacional"),
    "MAC006": (7, "PDN p.163: IDE não petrolífero <1% → 9% do PIB não petrolífero (2027); convertido para ≈7% do PIB (peso não petrolífero ≈80%)", "convertida"),
    "MAC002": (5000, "Sem meta no PDN (o plano projecta crescimento per capita quase nulo até 2027); meta operacional mantida", "operacional"),
    "MAC010": (None, "PDN p.193 fixa a entrada BRUTA de IDE em USD (6 → 14 mil milhões/ano, ≈10% do PIB); a série mede o IDE LÍQUIDO em % do PIB (negativo em 2016-2024): conceito diferente, sem meta comparável (a v7 usava 5)", "sem_meta"),
    "MAC033": (None, "PDN p.193 fixa a entrada BRUTA de IDE em USD (6 → 14 mil milhões/ano); a série (BNA) mede o IDE LÍQUIDO em % do PIB: conceito diferente, sem meta comparável (a v7 usava 5)", "sem_meta"),
    # Capital Humano — PDN p.55 (educação) e p.74 (saúde)
    "HUM028": (85, "PDN p.55 fixa a meta da alfabetização total (76% → 78%); a alfabetização masculina (84% em 2023) já a supera — meta operacional mantida", "operacional"),
    "HUM018": (None, "PDN p.55 fixa meta apenas para a alfabetização total (76% → 78%); a série feminina (54% em 2023) mede um conceito distinto — sem meta comparável (a v7 usava 75)", "sem_meta"),
    "HUM004": (70, "PDN p.55: taxa líquida de escolarização no ensino primário 64% (2022) → 70% (2027) → 90% (2050); a série (WDI) não tem observações em 2015-2025, pelo que a base não pôde ser confrontada", "pdn"),
    "HUM010": (5, "PDN p.55 fixa o peso da educação na despesa pública (10,0% → 11,8%), não em % do PIB; meta operacional mantida", "operacional"),
    "HUM030": (10, "Sem meta no PDN; variável de política (duração legal da escolaridade obrigatória) — meta operacional ajustada para os 10 anos em vigor", "operacional"),
    # Inclusão Social — PDN p.92
    "INC001": (28, "PDN p.92: população abaixo do limiar de pobreza (<2,15 USD/dia PPC 2017) 31% (2022*) → 28% (2027) → 18% (2050); a série regista 31,1% em 2018 (base coincide)", "pdn"),
    "INC003": (40, "PDN p.92 fixa a meta em número de segurados (2,5 → 4,3 milhões, ver INC033), não em % de cobertura; meta operacional mantida", "operacional"),
    "INC005": (4500, "Sem meta no PDN; meta operacional mantida", "operacional"),
    "INC06": (35, "Sem meta no PDN; meta operacional mantida", "operacional"),
    "INC04": (0.7, "Sem meta no PDN; meta operacional mantida", "operacional"),
    # Infraestruturas — PDN p.104 (energia) e p.125 (água e saneamento)
    "INF016": (None, "PDN p.104 fixa a meta em % da CAPACIDADE INSTALADA renovável (64% → 73%); a série mede a % da PRODUÇÃO eléctrica (91% em 2021): conceito diferente, sem meta comparável (a v7 usava 70)", "sem_meta"),
    "INF006": (30, "Sem meta no PDN (metas em investimento e km reabilitados); meta operacional mantida", "operacional"),
    "INF05": (35, "Sem meta no PDN; meta operacional mantida", "operacional"),
    "INF024": (None, "PDN p.117 fixa a meta portuária em toneladas de carga (17,4 → 29,8 Mt), não em TEU; sem meta comparável", "sem_meta"),
    # Mercado de Trabalho — PDN p.13 e p.163
    "LAB003": (None, "Sem meta no PDN para o desemprego jovem", "sem_meta"),
    "LAB016": (65, "Sem meta no PDN; meta operacional mantida", "operacional"),
    "LAB025": (None, "Sem meta no PDN", "sem_meta"),
    # Saúde e Segurança Alimentar — PDN p.74; MINPLAN painel anual 2025
    "SAU006": (80, "MINPLAN, Balanço PDN 2025: meta anual de vacinação contra o sarampo 80% (usada como proxy da meta 2027; o PDN não fixa meta 2027 explícita)", "minplan"),
    "SAU020": (215, "MINPLAN, Balanço PDN 2025: meta anual de incidência da malária 215 por 1 000 habitantes (usada como proxy da meta 2027; o PDN não fixa meta 2027 explícita)", "minplan"),
    "SAU002": (15, "Sem meta numérica no PDN (compromisso 'mais segurança alimentar'); meta operacional mantida", "operacional"),
    "SAU017": (20, "Sem meta no PDN; meta operacional mantida", "operacional"),
    # Diversificação — PDN p.163 e p.193
    "DIV015": (10, "PDN p.163: stock de crédito ao sector privado 10,9% → 12,5% do PIB não petrolífero (2027); convertido para ≈10% do PIB (peso não petrolífero ≈80%)", "convertida"),
    "MAC013": (10, "PDN p.163: crédito ao sector privado 12,5% do PIB não petrolífero (2027); convertido para ≈10% do PIB (peso não petrolífero ≈80%)", "convertida"),
    "DIV010": (70, "PDN p.193 fixa exportações não petrolíferas em USD (4,0 → 7,3 mil milhões), não em % das exportações; meta operacional mantida", "operacional"),
    "DIV007": (12, "Sem meta no PDN para a indústria transformadora; meta operacional mantida", "operacional"),
    "DIV001": (45, "Sem meta no PDN; meta operacional mantida", "operacional"),
})

# 1.2 Metas TRANSPOSTAS — a base 2022 do PDN não coincide com o valor 2022 da série (fonte ou definição distinta).
#     Regra: meta_série = valor2022_série + (meta_PDN − base_PDN)  [modo "add", níveis e proporções]
#            meta_série = valor2022_série × (meta_PDN ÷ base_PDN)  [modo "rel", taxas de mortalidade/desemprego e rácios de dívida]
#     Se a série não tem observação em 2022, usa-se a observação mais próxima (identificada no texto da origem).
#     "complemento": meta = 100 − meta de outra linha.
TRANSPOR = {
    "HUM002": dict(pagina="PDN p.74", conceito="esperança média de vida", base=62, meta=63, modo="add", fonte_serie="WDI/UN WPP", nd=1),
    "HUM006": dict(pagina="PDN p.74", conceito="mortalidade de menores de 1 ano (‰)", base=47, meta=37, modo="rel", fonte_serie="IGME/WDI", nd=1),
    "HUM003": dict(pagina="PDN p.55", conceito="taxa de alfabetização ≥15 anos (%)", base=76, meta=78, modo="add", fonte_serie="UNESCO-UIS/WDI", nd=1),
    "HUM024": dict(pagina="PDN p.74", conceito="esperança de vida saudável, HALE (anos)", base=56, meta=57, modo="add", fonte_serie="OMS GHO", nd=1),
    "HUM025": dict(pagina="PDN p.74", conceito="mortalidade de menores de 5 anos (‰)", base=69, meta=51, modo="rel", fonte_serie="IIMS 2023-24 (estimativa directa)", nd=1),
    "SAU04": dict(pagina="PDN p.74", conceito="mortalidade de menores de 5 anos (‰)", base=69, meta=51, modo="rel", fonte_serie="IGME/WDI", nd=1),
    "INC006": dict(pagina="PDN p.92", conceito="pobreza <2,15 USD/dia (%)", base=31, meta=28, modo="rel", fonte_serie="linha de pobreza nacional, IDREA 2018-19", nd=1,
                   extra="transposição por analogia: a linha nacional mede um conceito próximo mas distinto da linha internacional"),
    "INF001": dict(pagina="PDN p.104", conceito="taxa de electrificação (%)", base=43, meta=49, modo="add", fonte_serie="WDI (acesso à electricidade por qualquer fonte)", nd=1,
                   extra="o PDN mede a electrificação pela rede (on-grid); o WDI mede o acesso por qualquer fonte"),
    "INF002": dict(pagina="PDN p.125", conceito="população que utiliza serviços básicos de água potável (%)", base=57, meta=61, modo="add", fonte_serie="OMS/UNICEF JMP via WDI", nd=1),
    "INF003": dict(pagina="PDN p.125", conceito="população que utiliza serviços básicos de saneamento (%)", base=52, meta=55, modo="add", fonte_serie="OMS/UNICEF JMP via WDI", nd=1),
    "LAB002": dict(pagina="PDN p.13", conceito="taxa de desemprego, definição INE/IEA alargada (%)", base=30, meta=25, modo="rel", fonte_serie="OIT modelada, definição internacional estrita", nd=1,
                   extra="a meta do PDN aplica-se directamente às séries INE (LAB030/LAB031)"),
    "LAB018": dict(pagina="PDN p.13", conceito="taxa de desemprego, definição INE/IEA alargada (%)", base=30, meta=25, modo="rel", fonte_serie="INE/IEA harmonizada pela OIT, definição estrita", nd=1),
    "LAB005": dict(pagina="PDN p.163", conceito="taxa de formalização do emprego (%)", base=22, meta=31, modo="add", fonte_serie="INE/IEA, emprego formal", nd=1),
    "LAB004": dict(pagina="PDN p.163", conceito="emprego informal (%)", modo="complemento", ref="LAB005", fonte_serie="INE/IEA", nd=1),
    "SAU003": dict(pagina="PDN p.74", conceito="razão de mortalidade materna (por 100 mil nascimentos)", base=199, meta=165, modo="rel", fonte_serie="MMEIG/WDI", nd=0),
    "SAU029": dict(pagina="PDN p.74", conceito="razão de mortalidade materna (por 100 mil nascimentos)", base=199, meta=165, modo="rel", fonte_serie="IIMS 2023-24 (estimativa directa)", nd=0),
    "SAU004": dict(pagina="PDN p.74", conceito="despesa corrente em saúde (% do PIB)", base=3, meta=4, modo="add", fonte_serie="OMS GHED via WDI", nd=1),
    "MAC004": dict(pagina="PDN p.24 (Quadro Visão 2027; p.14/16: 134% em 2020 → 65% em 2022)", conceito="dívida pública (% do PIB), perímetro MINFIN 'dívida pública' e PIB pré-rebasing", base=66, meta=60, modo="rel",
                   fonte_serie="FMI WEO Abr-2026, governo geral, PIB rebaseado", nd=1, extra="o tecto de 60% coincide com a Lei de Sustentabilidade das Finanças Públicas; na série FMI o ponto de partida 2022 é 57,4%"),
    "MAC016": dict(pagina="PDN p.24 (Quadro Visão 2027)", conceito="dívida pública (% do PIB), perímetro MINFIN e PIB pré-rebasing", base=66, meta=60, modo="rel", fonte_serie="FMI WEO Abr-2026, governo geral, PIB rebaseado", nd=1),
}

META_CAT_LABEL = {"pdn": "PDN (valor directo)", "transposta": "PDN (transposta à base da série)", "convertida": "PDN (convertida de escala/unidade)",
                  "minplan": "MINPLAN (meta anual 2025, proxy)", "operacional": "Operacional (mantida/ajustada)", "sem_meta": "Sem meta comparável"}

# --------------------------------------------------------------------------------------------
# 2. NOVAS LINHAS DE CANDIDATOS (Feedback pontos 1 e 2)
# --------------------------------------------------------------------------------------------
NEW_ROWS = [
    dict(id="LAB028", dim="6. Mercado de Trabalho", subtema="6. Mercado de Trabalho", indicador="Taxa de emprego 15+ (INE/IEA, harmonizada OIT)",
         unidade="%", sentido="+", minimo=0, maximo=100, meta=None, peso=1, fonte="INE/IEA via OIT (ILOSTAT, LFS)", codigo="ILOSTAT EMP_DWAP_SEX_AGE_RT [BA:13951]",
         origem="OIT ILOSTAT — estimativa nacional (Inquérito sobre o Emprego em Angola, INE), harmonizada", nacional=1, grupo="LAB001", prioridade="Alta",
         url="https://ilostat.ilo.org/data/",
         values={2019: 62.873, 2020: 62.839, 2021: 63.474, 2022: 64.756, 2023: 62.903, 2024: 64.214, 2025: 66.617},
         nota="v8: série INE/IEA harmonizada pela OIT (definição internacional). Inelegível por cobertura (IEA existe desde 2019); usada na leitura complementar 12_Emprego_INE.",
         meta_origem="Sem meta no PDN", meta_cat="sem_meta"),
    dict(id="LAB029", dim="6. Mercado de Trabalho", subtema="6. Mercado de Trabalho", indicador="Desemprego jovem 15-24 (INE/IEA, harmonizada OIT)",
         unidade="%", sentido="-", minimo=0, maximo=50, meta=None, peso=1, fonte="INE/IEA via OIT (ILOSTAT, LFS)", codigo="ILOSTAT UNE_DEAP_SEX_AGE_RT Y15-24 [BA:13951]",
         origem="OIT ILOSTAT — estimativa nacional (IEA/INE), harmonizada", nacional=1, grupo="G_DESEMP_JUVENIL", prioridade="Alta",
         url="https://ilostat.ilo.org/data/",
         values={2019: 31.107, 2020: 29.081, 2021: 29.789, 2022: 26.585, 2023: 21.469, 2024: 26.203, 2025: 18.562},
         nota="v8: série INE/IEA harmonizada pela OIT. Inelegível por cobertura; usada em 12_Emprego_INE.", meta_origem="Sem meta no PDN", meta_cat="sem_meta"),
    dict(id="LAB030", dim="6. Mercado de Trabalho", subtema="6. Mercado de Trabalho", indicador="Taxa de desemprego 15+ (INE/IEA, 13.ª CIET, definição nacional alargada)",
         unidade="%", sentido="-", minimo=0, maximo=60, meta=25, peso=1, fonte="INE — Inquérito sobre o Emprego em Angola (IEA)", codigo="INE_IEA_TXDESEMP_13CIET",
         origem="INE — séries cronológicas IEA (antiga metodologia, 13.ª CIET): valor anual publicado pelo INE (2019-2022, 2024); 2023 = IV trim (único publicado); 2025 = Anuário IEA 2025 (Abril 2026)",
         nacional=1, grupo="G_DESEMPREGO", prioridade="Alta", url="https://www.ine.gov.ao/publicacoes/detalhes/NTM0NjU=",
         values={2019: 30.21, 2020: 32.18, 2021: 32.32, 2022: 30.16, 2023: 31.85, 2024: 31.47, 2025: 28.3},
         nota=("v8: definição alargada (inclui desencorajados) — nível ~2x superior à definição estrita OIT. 2019 = valor anual INE (inquérito iniciado no II trim); 2023 = apenas IV trim publicado; "
               "2025 = 28,3% (Anuário IEA 2025, Abril 2026). A série termina em 2025 com a mudança de metodologia. Inelegível por cobertura; usada em 12_Emprego_INE."),
         meta_origem="PDN p.13: taxa de desemprego 30% (2022) → 25% (2027), na definição INE (base coincide: 30,2% em 2022)", meta_cat="pdn"),
    dict(id="LAB031", dim="6. Mercado de Trabalho", subtema="6. Mercado de Trabalho", indicador="Taxa de desemprego 15+ (INE/IEA, 19.ª-21.ª CIET, nova metodologia)",
         unidade="%", sentido="-", minimo=0, maximo=60, meta=25, peso=1, fonte="INE — IEA (nova metodologia)", codigo="INE_IEA_TXDESEMP_19CIET",
         origem="INE — IEA IV trimestre 2025 (primeira publicação com as resoluções da 19.ª, 20.ª e 21.ª CIET)", nacional=1, grupo="G_DESEMPREGO", prioridade="Alta",
         url="https://www.ine.gov.ao/publicacoes/detalhes/NTA0Mzk=",
         values={2025: 20.12},
         nota="v8: quebra de série — nova metodologia desde o IV trim 2025 (20,1%; emprego 39,6%; informalidade 78,6%; desemprego jovem 43,3%). Série a acumular a partir de 2026.",
         meta_origem="PDN p.13: taxa de desemprego 30% → 25% (2027), fixada na definição antiga do INE; a nova metodologia não tem retropolação", meta_cat="pdn"),
    dict(id="INC031", dim="4. Inclusão Social e Protecção", subtema="4. Inclusão Social", indicador="Taxa de pobreza (< 3,00 USD/dia, PPC 2021)",
         unidade="%", sentido="-", minimo=0, maximo=80, meta=None, peso=1, fonte="Banco Mundial — PIP/WDI", codigo="SI.POV.DDAY (linha 3,00 USD PPC 2021)",
         origem="Banco Mundial — WDI (actualização Abr-2026): linha internacional de pobreza revista para 3,00 USD/dia (PPC 2021)", nacional=0, grupo="G_POBREZA", prioridade="Alta",
         url="https://data.worldbank.org/indicator/SI.POV.DDAY?locations=AO",
         values={2018: 39.3},
         nota="v8: novo padrão do Banco Mundial (Junho 2025). A meta do PDN (28%) refere-se à linha antiga de 2,15 USD (PPC 2017), mantida em INC001. Sem série anual: um inquérito (IDREA 2018-19) no período.",
         meta_origem="PDN p.92 fixa a meta na linha 2,15 USD (PPC 2017); sem meta comparável nesta linha", meta_cat="sem_meta"),
]

# --------------------------------------------------------------------------------------------
# 3. SÉRIES ACTUALIZADAS A PARTIR DE FONTES LOCAIS (ficheiros INE/BNA/FMI do repositório)
# --------------------------------------------------------------------------------------------
# PIB não petrolífero (MAC014/DIV017): a v7 já usava a folha Q5 (medidas de volume encadeadas, não ajustadas) do ficheiro trimestral do INE,
# cuja soma anual reproduz o PIB anual oficial; o ficheiro do repositório inclui o IV trim 2025 e não traz revisões → série VERIFICADA, sem alteração.
# (A folha Q2, com ajuste sazonal, dá valores até 0,22 p.p. diferentes e não é a referência anual.)
GDP_INE_Q = {2015: 0.740, 2016: -0.323, 2017: -0.206, 2018: -0.563, 2019: -1.048, 2020: -5.301, 2021: 1.200, 2022: 3.731, 2023: 1.431, 2024: 4.889, 2025: 3.029}

# FMI WEO (vintage Abril 2026, publicação 14-04-2026; cache local Calisto Ebo/_cache/imf_weo_ago.xml), precisão total da fonte:
# dívida bruta (GGXWDG_NGDP), saldo global (GGXCNL_NGDP) e saldo primário (GGXONLB_NGDP) do governo geral, % PIB.
IMF_DEBT = {2015: 50.385819, 2016: 65.691623, 2017: 59.607197, 2018: 81.635503, 2019: 100.79929, 2020: 119.838606, 2021: 75.502733, 2022: 57.427174, 2023: 75.739104, 2024: 57.106858, 2025: 51.310898}
IMF_BAL = {2015: -2.582085, 2016: -3.934178, 2017: -5.682721, 2018: 1.998796, 2019: -0.19541, 2020: -3.046261, 2021: 1.352043, 2022: 1.779136, 2023: -2.4918, 2024: -1.192327, 2025: -4.070237}
IMF_PRIM = {2015: -1.010223, 2016: -1.466382, 2017: -2.567522, 2018: 6.083505, 2019: 4.453149, 2020: 2.747432, 2021: 5.785409, 2022: 5.22853, 2023: 2.445, 2024: 3.401119, 2025: -0.38511}
# Denominador 2025 do WEO = PIB nominal das Contas Nacionais Trimestrais do INE (Quadro 12: 129 255 156 M Kz); as CN Anuais Preliminares 2025 (Maio 2026) reviram-no para 128 302 mil M Kz (−0,7%).
WEO_NGDP_2025 = 129255.157   # mil milhões Kz
INE_NGDP_2025_PRELIM = 128302.02

# INE IPCN: inflação média anual (idêntica ao WDI) e homóloga de Dezembro
INE_CPI_AVG = {2015: None, 2016: 30.69, 2017: 29.84, 2018: 19.63, 2019: 17.08, 2020: 22.27, 2021: 25.75, 2022: 21.36, 2023: 13.64, 2024: 28.24, 2025: 20.16}
INE_CPI_DEC = {2015: 12.09, 2016: 41.12, 2017: 23.67, 2018: 18.60, 2019: 16.90, 2020: 25.10, 2021: 27.03, 2022: 13.86, 2023: 20.01, 2024: 27.50, 2025: 15.70}

WGI_ORIGEM = "Banco Mundial — Worldwide Governance Indicators, edição 2025 (dados até 2024); herdado da v7, não reverificado na v8"

MAC004_NOTA = ("v8: substitui a série 'Compilação interna BDA' (valores arredondados, sem rastreabilidade) pela série oficial do FMI WEO Abr-2026 em precisão total (2025: 51,3%), "
               "que compila os dados do Ministério das Finanças (fonte histórica declarada pelo FMI: Ministry of Finance; GFSM 2014) — por isso mantém o marcador de fonte nacional. "
               "Mesma direcção que a leitura PDN/MINFIN (pico em 2020, forte descida até 2022), mas níveis não comparáveis: PDN/MINFIN em perímetro 'dívida pública' e PIB pré-rebasing "
               "(134% em 2020; 66% em 2022); FMI WEO em governo geral e PIB rebaseado (119,8% em 2020; 57,4% em 2022). Referências 2025: FMI 51,3%; Banco Mundial ≈52%; MINFIN dívida governamental 46,6%. "
               "O denominador de 2025 do WEO é o PIB nominal das Contas Nacionais Trimestrais do INE (129 255 mil M Kz); as CN Anuais Preliminares (Maio 2026) reviram-no para 128 302 mil M Kz, o que daria ≈51,7%.")

SERIES = {
    "MAC014": dict(values=None, tipo="metadados", origem="INE — Contas Nacionais Trimestrais, Quadro 5 (medidas de volume encadeadas, não ajustadas), ficheiro do repositório com o IV trim 2025; soma anual do PIB total menos Extracção e Refino de Petróleo",
                   nota_add="v8: série verificada contra a mesma publicação trimestral do INE (IV trim 2025 já incluído na v7; a soma anual da folha Q5 reproduz o PIB anual oficial em 2015-2025); sem alteração de valores."),
    "DIV017": dict(values=None, tipo="metadados", origem="INE — Contas Nacionais Trimestrais, Quadro 5 (medidas de volume encadeadas, não ajustadas), ficheiro do repositório com o IV trim 2025; soma anual do PIB total menos Extracção e Refino de Petróleo",
                   nota_add="v8: série verificada contra a publicação trimestral do INE (IV trim 2025 já incluído na v7); sem alteração de valores."),
    "MAC004": dict(values=IMF_DEBT, origem="FMI — World Economic Outlook, Abril 2026 (dívida bruta do governo geral, % PIB; compila dados do MINFIN)", fonte="MINFIN (via FMI WEO)",
                   codigo="WEO GGXWDG_NGDP", url=WEO_URL, nota_add=MAC004_NOTA),
    "MAC005": dict(values=IMF_BAL, origem="FMI — World Economic Outlook, Abril 2026 (saldo global do governo geral, % PIB; compila dados do MINFIN)", fonte="MINFIN (via FMI WEO)",
                   codigo="WEO GGXCNL_NGDP", url=WEO_URL,
                   nota_add="v8: substitui a série 'Compilação interna BDA' (ex.: 2017 = +1,5 quando o saldo oficial foi −5,7) pela série FMI WEO Abr-2026 em precisão total, que compila os dados do MINFIN (marcador de fonte nacional mantido)."),
    "MAC016": dict(values=None, tipo="metadados", origem="FMI — World Economic Outlook, Abril 2026", nota_add="v8: vintage Abr-2026 confirmado (série já presente na v7 em precisão total; sem alteração de valores)."),
    "MAC019": dict(values=None, tipo="metadados", origem="FMI — World Economic Outlook, Abril 2026", nota_add="v8: vintage Abr-2026 confirmado (série já presente na v7; sem alteração de valores)."),
    "MAC020": dict(values=None, tipo="metadados", origem="FMI — World Economic Outlook, Abril 2026", nota_add="v8: vintage Abr-2026 confirmado (série já presente na v7; sem alteração de valores)."),
    "MAC003": dict(values=None, tipo="metadados", origem="INE — IPC Nacional (média anual; idêntica ao WDI FP.CPI.TOTL.ZG)", fonte="INE",
                   nota_add="v8: validado com o IPCN do INE (2025: média anual 20,16%; homóloga de Dezembro 15,70%)."),
    "MAC027": dict(values=INE_CPI_DEC, origem="INE — IPC Nacional, variação homóloga de Dezembro (ficheiro 2026-05-09)", nota_add="v8: 2025 = 15,70% (Dez/2025)."),
}
for _g in WGI_PDN:
    SERIES[_g] = dict(values=None, tipo="metadados", origem=WGI_ORIGEM, nota_add="v8: origem documentada (WGI edição 2025, dados até 2024).")

# --------------------------------------------------------------------------------------------
# 4. RENOMEAÇÕES / CORRECÇÕES DE METADADOS
# --------------------------------------------------------------------------------------------
RENAME = {
    "INC001": dict(indicador="Taxa de pobreza (< 2,15 USD/dia, PPC 2017)", codigo="PIP (PPC 2017) — headcount 2,15 USD/dia [ex-SI.POV.DDAY, WDI ≤2024]",
                   origem="Banco Mundial — PIP/WDI (PPC 2017), via UNdata; valor 2018 (IDREA 2018-19) = 31,1%, já constante das versões anteriores do IGDA; PDN p.92 confirma 31% (2022*)",
                   url=PIP_URL,
                   nota_sub=[("Poverty headcount ratio at $3.00 a day (2021 PPP) (% of population)", "Poverty headcount ratio at $2.15 a day (2017 PPP) (% of population) — PIP (versão PPC 2017)/WDI até 2024")],
                   nota_add="v8: a linha de 2,15 USD (PPC 2017) é a referência do PDN (31% em 2022; 28% em 2027). Valor 2018 (IDREA 2018-19) = 31,1% (Banco Mundial/PIP). A nova linha de 3,00 USD (PPC 2021) consta em INC031. Sem novo inquérito de despesas desde o IDREA 2018-19.",
                   values={2018: 31.1}),
    "INC028": dict(indicador="Pobreza a 4,20 USD/dia (PPC 2021)", codigo="SI.POV.LMIC (linha 4,20 USD PPC 2021)", nota_add="v8: nome alinhado com a linha efectivamente extraída do WDI (Abr-2026): 53,7% em 2018 corresponde à linha de 4,20 USD (PPC 2021), sucessora da linha de 3,65 USD (PPC 2017)."),
    "INC029": dict(indicador="Pobreza a 8,30 USD/dia (PPC 2021)", codigo="SI.POV.UMIC (linha 8,30 USD PPC 2021)", nota_add="v8: nome alinhado com a linha efectivamente extraída do WDI (Abr-2026): 80% em 2018 corresponde à linha de 8,30 USD (PPC 2021), sucessora da linha de 6,85 USD (PPC 2017)."),
    "INC030": dict(indicador="Fosso de pobreza a 3,00 USD/dia (PPC 2021)", codigo="SI.POV.GAPS (linha 3,00 USD PPC 2021)", nota_add="v8: nome alinhado com a linha efectivamente extraída do WDI (Abr-2026): 16,3% em 2018 é o fosso à linha de 3,00 USD (PPC 2021)."),
}

# --------------------------------------------------------------------------------------------
# 5. ACTUALIZAÇÕES RESULTANTES DA RECOLHA WEB (INE, MINFIN, FMI) — ronda 1
# --------------------------------------------------------------------------------------------
SERIES.update({
    "MAC001": dict(values={2025: 3.13}, origem="INE — Contas Nacionais Anuais Preliminares 2025 (05-05-2026): PIB real +3,13%; 2015-2024 CN anuais definitivas",
                   nota_add="v8: 2025 confirmado pelas Contas Nacionais Anuais Preliminares do INE (+3,13%; sector petrolífero −5,23%). Coincide com o FMI (Art. IV 2026: 3,1%) e com o WDI."),
    "LAB004": dict(values={2025: 78.8}, origem="INE — IEA, Anuário 2025 (Abril 2026), metodologia 13.ª CIET; 2019-2024 séries cronológicas IEA (valor anual publicado)",
                   nota_add="v8: 2025 = 78,8% (Anuário IEA 2025, valor anual; a v7 usava o III trim). Quebra de série a partir do IV trim 2025 (nova metodologia: 78,6%). 2019 = valor anual publicado pelo INE (74,5%), que diverge da média dos trimestres II-IV (79,5%)."),
    "LAB005": dict(values={2019: 25.31, 2025: 21.2}, fonte="INE", origem="INE — IEA, taxa de emprego formal publicada (séries cronológicas 13.ª CIET); 2025 = Anuário IEA 2025",
                   nota_add="v8: 2025 = 21,2% (Anuário IEA 2025, complemento do emprego informal 78,8%); 2019 corrigido de 25,46 (= 100 − informal) para 25,31 (valor anual publicado pelo INE; no ano 2019 do INE formal + informal = 99,86%)."),
    "SAU003": dict(values={2024: None, 2025: None}, origem="Banco Mundial — WDI (estimativas modeladas MMEIG) até 2023",
                   nota_add=("v8: correcção — a v7 registava 185 (2024) e 170 (2025) sem suporte na série MMEIG; 2024-2025 passam a carry-forward (183) até nova ronda MMEIG. "
                             "A estimativa directa do IIMS 2023-24 (Quadro 17.4: 170 por 100 mil, IC 99-242, ≈7 anos anteriores ao inquérito) é compatível com a série mas não é comparável ponto a ponto "
                             "com as estimativas modeladas; figura em SAU029 como validação cruzada.")),
    "SAU029": dict(values=None, tipo="metadados", nota_add="v8: estimativa directa do IIMS 2023-24 (170, IC 99-242) mantida nesta linha como validação cruzada da série modelada MMEIG (SAU003); não incorporada na série do índice."),
    "HUM010": dict(values=None, tipo="metadados", origem="Banco Mundial/UNESCO — WDI (despesa executada, UIS) até 2023",
                   nota_add="v8: sem observação UIS para 2024-2025. As dotações do OGE confirmam a trajectória descendente sinalizada no feedback: 2,0% do PIB (OGE 2024), 1,8% (OGE 2025), 1,7% (OGE 2026), 6,4-6,9% da despesa total (UNICEF, Budget Briefs)."),
    "SAU004": dict(values=None, tipo="metadados", origem="Banco Mundial/OMS — GHED (despesa corrente em saúde) até 2023",
                   nota_add="v8: sem observação GHED para 2024-2025. Dotações OGE para a saúde: 5,5% do OGE (2024), 5,7-6,3% (2025), 6,32% do OGE e 1,5% do PIB (2026) — UNICEF Budget Briefs."),
})

# Kwenda — beneficiários acumulados (agregados familiares), MINPLAN painel PDN; cobertura = % dos agregados familiares (Censo 2024: 9 110 616 agregados)
KWENDA_AGREGADOS = {2022: 610832, 2023: 951203, 2024: 1070037, 2025: 1350850}
HOUSEHOLDS_2024 = 9110616
NEW_ROWS.append(dict(id="INC032", dim="4. Inclusão Social e Protecção", subtema="4. Inclusão Social", indicador="Cobertura do Kwenda (agregados beneficiários acumulados, % dos agregados familiares)",
    unidade="%", sentido="+", minimo=0, maximo=50, meta=19.8, peso=1, fonte="MINPLAN/MASFAMU — Balanço do PDN (painel anual 2025); INE Censo 2024 (agregados)", codigo="MINPLAN_PDN_KWENDA_COV",
    origem="MINPLAN — painel 'Implementation - Key Indicators by Domain' (beneficiários de transferências sociais monetárias, cumulativo) ÷ agregados familiares do Censo 2024 (9 110 616)", nacional=1, grupo="G_PROTECCAO_SOCIAL", prioridade="Alta",
    url="https://minplan.gov.ao/",
    values={y: round(100 * v / HOUSEHOLDS_2024, 2) for y, v in KWENDA_AGREGADOS.items()},
    nota="v8: novo indicador de protecção social não contributiva (feedback ponto 2). Denominador fixo (Censo 2024) para isolar a expansão do programa. Meta 2027 = 1,8 milhões de agregados (meta anual MINPLAN 2025, usada como proxy) ÷ 9,11 M. Inelegível por cobertura (programa iniciado em 2020; primeiro dado anual em 2021); documentado para incorporação futura.",
    meta_origem="MINPLAN, Balanço PDN 2025: meta anual de 1 800 000 agregados beneficiários (19,8% dos agregados do Censo 2024), usada como proxy da meta 2027; o PDN p.92 não fixa meta em % de cobertura", meta_cat="minplan"))

# --------------------------------------------------------------------------------------------
# 6. ACTUALIZAÇÕES RESULTANTES DA RECOLHA WEB (MINPLAN / MAPTSS / INSS / FAS-Kwenda) — ronda 1
# --------------------------------------------------------------------------------------------
# INSS — segurados inscritos (stock, milhões). Fontes: INSS/MAPTSS (Boletim Anual da Protecção Social Obrigatória 2025; Forbes África Lusófona 27-01-2025; Expansão 2026)
INSS_SEGURADOS_M = {2020: 1.97, 2022: 2.13, 2024: 3.004238, 2025: 3.341475}
NEW_ROWS.append(dict(id="INC033", dim="4. Inclusão Social e Protecção", subtema="4. Inclusão Social", indicador="Segurados inscritos na protecção social obrigatória (INSS, milhões)",
    unidade="milhões", sentido="+", minimo=0, maximo=8, meta=4.3, peso=1, fonte="INSS/MAPTSS — Boletim Anual da Protecção Social Obrigatória", codigo="MAPTSS_INSS_SEGURADOS",
    origem="INSS/MAPTSS: 2024 = 3 004 238 (Balanço 2024); 2025 = 3 341 475 (Boletim Anual 2025); 2020 = 1,97 M e Fev-2022 = 2,13 M (INSS via Expansão)", nacional=1, grupo="G_PROTECCAO_SOCIAL_CONTRIB", prioridade="Alta",
    url="https://www.maptss.gov.ao/2025/01/27/inss-tem-inscrito-mais-de-tres-milhoes-de-segurados/",
    values=INSS_SEGURADOS_M,
    nota="v8: indicador de protecção social contributiva (PDN, Programa 22: segurados 2,5 → 4,3 milhões). Stock acumulado de inscritos (inclui inactivos); 2022 refere-se a Fevereiro. Máximo 8 = tecto operacional para o horizonte 2027-2035 (a meta 2050 do PDN, 13,6 M, exigirá revisão da fronteira). Inelegível por cobertura (4 observações); série anual completa a obter junto do INSS.",
    meta_origem="PDN p.92: número de segurados registados na protecção social obrigatória 2,5 (2022) → 4,3 milhões (2027) → 13,6 (2050); indicador do próprio PDN", meta_cat="pdn"))
# Kwenda — 2021: pouco mais de 300 mil agregados pagos (FAS/IDL via Expansão)
for row in NEW_ROWS:
    if row["id"] == "INC032":
        row["values"][2021] = round(100 * 300000 / HOUSEHOLDS_2024, 2)
        row["nota"] += " 2021 ≈ 300 mil agregados pagos (FAS/IDL); 2022-2025 = painel MINPLAN (610 832; 951 203; 1 070 037; 1 350 850)."

# --------------------------------------------------------------------------------------------
# 7. GRAFIA — uniformização para a norma pré-Acordo Ortográfico usada nos documentos (nomes de dimensões, subtemas, indicadores e unidades)
# --------------------------------------------------------------------------------------------
DIM8_OLD = "8. Diversificação Produtiva e Setor Privado"
DIM8_NEW = "8. Diversificação Produtiva e Sector Privado"
ORTOGRAFIA = [("Setor", "Sector"), ("setor", "sector"), ("Proteção", "Protecção"), ("proteção", "protecção"), ("direta", "directa"), ("Eletricidade", "Electricidade"),
              ("eletricidade", "electricidade"), ("hidroelétric", "hidroeléctric"), ("Elétric", "Eléctric"), ("elétric", "eléctric"), ("afetada", "afectada"), ("eletrónicos", "electrónicos")]

# Nomes curtos dos indicadores do índice (documentos e apresentação)
NOMES_CURTOS = {"GOV01": "Percepção da Corrupção", "GOV05": "Qualidade Regulatória", "GOV03": "Eficiência do Governo", "GOV02": "Estado de Direito", "GOV004": "Estabilidade Política",
                "MAC014": "PIB não petrolífero", "MAC001": "Crescimento do PIB", "MAC005": "Saldo Orçamental", "MAC003": "Inflação", "MAC007": "Reservas Internacionais", "MAC004": "Dívida Pública", "MAC002": "PIB per capita",
                "HUM030": "Escolaridade obrigatória", "HUM002": "Esperança de vida", "HUM006": "Mortalidade infantil", "HUM010": "Despesa pública em educação",
                "INC06": "Mulheres no parlamento", "INC04": "Paridade de género (WEF)", "INC005": "RNB per capita (USD)",
                "INF002": "Acesso a água potável", "INF001": "Electrificação", "INF006": "Densidade rodoviária", "INF05": "Energia limpa para cozinhar", "INF024": "Tráfego portuário (TEU)",
                "LAB003": "Desemprego juvenil", "LAB002": "Taxa de desemprego", "LAB001": "Taxa de emprego", "LAB016": "Participação feminina", "LAB025": "Produtividade do trabalho",
                "SAU003": "Mortalidade materna", "SAU04": "Mortalidade < 5 anos", "SAU004": "Despesa em saúde", "SAU006": "Cobertura vacinal", "SAU017": "Despesa directa das famílias", "SAU002": "Desnutrição", "SAU020": "Incidência de malária",
                "DIV017": "PIB não petrolífero", "DIV007": "Indústria transformadora", "DIV010": "Combustíveis nas exportações", "DIV001": "Capacidades produtivas", "DIV015": "Crédito ao sector privado",
                "GOV027": "Liberdade de imprensa (RSF)", "MAC029": "Dívida externa pública", "MAC039": "Dívida externa total", "MAC010": "IDE líquido (% PIB)", "MAC033": "IDE líquido (BNA)",
                "HUM018": "Alfabetização feminina", "INF016": "Electricidade renovável", "HUM004": "Matrícula no primário", "HUM003": "Alfabetização", "HUM024": "HALE", "HUM025": "Mortalidade < 5 anos (IIMS)",
                "INC001": "Pobreza 2,15 USD", "INC006": "Pobreza (linha nacional)", "INF003": "Saneamento básico", "LAB018": "Desemprego (INE harmonizada)", "LAB005": "Formalização", "LAB004": "Emprego informal",
                "SAU029": "Mortalidade materna (IIMS)", "MAC016": "Dívida bruta (FMI)", "MAC028": "Crescimento do PIB (INE)", "MAC006": "IDE não petrolífero", "MAC013": "Crédito ao sector privado (% PIB)",
                "MAC034": "Reservas (BNA)", "INC032": "Cobertura do Kwenda", "INC033": "Segurados INSS", "LAB030": "Desemprego INE (13.ª CIET)", "LAB031": "Desemprego INE (nova metodologia)"}

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

if __name__ == "__main__":
    for k, (nome, p22, p27, e22) in WGI_PDN.items():
        print(k, nome, METAS[k][0], "sens. p.p.:", wgi_meta_pp(e22, p22, p27))
