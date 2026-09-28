"""
Actualiza a apresentação CEX (v9 -> v10) com os resultados da v8: textos, tabelas, gráficos e um novo diapositivo de cruzamento com o PDN.
"""
from __future__ import annotations
import os
import json, sys, copy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx_tools import set_text_preserve, set_table_cell, find_shape, replace_chart_series
from docx_tools import fmt

SC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "")
A = json.load(open(SC + "analysis.json"))["v8"]
YEARS = [str(y) for y in range(2015, 2026)]
DIMS8 = ["Governança", "Macroeconomia", "Capital Humano", "Inclusão Social", "Infraestruturas", "Mercado Trabalho", "Saúde/Alimentar", "Diversificação"]
FULL = {"Governança": "1. Governança e Estado de Direito", "Macroeconomia": "2. Estabilidade Macroeconómica", "Capital Humano": "3. Capital Humano", "Inclusão Social": "4. Inclusão Social e Protecção",
        "Infraestruturas": "5. Infraestruturas e Serviços", "Mercado Trabalho": "6. Mercado de Trabalho", "Saúde/Alimentar": "7. Segurança Alimentar e Saúde", "Diversificação": "8. Diversificação Produtiva e Setor Privado"}
NAMES = {"GOV01": "Percepção da Corrupção", "GOV05": "Qualidade Regulatória", "GOV03": "Eficiência do Governo", "GOV02": "Estado de Direito", "GOV004": "Estabilidade Política",
         "MAC014": "PIB não petrolífero", "MAC001": "Crescimento do PIB", "MAC005": "Saldo Orçamental", "MAC003": "Inflação", "MAC007": "Reservas Internacionais", "MAC004": "Dívida Pública", "MAC002": "PIB per capita",
         "HUM030": "Escolaridade obrigatória", "HUM002": "Esperança de vida", "HUM006": "Mortalidade infantil", "HUM010": "Despesa pública em educação",
         "INC06": "Mulheres no parlamento", "INC04": "Paridade de género (WEF)", "INC005": "RNB per capita (USD)",
         "INF002": "Acesso a água potável", "INF001": "Electrificação", "INF006": "Densidade rodoviária", "INF05": "Energia limpa para cozinhar", "INF024": "Tráfego portuário (TEU)",
         "LAB003": "Desemprego juvenil", "LAB002": "Taxa de desemprego", "LAB001": "Taxa de emprego", "LAB016": "Participação feminina", "LAB025": "Produtividade do trabalho",
         "SAU003": "Mortalidade materna", "SAU04": "Mortalidade < 5 anos", "SAU004": "Despesa em saúde", "SAU006": "Cobertura vacinal", "SAU017": "Despesa directa das famílias", "SAU002": "Desnutrição", "SAU020": "Incidência de malária",
         "DIV017": "PIB não petrolífero", "DIV007": "Indústria transformadora", "DIV010": "Combustíveis nas exportações", "DIV001": "Capacidades produtivas", "DIV015": "Crédito ao sector privado"}

sub = A["subindices"]; igda = A["igda"]; esc = A["escala"]; ind = A["indicators"]; cov = A["coverage"]


def f1(x, sign=False):
    return fmt(x, 1, sign=sign)


def drivers(dim):
    items = [i for i in ind if i["dim"] == dim and i["d_15_25"] is not None]
    pos = sorted([i for i in items if i["d_15_25"] >= 0], key=lambda x: -x["d_15_25"])
    neg = sorted([i for i in items if i["d_15_25"] < 0], key=lambda x: x["d_15_25"])
    return pos, neg


def line(i):
    return f"{NAMES.get(i['id'], i['nome'])}  {f1(i['d_15_25'], sign=True)}"


def main(src, out):
    prs = Presentation(src)
    S = prs.slides
    d15 = {d: sub[d][-1] - sub[d][0] for d in DIMS8}
    d23 = {d: sub[d][-1] - sub[d][8] for d in DIMS8}
    n_up = sum(1 for d in DIMS8 if d15[d] > 0)
    imin = min(range(11), key=lambda k: igda[k])

    # ---- slide 1 ----
    s = S[0]
    set_text_preserve(find_shape(s, shape_id=263), ["Proposta actualizada (v8)", "Setembro - 2026"])
    set_text_preserve(find_shape(s, shape_id=3), [f"IGDA-BDA  ·  2015–2025  ·  8 dimensões  ·  {cov['n_sel']} indicadores  ·  metas alinhadas com o PDN 2023-2027"])

    # ---- slide 7 ----
    s = S[6]
    kpis = {"Agrupar 7": (f1(igda[-1]), "IGDA 2025"), "Agrupar 15": (f1(igda[imin]), f"Mínimo da série ({2015+imin})"), "Agrupar 20": (f1(igda[-1]-igda[0], sign=True).replace("−", "−") + " ", "Variação 2015–2025 (p.p.)"),
            "Agrupar 26": (f"{n_up} / 8", "Dimensões em progresso"), "Agrupar 33": (f1(igda[0]), "IGDA 2015"), "Agrupar 37": ("Médio", "Classificação IGDA 2025 (40–60) ")}
    for sh in s.shapes:
        if sh.shape_type == 6 and sh.name in kpis:
            val, lab = kpis[sh.name]
            for sub_sh in sh.shapes:
                if sub_sh.name == "TextBox 9034": set_text_preserve(sub_sh, [val])
                if sub_sh.name == "TextBox 9035": set_text_preserve(sub_sh, [lab])
    order = sorted(DIMS8, key=lambda d: -esc[d][-1])
    dest = find_shape(s, shape_id=902)
    set_text_preserve(dest, [
        f"IGDA 2025: {f1(igda[-1])} — máximo da série (v8); classificação Médio (40–60)",
        f"Variação 2015–2025: {f1(igda[-1]-igda[0], sign=True)} p.p.",
        f"Mínimo da série: {f1(igda[imin])} ({2015+imin}, choque pandémico)",
        f"Trajectória: {n_up} de 8 dimensões melhoraram face a 2015",
        f"Evolução — lideram: {order[0]} {fmt(esc[order[0]][-1]-100,1,pct=True,sign=True)} e {order[1]} {fmt(esc[order[1]][-1]-100,1,pct=True,sign=True)} (2015 = 100)",
        f"Evolução — recuam: {order[-1]} {fmt(esc[order[-1]][-1]-100,1,pct=True,sign=True)} e {order[-2]} {fmt(esc[order[-2]][-1]-100,1,pct=True,sign=True)}",
    ])
    tbl = find_shape(s, shape_id=910).table
    for ri, d in enumerate(DIMS8, start=1):
        for cj, v in enumerate(sub[d], start=1):
            set_table_cell(tbl, ri, cj, f1(v))
    for cj, v in enumerate(igda, start=1):
        set_table_cell(tbl, 9, cj, f1(v))

    # ---- slide 8 ----
    s = S[7]
    for sh in s.shapes:
        if sh.shape_type == 6:
            for sub_sh in sh.shapes:
                if getattr(sub_sh, "has_chart", False) and sub_sh.has_chart:
                    replace_chart_series(sub_sh.chart, YEARS, {"EVOLUÇÃO DO IGDA-BDA (Índice Global)": igda})
                if sub_sh.name == "Retângulo: Cantos Arredondados 50":
                    set_text_preserve(sub_sh, [
                        f"2015–2019 — estagnação em torno de {f1(sum(igda[:5])/5)} pontos: ganhos de governança anulados pela deterioração macroeconómica.",
                        f"2020 — choque COVID-19 / petróleo: {f1(igda[5]-igda[4], sign=True)} p.p., para {f1(igda[5])} — único ano na banda Baixo.",
                        f"2021–2024 — recuperação sustentada: {f1(igda[9]-igda[5], sign=True)} p.p. face a 2020, com máximos da série.",
                        f"2025 — {f1(igda[-1])} pontos: estável ({f1(igda[-1]-igda[-2], sign=True)} p.p.); valor provisório (v8: dados INE/FMI de 2026).",
                    ])
    up = sorted([d for d in DIMS8 if d15[d] > 0], key=lambda d: -d15[d]); down = sorted([d for d in DIMS8 if d15[d] <= 0], key=lambda d: d15[d])
    low = sorted(DIMS8, key=lambda d: sub[d][-1])[:2]
    set_text_preserve(find_shape(s, shape_id=47), [
        f"No conjunto do período, o IGDA sobe de {f1(igda[0])} (2015) para {f1(igda[-1])} (2025) — {f1(igda[-1]-igda[0], sign=True)} p.p. —, permanecendo na banda Médio (40–60). O progresso é puxado por " + ", ".join(f"{d} ({f1(d15[d], sign=True)})" for d in up[:3]) + "; recuam " + " e ".join(f"{d} ({f1(d15[d], sign=True)})" for d in down) + ".",
        f"O nível continua condicionado por duas dimensões na banda Baixo — {low[0]} ({f1(sub[low[0]][-1])}) e {low[1]} ({f1(sub[low[1]][-1])}) — e sensível a choques macroeconómicos, como em 2016–2017 e, sobretudo, em 2020.",
        "Leitura: melhoria lenta mas consistente desde 2020. A Diversificação trava o nível do índice por ser a dimensão de score mais baixo — mas é, ao mesmo tempo, a que mais progride desde 2015. A v8 confirma a leitura da v7 com dados oficiais de 2026 e metas do PDN.",
    ])

    # ---- slide 9 ----
    s = S[8]
    set_text_preserve(find_shape(s, shape_id=23), [
        f"Entre 2023 e 2025 o IGDA sobe de {f1(igda[8])} para {f1(igda[-1])} pontos ({f1(igda[-1]-igda[8], sign=True)} p.p.), o máximo da série. O ganho concentra-se em 2024 ({f1(igda[9]-igda[8], sign=True)} p.p.); em 2025 o índice fica estável ({f1(igda[-1]-igda[-2], sign=True)} p.p.).",
        f"2025 é provisório: {cov['last2025']} dos {cov['n_sel']} indicadores têm observação nesse ano ({cov['last2024']} até 2024 e {cov['last2023']} até 2023); os restantes entram por carry-forward.",
    ])
    def top_driver(dim, positive):
        items = [i for i in ind if i["dim"] == dim and i["d_23_25"] is not None]
        items = sorted(items, key=lambda x: (-x["d_23_25"] if positive else x["d_23_25"]))
        i = items[0]; return f"{NAMES.get(i['id'], i['nome']).lower()} ({f1(i['d_23_25'], sign=True)})"
    rec = sorted([d for d in DIMS8 if d23[d] < 0], key=lambda d: d23[d]); mel = sorted([d for d in DIMS8 if d23[d] >= 0], key=lambda d: -d23[d])
    set_text_preserve(find_shape(s, shape_id=261), ["Dimensões em recuo"] + [f"•  {FULL[d][3:]} {f1(d23[d], sign=True)} — {top_driver(d, False)}" for d in rec] +
                      ["Leitura: a recuperação pós-2020 mantém-se, mas o ritmo abranda em 2025. A leitura INE do emprego (folha 12_Emprego_INE) mostra melhoria em 2025 que as séries OIT ainda não captam."])
    set_text_preserve(find_shape(s, shape_id=10), ["Dimensões em melhoria"] + [f"•  {FULL[d][3:]} {f1(d23[d], sign=True)} — {top_driver(d, True)}" for d in mel])

    # ---- slide 10 ----
    s = S[9]
    set_text_preserve(find_shape(s, shape_id=23), [
        "Cada dimensão é reexpressa face ao seu próprio valor de 2015. Como todas partem de 100, as trajectórias tornam-se directamente comparáveis — o que os níveis, com fronteiras de normalização distintas, não permitem.",
        "",
        f"A leitura inverte-se: a Diversificação, a dimensão de nível mais baixo ({f1(sub['Diversificação'][-1])}), é a que mais progrediu desde 2015 ({fmt(esc['Diversificação'][-1]-100,1,pct=True,sign=True)}). O Mercado de Trabalho, o nível mais alto, é dos que menos avança ({fmt(esc['Mercado Trabalho'][-1]-100,1,pct=True,sign=True)}). Recuam a Inclusão Social e a Saúde/Segurança Alimentar.",
    ])
    tbl = find_shape(s, shape_id=920).table
    for ri, d in enumerate(order, start=1):
        set_table_cell(tbl, ri, 0, d)
        for cj, v in enumerate(esc[d], start=1):
            set_table_cell(tbl, ri, cj, fmt(v, 0))
        set_table_cell(tbl, ri, 12, fmt(esc[d][-1]-100, 1, pct=True, sign=True))
    for cj, v in enumerate(esc["IGDA"], start=1):
        set_table_cell(tbl, 9, cj, fmt(v, 0))
    set_table_cell(tbl, 9, 12, fmt(esc["IGDA"][-1]-100, 1, pct=True, sign=True))
    set_text_preserve(find_shape(s, shape_id=922), [f"Esta escala compara ritmos de progresso, não patamares de desenvolvimento: {fmt(esc['Diversificação'][-1],0)} na Diversificação e {fmt(esc['Mercado Trabalho'][-1],0)} no Mercado de Trabalho significam que a primeira cresceu mais desde 2015, não que esteja num nível superior. Fonte: construtor IGDA-BDA v8, folha 11_Escala_Comum."])

    # ---- slides 11-14: drivers + charts ----
    def set_drivers(slide, dim, pos_id, neg_id, extra_neg=None, extra_pos=None):
        pos, neg = drivers(dim)
        pl = [line(i) for i in pos] + (extra_pos or [])
        nl = [line(i) for i in neg] + (extra_neg or [])
        set_text_preserve(find_shape(slide, shape_id=pos_id), pl)
        set_text_preserve(find_shape(slide, shape_id=neg_id), nl)

    def set_chart(slide, dim, chart_name=None, shape_id=None, in_group=None):
        series = {FULL[dim]: sub[dim]}
        if in_group:
            for sh in slide.shapes:
                if sh.shape_type == 6 and sh.name == in_group:
                    for sub_sh in sh.shapes:
                        if getattr(sub_sh, "has_chart", False) and sub_sh.has_chart:
                            replace_chart_series(sub_sh.chart, YEARS, series)
            return
        sh = find_shape(slide, shape_id=shape_id)
        replace_chart_series(sh.chart, YEARS, series)

    gov_pv = next(i for i in ind if i["id"] == "GOV004")
    s = S[10]
    set_chart(s, "Governança", in_group="Agrupar 266"); set_chart(s, "Macroeconomia", shape_id=63)
    set_drivers(s, "Governança", 60, 286, extra_neg=[f"(recuo de {f1(gov_pv['d_23_25'], sign=True)} desde 2023)"])
    pos, neg = drivers("Macroeconomia")
    set_text_preserve(find_shape(s, shape_id=287), [line(i) for i in pos])
    set_text_preserve(find_shape(s, shape_id=288), [line(i) for i in neg] + ["(Δ do score normalizado, p.p., 2015–2025; saldo e dívida com séries FMI WEO Abr-2026)"])
    s = S[11]
    set_chart(s, "Capital Humano", shape_id=2); set_chart(s, "Inclusão Social", shape_id=6)
    set_drivers(s, "Capital Humano", 60, 286, extra_neg=["(único indicador da dimensão em queda; dotação OGE 2026 = 1,7% do PIB)"])
    pos, neg = drivers("Inclusão Social")
    set_text_preserve(find_shape(s, shape_id=287), [line(i) for i in pos])
    set_text_preserve(find_shape(s, shape_id=288), [line(i) for i in neg] + ["(quebra cambial de 2015–2020 ainda não recuperada)", "(Δ do score normalizado, p.p., 2015–2025)"])
    s = S[12]
    set_chart(s, "Infraestruturas", shape_id=2); set_chart(s, "Mercado Trabalho", shape_id=6)
    teu = next(i for i in ind if i["id"] == "INF024")
    set_drivers(s, "Infraestruturas", 60, 286, extra_neg=[f"(recupera {f1(teu['d_23_25'], sign=True)} desde 2023)"])
    pos, neg = drivers("Mercado Trabalho")
    set_text_preserve(find_shape(s, shape_id=287), [line(i) for i in pos])
    set_text_preserve(find_shape(s, shape_id=288), [line(i) for i in neg] + ["(Δ do score normalizado, p.p., 2015–2025; leitura INE em 12_Emprego_INE)"])
    s = S[13]
    set_chart(s, "Saúde/Alimentar", shape_id=5); set_chart(s, "Diversificação", shape_id=14)
    pos, neg = drivers("Saúde/Alimentar")
    vac = next(i for i in ind if i["id"] == "SAU006")
    set_text_preserve(find_shape(s, shape_id=60), [line(i) for i in pos])
    set_text_preserve(find_shape(s, shape_id=286), [line(i) for i in neg] + [f"Cobertura vacinal (2023–25)  {f1(vac['d_23_25'], sign=True)}"])
    pos, neg = drivers("Diversificação")
    fuel = next(i for i in ind if i["id"] == "DIV010")
    set_text_preserve(find_shape(s, shape_id=901), [line(i) for i in pos])
    set_text_preserve(find_shape(s, shape_id=902), [line(i) for i in neg] + [f"Exportações concentradas em combustíveis (score {f1(fuel['score_2025'])})", "(Δ do score normalizado, p.p., 2015–2025)"])

    # ---- slide 15 ----
    s = S[14]
    set_text_preserve(find_shape(s, shape_id=261), ["Próximos passos (v8 → v9)",
        "•  Concluído (v8): metas 2027 alinhadas com o PDN; séries de dívida/saldo com fonte oficial (FMI WEO); leitura INE do emprego (folha 12)",
        "•  Actualizar emprego quando a OIT incorporar o IEA 2025 (Nov-2026); a partir de 2028 avaliar séries INE no índice",
        "•  Inclusão: obter série anual INSS 2015-2025 e próximo inquérito de despesas (pobreza)",
        "•  Substituir a escolaridade obrigatória (variável legal) por um indicador de resultado",
        "•  Rever 2024–2025 com dados definitivos; executar a análise de sensibilidade (passo 7 OCDE/JRC)"])
    set_text_preserve(find_shape(s, shape_id=10), ["Limitações actuais",
        f"•  2025 provisório — {cov['n_sel']-cov['last2025']} dos {cov['n_sel']} indicadores entram por extrapolação",
        "•  Quebra de série no IEA (nova metodologia desde o IV trim 2025) e rebasing do PIB (INE 2025)",
        "•  Inclusão Social com 3 indicadores: pobreza, Kwenda e INSS documentados mas sem série elegível",
        "•  Níveis não comparáveis entre dimensões: cada uma tem goalposts próprios",
        "•  Sensibilidade às escolhas de desenho ainda não quantificada"])

    # ---- novo diapositivo: cruzamento com o PDN (inserido após o 10) ----
    layout = prs.slide_layouts[4]  # Blank
    new = prs.slides.add_slide(layout)
    # footer image + title style copied from slide 10
    ref = S[9]
    from lxml import etree
    def has_rel(el):
        return any(k.endswith("}id") or k.endswith("}embed") or k.endswith("}link") for e in el.iter() for k in e.attrib.keys())
    for sh in ref.shapes:
        if sh.shape_type == 13 and sh.name == "Imagem 257":
            new.shapes.add_picture(__import__("io").BytesIO(sh.image.blob), sh.left, sh.top, sh.width, sh.height)
        elif sh.shape_id in (259, 6, 260) and not has_rel(sh._element):
            new.shapes._spTree.append(copy.deepcopy(sh._element))
    tshape = find_shape(new, contains="ESCALA COMUM")
    if tshape is None:
        tshape = new.shapes.add_textbox(Emu(559886), Emu(98690), Emu(11684286), Emu(474489))
        tshape.text_frame.text = "x"; r = tshape.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(24); r.font.bold = True; r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    set_text_preserve(tshape, ["3 – CRUZAMENTO COM O PDN 2023-2027 _2023 – 2025   "])
    tb = new.shapes.add_textbox(Emu(559886), Emu(700000), Emu(11100000), Emu(600000)); tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = ("Em seis das oito dimensões o IGDA e o Balanço do PDN contam a mesma história — uma validação externa da construção. Nos dois pilares do PDN as leituras divergem: "
                                   "o capital humano avança devagar e a segurança alimentar regride nos indicadores de processo antes de as metas de mortalidade o reflectirem.")
    p.runs[0].font.size = Pt(12); p.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    rows = [("Governança", "IPC 33 → 34; percentis WGI +4 a +7 p.p.", f"{f1(d23['Governança'], sign=True)} p.p.; estabilidade política −5,2", "Consistente"),
            ("Macroeconomia", "Dívida 60% do PIB superada (FMI 51,3%); inflação 15,7% Dez-25; défice 4,1%", f"{f1(d23['Macroeconomia'], sign=True)} p.p.; dívida +16,3, inflação −13,0, saldo −6,3", "Consistente"),
            ("Capital Humano", "Esperança de vida 62 → 63; educação 11,8% da despesa", f"{f1(d23['Capital Humano'], sign=True)} p.p.; despesa em educação em queda (OGE 1,7% PIB 2026)", "Consistente, alerta"),
            ("Inclusão Social", "Pobreza 31 → 28%; Kwenda 1,35 M agregados (meta 1,8 M)", f"{f1(d23['Inclusão Social'], sign=True)} p.p., só por representação política", "Parcial — não mede pobreza"),
            ("Infraestruturas", "Electrificação 43 → 49% (48% em 2025); PIP paralisado", f"{f1(d23['Infraestruturas'], sign=True)} p.p., lento", "Consistente"),
            ("Mercado de Trabalho", "Desemprego 30 → 25%: 28,3% (2025, IEA); 20,1% nova metodologia", f"{f1(d23['Mercado Trabalho'], sign=True)} p.p.; leitura INE melhora — desfasamento OIT", "Inconsistente — vintage + produtividade"),
            ("Saúde e Seg. Alimentar", "Mortalidade <5 69 → 51; materna 199 → 165; vacinação 76% vs 80%", f"{f1(d23['Saúde/Alimentar'], sign=True)} p.p.; vacinal −10,0, malária e desnutrição pesam", "Consistente — alerta principal"),
            ("Diversificação", "Não petrolífero 4,6%/ano (5,2% em 2025); IDE e exportações aquém", f"{f1(d23['Diversificação'], sign=True)} p.p.; PIB n.p. +10,8; crédito −1,1", "Parcial")]
    gt = new.shapes.add_table(len(rows) + 1, 4, Emu(559886), Emu(1400000), Emu(11100000), Emu(4200000)).table
    widths = [Emu(2000000), Emu(3900000), Emu(3400000), Emu(1800000)]
    for j, w in enumerate(widths): gt.columns[j].width = w
    hdrs = ["Dimensão", "PDN 2023-2027 e Balanço", "IGDA 2023–2025 (v8)", "Veredicto"]
    for j, h in enumerate(hdrs):
        c = gt.cell(0, j); c.text = h
        for pp in c.text_frame.paragraphs:
            for r in pp.runs: r.font.size = Pt(11); r.font.bold = True
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            c = gt.cell(i, j); c.text = v
            for pp in c.text_frame.paragraphs:
                for r in pp.runs: r.font.size = Pt(9)
    note = new.shapes.add_textbox(Emu(559886), Emu(5680000), Emu(11100000), Emu(420000)); nf = note.text_frame; nf.word_wrap = True
    np_ = nf.paragraphs[0]; np_.text = ("Fontes: PDN 2023-2027 (Diário da República); MINPLAN, Balanço do PDN (I trim e anual 2025: 1 036 indicadores, 856 sem execução reportada); INE (CN 2025, IPCN, IEA); FMI Art. IV 2026; construtor IGDA-BDA v8.")
    np_.runs[0].font.size = Pt(8); np_.runs[0].font.italic = True
    # mover o novo diapositivo para a posição 11 (após o 10)
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    new_el = ids[-1]
    sldIdLst.remove(new_el); sldIdLst.insert(10, new_el)
    prs.save(out)
    return out


if __name__ == "__main__":
    print(main("/home/user/bda-indice/Versão Final/Finalíssima/files/IGDA_Evolucao_CEX_v9.pptx", sys.argv[1]))
