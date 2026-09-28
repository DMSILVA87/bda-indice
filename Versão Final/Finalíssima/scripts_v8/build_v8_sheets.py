"""
Folhas novas do construtor v8:
  12_Emprego_INE  — leitura complementar do Mercado de Trabalho com as séries INE/IEA (feedback, ponto 1)
  13_Alteracoes_v8 — registo de alterações v7 → v8, comparação de resultados, metas alteradas e conversão WGI
Também: nota de versão em 09_Metodologia, regeneração das caches dos gráficos e opção de recálculo total ao abrir.
"""
from __future__ import annotations
import os, sys
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.chart.data_source import NumData, NumVal
from openpyxl.worksheet.hyperlink import Hyperlink

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from xlsx_tools import find_row_by_id, sheet_columns, BASE
from igda_engine import DIM_ORDER, DIM_SHORT
import v8_plan as P

YEARS = list(range(2015, 2026))
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SUBHDR_FILL = PatternFill("solid", fgColor="D9E1F2")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
H2_FONT = Font(bold=True, size=11, color="1F3864")
SUB_FONT = Font(italic=True, size=9, color="595959")
NOTE_FONT = Font(size=9, color="404040")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LINK_FONT = Font(size=10, color="1F4E79", underline="single")
SRC_FILL = PatternFill("solid", fgColor="F2F2F2")
DIM_MT = "6. Mercado de Trabalho"
# série alternativa -> série oficial que substitui (fronteiras e sentido da oficial, para que só os dados mudem)
ALT_PAIRS = [("LAB018", "LAB002", "Desemprego 15+ — INE harmonizada"), ("LAB028", "LAB001", "Taxa de emprego 15+ — INE harmonizada"),
             ("LAB029", "LAB003", "Desemprego jovem 15-24 — INE harmonizada"), ("LAB016", "LAB016", "Participação feminina — OIT (mantida)"),
             ("LAB025", "LAB025", "Produto por trabalhador — OIT (mantido)")]


def _fmt(v, vazio="vazio"):
    if v is None:
        return vazio
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return f"{v:g}"
    return str(v)


def _pt(x, nd=1, sign=False):
    if x is None:
        return "–"
    s = f"{x:+.{nd}f}" if sign else f"{x:.{nd}f}"
    return s.replace(".", ",").replace("-", "−")


def _hdr(ws, row, col, text, width=None):
    c = ws.cell(row, col, text); c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BORDER
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c


def _cell(ws, row, col, val, fmt=None, bold=False, fill=None, italic=False):
    c = ws.cell(row, col, val); c.border = BORDER; c.font = Font(size=10, bold=bold, italic=italic)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    return c


def _note(ws, r, text, height=40, last_col=10):
    ws.cell(r, 2, text).font = NOTE_FONT
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col); ws.row_dimensions[r].height = height
    ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")


def emprego_alt_diffs(model):
    """Replica em Python o bloco D da folha 12_Emprego_INE: subíndice alternativo (séries INE com as fronteiras das séries oficiais) vs oficial, 2019-2025."""
    rows = {r.id: r for r in model.rows}
    p = model.p
    out = {}
    for t, y in enumerate(YEARS):
        if y < 2019:
            continue
        scs = []
        for alt, ref, _ in ALT_PAIRS:
            ra, rr = rows[alt], rows[ref]
            v = ra.values[t]
            if v is None:
                continue
            lo, hi = float(rr.minimo), float(rr.maximo)
            x = (v - lo) / (hi - lo) if rr.sentido == "+" else (hi - v) / (hi - lo)
            scs.append(min(100.0, max(0.0, 100 * x)))
        alt = sum(scs) / len(scs) if len(scs) >= 4 else None
        off = model.subindices[DIM_MT][t]
        out[y] = dict(alt=alt, oficial=off, diff=(alt - off) if alt is not None and off is not None else None)
    return out


def add_emprego_ine(wb, labour: dict, web_new_method: dict, model):
    """labour: labour_data.json; web_new_method: {"2026T1": {...}, "2026T2": {...}}; model: igda_engine.Model já corrido sobre o livro v8."""
    ws = wb.create_sheet("12_Emprego_INE")
    base = wb[BASE]
    cols = sheet_columns(base)
    diffs = emprego_alt_diffs(model)
    ws.column_dimensions["A"].width = 3
    ws["A1"] = "← Painel"; ws["A1"].font = LINK_FONT; ws["A1"].hyperlink = Hyperlink(ref="A1", location="'00_Painel'!A1", display="← Painel")
    ws["B2"] = "LEITURA COMPLEMENTAR — MERCADO DE TRABALHO COM DADOS DO INE (IEA)"; ws["B2"].font = TITLE_FONT
    ws["B3"] = ("Resposta ao ponto 1 do feedback (Fable 5.1): o índice usa as estimativas modeladas da OIT (2015-2025, definição internacional/estrita) porque o Inquérito sobre o Emprego em Angola "
                "(IEA) só existe desde 2019 e, com o limiar de cobertura de 80%, as séries INE são inelegíveis para o índice-mãe. Esta folha mostra (A) as séries INE originais, (B) a nova metodologia "
                "em vigor desde o IV trim 2025, (C) as séries INE harmonizadas pela OIT (definição comparável) e (D) um subíndice alternativo 2019-2025 calculado com as séries INE normalizadas com as "
                "fronteiras das séries oficiais que substituem, para aproximar o efeito do desfasamento de vintage.")
    ws["B3"].font = SUB_FONT; ws["B3"].alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells("B3:N3"); ws.row_dimensions[3].height = 62
    ws.column_dimensions["B"].width = 62
    for i in range(3, 15):
        ws.column_dimensions[get_column_letter(i)].width = 10

    r = 5
    # ---------------- A ----------------
    ws.cell(r, 2, "A · Séries INE/IEA — metodologia antiga (13.ª CIET; definição nacional alargada de desemprego)").font = H2_FONT; r += 1
    yrs = list(range(2019, 2026))
    _hdr(ws, r, 2, "Indicador (população 15+, salvo indicação)")
    for j, y in enumerate(yrs): _hdr(ws, r, 3 + j, y)
    _hdr(ws, r, 3 + len(yrs), "Fonte", 44)
    r += 1
    ine = labour["ine_13ciet"]
    def ann(code, y):
        d = ine[code].get(str(y), {})
        if "Anual" in d and d["Anual"] is not None:
            return d["Anual"]
        q = [v for k, v in d.items() if k != "Anual" and v is not None]
        return sum(q) / len(q) if q else None
    seriesA = [("Taxa de desemprego (%)", "desemprego", {2025: 28.3}), ("Taxa de emprego (%)", "emprego", {2025: 63.7}), ("Taxa de desemprego jovem 15-24 (%)", "desemprego_jovem", {2025: 51.8}),
               ("Taxa de emprego informal (%)", "informal", {2025: 78.8}), ("Taxa de emprego formal (%)", "formal", {2025: 21.2}), ("Taxa de actividade (%)", "actividade", {})]
    for label, code, override in seriesA:
        _cell(ws, r, 2, label)
        for j, y in enumerate(yrs):
            v = override.get(y, ann(code, y))
            _cell(ws, r, 3 + j, None if v is None else round(v, 2), "0.0")
        _cell(ws, r, 3 + len(yrs), "INE, Séries cronológicas IEA (13.ª CIET), valor anual; 2025 = Anuário IEA 2025 (Abr-2026; metodologia antiga, média I-III trim)" if override else "INE, Séries cronológicas IEA (13.ª CIET), valor anual; 2025 = média I-III trim (metodologia antiga)", fill=SRC_FILL)
        r += 1
    _note(ws, r, "Notas: 2019-2022 e 2024 = valor anual publicado pelo INE (2019: inquérito iniciado no II trim; o valor anual do emprego informal/formal de 2019, 74,5%/25,3%, diverge da média dos trimestres II-IV, 79,5%/20,5%, "
                 "e formal + informal = 99,9%); 2023 = apenas IV trim publicado; 2025 = Anuário IEA 2025 (Abril 2026), que na metodologia antiga agrega apenas os trimestres I-III porque o IV trim já segue a nova metodologia (bloco B). "
                 "A definição nacional inclui a população desencorajada, pelo que o nível é cerca do dobro da definição estrita da OIT (bloco C).", height=54)
    r += 2

    # ---------------- B ----------------
    ws.cell(r, 2, "B · Nova metodologia INE (19.ª-21.ª CIET, em vigor desde o IV trim 2025) — quebra de série").font = H2_FONT; r += 1
    periods = ["IV trim 2025", "I trim 2026", "II trim 2026"]
    _hdr(ws, r, 2, "Indicador (população 15+)")
    for j, p_ in enumerate(periods): _hdr(ws, r, 3 + j, p_, 13)
    _hdr(ws, r, 3 + len(periods), "Fonte", 44); r += 1
    nova = labour["ine_nova_2025Q4"]
    rowsB = [("Taxa de desemprego (%)", nova.get("15+|Taxa de desemprego"), web_new_method.get("2026T1", {}).get("desemprego"), web_new_method.get("2026T2", {}).get("desemprego")),
             ("Taxa de emprego (%)", nova.get("15+|Taxa de emprego"), web_new_method.get("2026T1", {}).get("emprego"), web_new_method.get("2026T2", {}).get("emprego")),
             ("Taxa de desemprego jovem 15-24 (%)", nova.get("15-24|Taxa de desemprego"), web_new_method.get("2026T1", {}).get("desemprego_jovem"), web_new_method.get("2026T2", {}).get("desemprego_jovem")),
             ("Taxa de emprego informal (%)", nova.get("15+|Taxa de Emprego Informal"), web_new_method.get("2026T1", {}).get("informal"), web_new_method.get("2026T2", {}).get("informal")),
             ("Taxa da força de trabalho (%)", nova.get("15+|Taxa da força de trabalho"), web_new_method.get("2026T1", {}).get("actividade"), web_new_method.get("2026T2", {}).get("actividade"))]
    for label, *vals in rowsB:
        _cell(ws, r, 2, label)
        for j, v in enumerate(vals):
            _cell(ws, r, 3 + j, None if v is None else round(float(v), 2), "0.0")
        _cell(ws, r, 3 + len(periods), "INE, IEA IV trim 2025 (FIR); I e II trim 2026 (FIR, via Lusa/Expansão/Forbes África Lusófona)", fill=SRC_FILL)
        r += 1
    _note(ws, r, "A nova metodologia deixa de classificar a produção para autoconsumo como emprego: a taxa de emprego cai de ~63% para ~40% e o desemprego de 26,9% (III trim 2025) para 20,1% (IV trim 2025). "
                 "O INE não publicou retropolação; os valores não são comparáveis com o bloco A. A meta do PDN (25% em 2027) foi fixada na definição antiga.")
    r += 2

    # ---------------- C ----------------
    ws.cell(r, 2, "C · Séries INE/IEA harmonizadas pela OIT (definição internacional) vs estimativas modeladas usadas no índice — ligadas à Base_Potencial").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "Indicador")
    for j, y in enumerate(yrs): _hdr(ws, r, 3 + j, y)
    _hdr(ws, r, 3 + len(yrs), "Linha na Base_Potencial", 44); r += 1
    pairs = [("Desemprego 15+ — INE harmonizada (LAB018)", "LAB018"), ("Desemprego 15+ — OIT modelada, no índice (LAB002)", "LAB002"),
             ("Desemprego jovem 15-24 — INE harmonizada (LAB029)", "LAB029"), ("Desemprego jovem 15-24 — OIT modelada, no índice (LAB003)", "LAB003"),
             ("Taxa de emprego 15+ — INE harmonizada (LAB028)", "LAB028"), ("Taxa de emprego 15+ — OIT modelada, no índice (LAB001)", "LAB001"),
             ("Emprego informal — INE (LAB004)", "LAB004"), ("Taxa de formalização — INE (LAB005)", "LAB005")]
    for label, iid in pairs:
        br = find_row_by_id(base, iid)
        _cell(ws, r, 2, label)
        for j, y in enumerate(yrs):
            col = get_column_letter(cols[y])
            _cell(ws, r, 3 + j, f"=IF('{BASE}'!{col}{br}=\"\",\"\",'{BASE}'!{col}{br})", "0.0")
        _cell(ws, r, 3 + len(yrs), f"'{BASE}' linha {br} ({iid})", fill=SRC_FILL)
        r += 1
    _note(ws, r, "Leitura: a série INE harmonizada regista em 2025 a queda do desemprego (13,9% → 10,4%) que a estimativa modelada da OIT, produzida antes dos dados de 2025, ainda não incorpora (14,0% → 14,1%). "
                 "É este desfasamento de vintage que o feedback identifica; a parte 'substantiva' da divergência (produtividade em queda) mantém-se no bloco D.")
    r += 2

    # ---------------- D ----------------
    ws.cell(r, 2, "D · Subíndice Mercado de Trabalho 2019-2025 — leitura oficial (OIT modelada) vs leitura alternativa (séries INE harmonizadas, fronteiras das séries oficiais)").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "Linha")
    for j, y in enumerate(yrs): _hdr(ws, r, 3 + j, y)
    _hdr(ws, r, 3 + len(yrs), "Cálculo", 44); r += 1
    def score_formula(iid, ref_iid, y):
        br = find_row_by_id(base, iid); rr = find_row_by_id(base, ref_iid); col = get_column_letter(cols[y])
        v = f"'{BASE}'!{col}{br}"; lo = f"'{BASE}'!$G${rr}"; hi = f"'{BASE}'!$H${rr}"; s = f"'{BASE}'!$F${rr}"
        return f"=IF({v}=\"\",\"\",MIN(100,MAX(0,IF({s}=\"+\",100*({v}-{lo})/({hi}-{lo}),100*({hi}-{v})/({hi}-{lo})))))"
    first_score_row = r
    for alt, ref, label in ALT_PAIRS:
        rr = find_row_by_id(base, ref)
        lo, hi, s = base.cell(rr, cols["Mínimo"]).value, base.cell(rr, cols["Máximo"]).value, base.cell(rr, cols["Sentido"]).value
        _cell(ws, r, 2, f"Score · {label} ({alt}; fronteiras de {ref}: {lo:g}-{hi:g}, {s.replace('-', '−')})", italic=True)
        for j, y in enumerate(yrs):
            _cell(ws, r, 3 + j, score_formula(alt, ref, y), "0.0")
        _cell(ws, r, 3 + len(yrs), f"Min-max com as fronteiras e o sentido da série oficial ({ref}), para que só os dados mudem" if alt != ref else "Min-max com as fronteiras da Base_Potencial (série do índice)", fill=SRC_FILL); r += 1
    last_score_row = r - 1
    _cell(ws, r, 2, "Subíndice alternativo — leitura INE/IEA (média aritmética dos 5 scores)", bold=True)
    for j, y in enumerate(yrs):
        L = get_column_letter(3 + j)
        _cell(ws, r, 3 + j, f"=IF(COUNT({L}{first_score_row}:{L}{last_score_row})<4,\"\",AVERAGE({L}{first_score_row}:{L}{last_score_row}))", "0.0", bold=True)
    _cell(ws, r, 3 + len(yrs), "Mesma regra de cobertura mínima (≥75%)", fill=SRC_FILL); alt_row = r; r += 1
    _cell(ws, r, 2, "Subíndice oficial — Mercado de Trabalho (07_Subindices)", bold=True)
    for j, y in enumerate(yrs):
        col07 = get_column_letter(6 + (y - 2015))
        _cell(ws, r, 3 + j, f"=IF('07_Subindices'!{col07}10=\"\",\"\",'07_Subindices'!{col07}10)", "0.0", bold=True)
    _cell(ws, r, 3 + len(yrs), "Ligado a 07_Subindices (linha Mercado Trabalho)", fill=SRC_FILL); off_row = r; r += 1
    _cell(ws, r, 2, "Diferença (alternativa − oficial), p.p.")
    for j, y in enumerate(yrs):
        L = get_column_letter(3 + j)
        _cell(ws, r, 3 + j, f"=IF(OR({L}{alt_row}=\"\",{L}{off_row}=\"\"),\"\",{L}{alt_row}-{L}{off_row})", "+0.0;-0.0;0.0")
    r += 1
    _cell(ws, r, 2, "Variação 2023→2025 — alternativa vs oficial (p.p.)")
    _cell(ws, r, 3, f"=IF(OR(I{alt_row}=\"\",G{alt_row}=\"\"),\"\",I{alt_row}-G{alt_row})", "+0.0;-0.0;0.0")
    _cell(ws, r, 4, f"=IF(OR(I{off_row}=\"\",G{off_row}=\"\"),\"\",I{off_row}-G{off_row})", "+0.0;-0.0;0.0")
    _cell(ws, r, 5, "← coluna C: alternativa | coluna D: oficial", italic=True)
    r += 2
    d = diffs
    prev = [d[y]["diff"] for y in range(2019, 2025) if d[y]["diff"] is not None]
    mean_prev = sum(prev) / len(prev) if prev else None
    seq = "; ".join(f"{y}: {_pt(d[y]['diff'], 1, sign=True)}" for y in range(2019, 2026) if d[y]["diff"] is not None)
    _note(ws, r, ("Como ler: a leitura alternativa substitui as três séries modeladas da OIT (emprego, desemprego, desemprego jovem) pelas séries INE/IEA harmonizadas, normalizadas com as fronteiras das séries que substituem, "
                  f"mantendo a participação feminina e a produtividade. Diferença alternativa − oficial (p.p.): {seq}. Mesmo com fronteiras iguais, a diferença não é nula em 2019-2024 (média {_pt(mean_prev, 1, sign=True)} p.p.; "
                  "reflecte diferenças de nível e de definição entre as séries harmonizadas e as modeladas, e em 2023 a publicação de um único trimestre pelo INE). "
                  f"O efeito do desfasamento de vintage é aproximado pelo salto de 2025 ({_pt(d[2025]['diff'], 1, sign=True)} p.p.) face a essa média, e não pela diferença bruta. "
                  "Nenhuma célula desta folha altera o índice-mãe; quando a OIT publicar estimativas modeladas que incorporem o IEA 2025, a linha oficial deverá aproximar-se da alternativa."), height=80)
    ws.freeze_panes = "C5"
    return ws, diffs


def add_changelog(wb, log: list, results_v7: dict, results_v8: dict, meta_changes: list, transposicoes: dict | None = None, version_label="v8", date_label="28-09-2026"):
    """results_*: {"dims": {short: [11 values]}, "igda": [11]}; meta_changes: (id, antes, depois, origem, categoria)."""
    ws = wb.create_sheet(f"13_Alteracoes_{version_label}")
    ws.column_dimensions["A"].width = 3; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 14; ws.column_dimensions["D"].width = 9
    ws["A1"] = "← Painel"; ws["A1"].font = LINK_FONT; ws["A1"].hyperlink = Hyperlink(ref="A1", location="'00_Painel'!A1", display="← Painel")
    ws["B2"] = f"REGISTO DE ALTERAÇÕES v7 → {version_label}  ·  {date_label}"; ws["B2"].font = TITLE_FONT
    ws["B3"] = ("Actualização de dados (INE, MINFIN/FMI, MINPLAN, INSS/MAPTSS; séries internacionais mantidas no vintage da v7), alinhamento das metas 2027 com o PDN 2023-2027, "
                "novos candidatos de emprego (INE/IEA) e protecção social, correcção de séries sem rastreabilidade e resposta ao feedback Fable 5.1. "
                "Todas as fórmulas do construtor foram preservadas; o livro recalcula integralmente ao abrir. "
                "Regras de alinhamento das metas 2027 (a coluna 'Meta 2027' não entra em nenhuma fórmula), em seis categorias identificadas entre parênteses rectos na coluna 'Origem da Meta 2027': "
                "(i) valor do PDN aplicado directamente quando a base coincide com a série, o indicador é o do PDN ou a série não permite confrontar a base; "
                "(ii) TRANSPOSIÇÃO quando a base 2022 do PDN difere do valor 2022 da série (fonte, definição ou vintage distintos): meta = valor2022 + (meta_PDN − base_PDN) para níveis e proporções; "
                "meta = valor2022 × meta_PDN ÷ base_PDN para taxas de mortalidade, de desemprego e rácios de dívida; quando a série não tem observação em 2022 usa-se a observação mais próxima "
                "(HUM003 2023; HUM024 2021; INC006 2018; HUM025 e SAU029 2024) e a variação é aplicada integralmente; LAB004 = 100 − meta de LAB005; as metas em % do PIB não petrolífero são primeiro convertidas para % do PIB (×0,8) e depois transpostas; "
                "(iii) CONVERSÃO de escala: percentis WGI → estimativas, meta = ê2022 + [Φ⁻¹(p2027) − Φ⁻¹(p2022)] (secção 4); IDE não petrolífero em % do PIB não petrolífero → % do PIB; "
                "(iv) meta anual 2025 do Balanço do PDN (MINPLAN/MINSA) como proxy quando o PDN não fixa 2027, transposta à série do índice (sarampo, malária); "
                "(v) meta operacional mantida ou ajustada quando não há correspondência; (vi) sem meta quando o conceito do PDN difere do da série. "
                "Em seis indicadores do índice (HUM002, HUM006, SAU04, INF001, INF002, LAB002) a última observação já atinge o valor absoluto do PDN; a meta transposta é deliberadamente mais exigente porque preserva a variação, e não o nível, do plano.")
    ws["B3"].font = SUB_FONT; ws["B3"].alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells("B3:P3"); ws.row_dimensions[3].height = 150
    r = 5
    ws.cell(r, 2, "1 · Comparação de resultados (subíndices e IGDA)").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "Dimensão"); _hdr(ws, r, 3, "Versão")
    for j, y in enumerate(YEARS): _hdr(ws, r, 4 + j, y, 9)
    _hdr(ws, r, 4 + len(YEARS), "Δ 2015-25", 10); r += 1
    for short in list(results_v7["dims"].keys()) + ["IGDA-BDA"]:
        for ver, res in (("v7", results_v7), (version_label, results_v8)):
            vals = res["igda"] if short == "IGDA-BDA" else res["dims"][short]
            _cell(ws, r, 2, short, bold=(short == "IGDA-BDA")); _cell(ws, r, 3, ver, bold=(ver != "v7"))
            for j, v in enumerate(vals):
                _cell(ws, r, 4 + j, None if v is None else round(v, 2), "0.0", bold=(ver != "v7"), fill=(SUBHDR_FILL if ver != "v7" else None))
            if vals[0] is not None and vals[-1] is not None:
                _cell(ws, r, 4 + len(YEARS), round(vals[-1] - vals[0], 2), "+0.0;-0.0;0.0", bold=(ver != "v7"), fill=(SUBHDR_FILL if ver != "v7" else None))
            r += 1
    _note(ws, r, "Valores v7 = recálculo integral das fórmulas da v7 (o ficheiro v7 distribuído continha valores em cache da v6, ex.: IGDA 2025 = 46,4 em vez de 45,7). "
                 f"Valores {version_label} = recálculo integral após as alterações abaixo (motor de cálculo independente que replica as fórmulas do construtor).", height=28, last_col=16)
    r += 2
    ws.cell(r, 2, "2 · Alterações linha a linha").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "Tipo"); _hdr(ws, r, 3, "ID"); _hdr(ws, r, 4, "Descrição"); r += 1
    for e in log:
        _cell(ws, r, 2, e["tipo"]); _cell(ws, r, 3, e["id"])
        c = _cell(ws, r, 4, e["descricao"]); c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=16)
        ws.row_dimensions[r].height = max(15, min(160, 15 * (1 + len(e["descricao"]) // 140)))
        r += 1
    r += 1
    ws.cell(r, 2, "3 · Metas 2027 alteradas (alinhamento com o PDN)").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "ID"); _hdr(ws, r, 3, "v7 → v8"); _hdr(ws, r, 4, "Categoria"); _hdr(ws, r, 7, "Origem"); r += 1
    for mid, before, after, origem, cat in meta_changes:
        _cell(ws, r, 2, mid); _cell(ws, r, 3, f"{_fmt(before, 'sem meta')} → {_fmt(after, 'sem meta')}")
        c = _cell(ws, r, 4, P.META_CAT_LABEL[cat]); ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        c = _cell(ws, r, 7, origem); c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=16)
        ws.row_dimensions[r].height = max(15, min(75, 15 * (1 + len(origem) // 120)))
        r += 1
    r += 1
    ws.cell(r, 2, "4 · Conversão das metas de governança: percentis WGI (PDN p.30) → estimativas WGI").font = H2_FONT; r += 1
    for j, h in enumerate(["ID", "Indicador", "", "p2022 (%)", "p2027 (%)", "ê2022", "Meta 2027", "Sensib. (p.p.)"]):
        if h: _hdr(ws, r, 2 + j, h)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4); r += 1
    for gid, (nome, p22, p27, e22) in P.WGI_PDN.items():
        _cell(ws, r, 2, gid); _cell(ws, r, 3, nome); ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        _cell(ws, r, 5, p22, "0.0"); _cell(ws, r, 6, p27, "0.0"); _cell(ws, r, 7, e22, "0.00"); _cell(ws, r, 8, P.METAS[gid][0], "0.00", bold=True); _cell(ws, r, 9, P.wgi_meta_pp(e22, p22, p27), "0.00")
        r += 1
    _note(ws, r, "meta = ê2022 + [Φ⁻¹(p2027) − Φ⁻¹(p2022)], em que ê2022 é a estimativa WGI da série (edição 2025) e p são os percentis do PDN (p.30, '2022 ou ano mais recente disponível'). "
                 "Os percentile ranks do WGI são posições empíricas de um vintage anterior, pelo que Φ(ê2022) difere de p2022 (ex.: 29,4% vs 20,8% em GOV004); o deslocamento em z é invariante a essa diferença de nível. "
                 "Sensibilidade: aplicar o deslocamento em pontos percentuais ao percentil implícito Φ(ê2022) daria as metas da última coluna (diferença ≤0,04 na escala WGI).", height=54, last_col=16)
    ws.freeze_panes = "B5"
    return ws


def add_version_notes(wb, n_cand: int, version_label="v8", date_label="Setembro de 2026"):
    ws = wb["09_Metodologia"]
    r = ws.max_row + 2
    ws.cell(r, 2, f"10. REVISÃO {version_label.upper()} ({date_label})").font = Font(bold=True, size=11, color="1F3864"); r += 1
    notes = [
        f"  Catálogo: {n_cand} candidatos (331 na v7 + 7 na v8). Dados: séries actualizadas com fontes primárias nacionais (INE Contas Nacionais preliminares 2025 e trimestrais; INE IEA Anuário 2025; INE IPCN; FMI WEO Abr-2026, que compila dados do MINFIN; MINPLAN Balanço PDN 2025; INSS/MAPTSS). As séries internacionais (WDI, WGI, OIT, OMS/UNICEF, UNCTAD, TI, WEF, IPU) mantêm o vintage da v7, não reverificado nesta revisão.",
        "  Metas 2027: coluna I alinhada com o PDN 2023-2027 em seis categorias — valor directo (base coincide, indicador do PDN ou base não confrontável); transposição à base da série quando a base 2022 do PDN difere (variação aditiva para níveis/proporções, relativa para mortalidade, desemprego e dívida; observação mais próxima quando não há 2022; LAB004 = 100 − LAB005; metas em % do PIB não petrolífero convertidas ×0,8 antes de transpor); conversão de escala (percentis WGI → estimativas); meta anual 2025 do MINPLAN como proxy, transposta à série; meta operacional mantida ou ajustada; sem meta quando o conceito difere. A coluna 'Origem da Meta 2027' identifica a categoria [entre parênteses rectos] e a fonte de cada meta; regras e fórmulas em 13_Alteracoes_v8. A coluna I não entra em nenhuma fórmula.",
        "  Marcador de fonte nacional (+12 pontos): assinala as séries cujos dados de base são produzidos por instituições nacionais, incluindo estimativas de organismos internacionais construídas sobre esses dados (ex.: estimativas modeladas da OIT a partir do IEA; dívida e saldo do MINFIN via FMI WEO). A composição de Macroeconomia é sensível a esta convenção (ver Nota v8, 9.2).",
        "  Emprego (feedback ponto 1): novas linhas LAB028-LAB031 com as séries INE/IEA (originais e harmonizadas pela OIT); folha 12_Emprego_INE com leitura complementar 2019-2025. O índice-mãe mantém as séries OIT 2015-2025 por exigência de cobertura.",
        "  Inclusão (feedback ponto 2): pobreza a 2,15 USD (PDN) e 3,00 USD (novo padrão BM) separadas (INC001/INC031); novas linhas INC032 (cobertura do Kwenda) e INC033 (segurados inscritos no INSS, milhões). Continuam inelegíveis por cobertura temporal — documentadas para incorporação futura.",
        "  Correcções: séries 'Compilação interna BDA' de dívida e saldo orçamental substituídas pelo FMI WEO Abr-2026 (marcador de fonte nacional mantido: produtor primário MINFIN); mortalidade materna 2024-2025 sem suporte na série MMEIG removida (a estimativa do IIMS 2023-24 fica em SAU029 como validação cruzada); valores em cache das células (v7, herdados da v6) eliminados e caches dos gráficos regeneradas com os resultados v8 — o livro recalcula integralmente ao abrir.",
        f"  Registo completo em 13_Alteracoes_{version_label}.",
    ]
    for n in notes:
        ws.cell(r, 2, n).alignment = Alignment(wrap_text=True, vertical="top"); r += 1
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for c in row:
            if isinstance(c.value, str) and c.value.strip() == "IGDA-BDA · Selecção Final":
                c.value = f"IGDA-BDA · Selecção Final · {version_label} ({date_label})"
    wb.calculation = CalcProperties(fullCalcOnLoad=True)


def refresh_chart_caches(wb, model):
    """Substitui as caches numéricas dos gráficos (herdadas da v6) pelos resultados v8 do motor; onde o intervalo não é reconhecido, remove a cache."""
    sub = [model.subindices[d] for d in DIM_ORDER]
    igda = list(model.igda)
    inc = [model.p.dims_incluir.get(d, 0) == 1 for d in DIM_ORDER]
    ranges = {f"'07_Subindices'!$F${5 + i}:$P${5 + i}": sub[i] for i in range(11)}
    ranges["'07_Subindices'!$F$18:$P$18"] = igda
    ranges["'07_Subindices'!$T$5:$T$15"] = [(sub[i][10] if inc[i] else None) for i in range(11)]
    v8 = wb["08_Visualizacoes"]
    ya, yb = v8["C23"].value, v8["F23"].value
    if isinstance(ya, int) and isinstance(yb, int) and 2015 <= ya <= 2025 and 2015 <= yb <= 2025:
        A, B = ya - 2015, yb - 2015
        ranges["'08_Visualizacoes'!$C$26:$C$33"] = [sub[i][A] for i in range(8)]
        ranges["'08_Visualizacoes'!$D$26:$D$33"] = [sub[i][B] for i in range(8)]
        ranges["'08_Visualizacoes'!$E$26:$E$33"] = [(sub[i][B] - sub[i][A]) if sub[i][A] is not None and sub[i][B] is not None else None for i in range(8)]
    n_set = n_clear = 0
    for ws in wb.worksheets:
        for ch in ws._charts:
            for s in ch.series:
                ds = s.val
                if ds is None or ds.numRef is None:
                    continue
                vals = ranges.get(ds.numRef.f)
                if vals is None:
                    ds.numRef.numCache = None; n_clear += 1
                else:
                    fc = ds.numRef.numCache.formatCode if ds.numRef.numCache is not None else None
                    ds.numRef.numCache = NumData(formatCode=fc, ptCount=len(vals), pt=[NumVal(idx=i, v=round(v, 6)) for i, v in enumerate(vals) if v is not None]); n_set += 1
    return n_set, n_clear
