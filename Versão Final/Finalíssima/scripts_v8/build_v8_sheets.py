"""
Folhas novas do construtor v8:
  12_Emprego_INE  — leitura complementar do Mercado de Trabalho com as séries INE/IEA (feedback, ponto 1)
  13_Alteracoes_v8 — registo de alterações v7 → v8 e comparação de resultados
Também: nota de versão em 09_Metodologia e opção de recálculo total ao abrir.
"""
from __future__ import annotations
import os
import json
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_tools import find_row_by_id, sheet_columns, BASE

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


def add_emprego_ine(wb, labour: dict, web_new_method: dict):
    """labour: labour_data.json; web_new_method: {"2026T1": {...}, "2026T2": {...}} valores da nova metodologia recolhidos na web."""
    ws = wb.create_sheet("12_Emprego_INE")
    base = wb[BASE]
    cols = sheet_columns(base)
    ws.column_dimensions["A"].width = 3
    ws["A1"] = "← Painel"; ws["A1"].font = LINK_FONT; ws["A1"].hyperlink = "#'00_Painel'!A1"
    ws["B2"] = "LEITURA COMPLEMENTAR — MERCADO DE TRABALHO COM DADOS DO INE (IEA)"; ws["B2"].font = TITLE_FONT
    ws["B3"] = ("Resposta ao ponto 1 do feedback (Fable 5.1): o índice usa as estimativas modeladas da OIT (2015-2025, definição internacional/estrita) porque o Inquérito sobre o Emprego em Angola "
                "(IEA) só existe desde 2019 e, com o limiar de cobertura de 80%, as séries INE são inelegíveis para o índice-mãe. Esta folha mostra (A) as séries INE originais, (B) a nova metodologia "
                "em vigor desde o IV trim 2025, (C) as séries INE harmonizadas pela OIT (definição comparável) e (E) um subíndice alternativo 2019-2025 calculado com as séries INE, para isolar o desfasamento de vintage.")
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
        _cell(ws, r, 3 + len(yrs), "INE, Séries cronológicas IEA (13.ª CIET); 2025 = Anuário IEA 2025 (Abr-2026)" if override else "INE, Séries cronológicas IEA (13.ª CIET); 2025 = média I-III trim", fill=SRC_FILL)
        r += 1
    ws.cell(r, 2, "Notas: 2019 = média II-IV trim (início do inquérito); 2023 = apenas IV trim publicado; 2024 = valor anual INE. A definição nacional inclui a população desencorajada, "
                  "pelo que o nível é cerca do dobro da definição estrita da OIT (bloco C).").font = NOTE_FONT
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10); ws.row_dimensions[r].height = 28; ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    r += 2

    # ---------------- B ----------------
    ws.cell(r, 2, "B · Nova metodologia INE (19.ª-21.ª CIET, em vigor desde o IV trim 2025) — quebra de série").font = H2_FONT; r += 1
    periods = ["IV trim 2025", "I trim 2026", "II trim 2026"]
    _hdr(ws, r, 2, "Indicador (população 15+)")
    for j, p in enumerate(periods): _hdr(ws, r, 3 + j, p, 13)
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
    ws.cell(r, 2, "A nova metodologia deixa de classificar a produção para autoconsumo como emprego: a taxa de emprego cai de ~63% para ~40% e o desemprego de 26,9% (III trim 2025) para 20,1% (IV trim 2025). "
                  "O INE não publicou retropolação; os valores não são comparáveis com o bloco A. A meta do PDN (25% em 2027) foi fixada na definição antiga.").font = NOTE_FONT
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10); ws.row_dimensions[r].height = 40; ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
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
    rowref = {}
    for label, iid in pairs:
        br = find_row_by_id(base, iid); rowref[iid] = br
        _cell(ws, r, 2, label)
        for j, y in enumerate(yrs):
            col = get_column_letter(cols[y])
            c = _cell(ws, r, 3 + j, f"=IF('{BASE}'!{col}{br}=\"\",\"\",'{BASE}'!{col}{br})", "0.0")
        _cell(ws, r, 3 + len(yrs), f"'{BASE}' linha {br} ({iid})", fill=SRC_FILL)
        r += 1
    ws.cell(r, 2, "Leitura: a série INE harmonizada regista em 2025 a queda do desemprego (13,9% → 10,4%) que a estimativa modelada da OIT, produzida antes dos dados de 2025, ainda não incorpora (14,0% → 14,1%). "
                  "É este desfasamento de vintage que o feedback identifica; a parte 'substantiva' da divergência (produtividade em queda) mantém-se no bloco E.").font = NOTE_FONT
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10); ws.row_dimensions[r].height = 40; ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    r += 2

    # ---------------- E ----------------
    ws.cell(r, 2, "E · Subíndice Mercado de Trabalho 2019-2025 — leitura oficial (OIT modelada) vs leitura alternativa (séries INE harmonizadas)").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "Linha")
    for j, y in enumerate(yrs): _hdr(ws, r, 3 + j, y)
    _hdr(ws, r, 3 + len(yrs), "Cálculo", 44); r += 1
    # scores normalizados alternativos: LAB018 (−), LAB028 (+), LAB029 (−), LAB016 (+, modelada), LAB025 (+)
    def score_formula(iid, y):
        br = find_row_by_id(base, iid); col = get_column_letter(cols[y])
        v = f"'{BASE}'!{col}{br}"; lo = f"'{BASE}'!$G${br}"; hi = f"'{BASE}'!$H${br}"; s = f"'{BASE}'!$F${br}"
        return f"=IF({v}=\"\",\"\",MIN(100,MAX(0,IF({s}=\"+\",100*({v}-{lo})/({hi}-{lo}),100*({hi}-{v})/({hi}-{lo})))))"
    alt_ids = [("Score · Desemprego INE harmonizada (LAB018, 0-50, −)", "LAB018"), ("Score · Taxa de emprego INE harmonizada (LAB028, 0-100, +)", "LAB028"),
               ("Score · Desemprego jovem INE harmonizada (LAB029, 0-50, −)", "LAB029"), ("Score · Participação feminina OIT (LAB016, 0-100, +)", "LAB016"),
               ("Score · Produto por trabalhador OIT (LAB025, 0-25 000, +)", "LAB025")]
    first_score_row = r
    for label, iid in alt_ids:
        _cell(ws, r, 2, label, italic=True)
        for j, y in enumerate(yrs):
            _cell(ws, r, 3 + j, score_formula(iid, y), "0.0")
        _cell(ws, r, 3 + len(yrs), "Min-max com as fronteiras da Base_Potencial", fill=SRC_FILL); r += 1
    last_score_row = r - 1
    _cell(ws, r, 2, "Subíndice alternativo — leitura INE/IEA (média aritmética dos 5 scores)", bold=True)
    for j, y in enumerate(yrs):
        L = get_column_letter(3 + j)
        _cell(ws, r, 3 + j, f"=IF(COUNT({L}{first_score_row}:{L}{last_score_row})<4,\"\",AVERAGE({L}{first_score_row}:{L}{last_score_row}))", "0.0", bold=True)
    _cell(ws, r, 3 + len(yrs), "Mesma regra de cobertura mínima (≥75%)", fill=SRC_FILL); alt_row = r; r += 1
    _cell(ws, r, 2, "Subíndice oficial — Mercado de Trabalho (07_Subindices)", bold=True)
    # 07_Subindices: Mercado Trabalho row 10, years 2015..2025 in F..P -> 2019 = J
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
    _cell(ws, r, 5, "← alternativa (C) | oficial (D)", italic=True)
    r += 2
    ws.cell(r, 2, "Como ler: a leitura alternativa substitui as três séries modeladas da OIT (emprego, desemprego, desemprego jovem) pelas séries INE/IEA harmonizadas, mantendo a participação feminina e a produtividade. "
                  "A diferença entre as duas linhas em 2025 mede o efeito de vintage; a evolução comum (produtividade em queda) é a divergência substantiva face à leitura quantitativa do PDN. "
                  "Nenhuma célula desta folha altera o índice-mãe; quando a OIT publicar estimativas modeladas que incorporem o IEA 2025, a linha oficial converge para a alternativa.").font = NOTE_FONT
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10); ws.row_dimensions[r].height = 54; ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "C5"
    return ws


def add_changelog(wb, log: list, results_v7: dict, results_v8: dict, meta_changes: list, version_label="v8", date_label="28-09-2026"):
    """results_*: {"dims": {short: [11 values]}, "igda": [11]}"""
    ws = wb.create_sheet(f"13_Alteracoes_{version_label}")
    ws.column_dimensions["A"].width = 3; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 12; ws.column_dimensions["D"].width = 120
    ws["A1"] = "← Painel"; ws["A1"].font = LINK_FONT; ws["A1"].hyperlink = "#'00_Painel'!A1"
    ws["B2"] = f"REGISTO DE ALTERAÇÕES v7 → {version_label}  ·  {date_label}"; ws["B2"].font = TITLE_FONT
    ws["B3"] = ("Actualização de dados (INE, MINFIN/FMI, MINPLAN, BNA, Banco Mundial, OIT, OMS/UNICEF, UNCTAD, TI, WEF, IPU), alinhamento das metas 2027 com o PDN 2023-2027, "
                "novos candidatos de emprego (INE/IEA) e protecção social, correcção de séries sem rastreabilidade e resposta ao feedback Fable 5.1. "
                "Todas as fórmulas do construtor foram preservadas; o livro recalcula integralmente ao abrir.")
    ws["B3"].font = SUB_FONT; ws["B3"].alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells("B3:D3"); ws.row_dimensions[3].height = 44
    r = 5
    ws.cell(r, 2, "1 · Comparação de resultados (subíndices e IGDA)").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "Dimensão"); _hdr(ws, r, 3, "Versão")
    for j, y in enumerate(YEARS): _hdr(ws, r, 4 + j, y, 9)
    _hdr(ws, r, 4 + len(YEARS), "Δ 2015-25", 10); r += 1
    ws.column_dimensions["D"].width = 9
    for short in list(results_v7["dims"].keys()) + ["IGDA-BDA"]:
        for ver, res in (("v7", results_v7), (version_label, results_v8)):
            vals = res["igda"] if short == "IGDA-BDA" else res["dims"][short]
            _cell(ws, r, 2, short, bold=(short == "IGDA-BDA")); _cell(ws, r, 3, ver, bold=(ver != "v7"))
            for j, v in enumerate(vals):
                _cell(ws, r, 4 + j, None if v is None else round(v, 2), "0.0", bold=(ver != "v7"), fill=(SUBHDR_FILL if ver != "v7" else None))
            if vals[0] is not None and vals[-1] is not None:
                _cell(ws, r, 4 + len(YEARS), round(vals[-1] - vals[0], 2), "+0.0;-0.0;0.0", bold=(ver != "v7"), fill=(SUBHDR_FILL if ver != "v7" else None))
            r += 1
    ws.cell(r, 2, "Valores v7 = recálculo integral das fórmulas da v7 (o ficheiro v7 distribuído continha valores em cache da v6, ex.: IGDA 2025 = 46,4 em vez de 45,7). "
                  f"Valores {version_label} = recálculo integral após as alterações abaixo.").font = NOTE_FONT
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=16); ws.row_dimensions[r].height = 28; ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    r += 2
    ws.cell(r, 2, "2 · Alterações linha a linha").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "Tipo"); _hdr(ws, r, 3, "ID"); _hdr(ws, r, 4, "Descrição"); r += 1
    ws.column_dimensions["D"].width = 9
    # descrição ocupa D:P (merge) para largura
    for e in log:
        _cell(ws, r, 2, e["tipo"]); _cell(ws, r, 3, e["id"])
        c = _cell(ws, r, 4, e["descricao"]); c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=16)
        ws.row_dimensions[r].height = max(15, min(120, 15 * (1 + len(e["descricao"]) // 140)))
        r += 1
    r += 1
    ws.cell(r, 2, "3 · Metas 2027 alteradas (alinhamento PDN)").font = H2_FONT; r += 1
    _hdr(ws, r, 2, "ID"); _hdr(ws, r, 3, "v7 → v8"); _hdr(ws, r, 4, "Origem"); r += 1
    for mid, before, after, origem in meta_changes:
        _cell(ws, r, 2, mid); _cell(ws, r, 3, f"{before} → {after}")
        c = _cell(ws, r, 4, origem); c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=16)
        ws.row_dimensions[r].height = max(15, min(60, 15 * (1 + len(origem) // 140)))
        r += 1
    ws.freeze_panes = "B5"
    return ws


def add_version_notes(wb, version_label="v8", date_label="Setembro de 2026"):
    ws = wb["09_Metodologia"]
    r = ws.max_row + 2
    ws.cell(r, 2, f"10. REVISÃO {version_label.upper()} ({date_label})").font = Font(bold=True, size=11, color="1F3864"); r += 1
    notes = [
        "  Dados: séries actualizadas com as publicações mais recentes (INE Contas Nacionais preliminares 2025 e trimestrais; INE IEA Anuário 2025; INE IPCN; FMI WEO Abr-2026; MINPLAN Balanço PDN 2025; BNA; Banco Mundial WDI; OIT; OMS/UNICEF; UNCTAD; TI; WEF; IPU).",
        "  Metas 2027: coluna I alinhada com o PDN 2023-2027 (metas oficiais) sempre que existe correspondência; a nova coluna 'Origem da Meta 2027' documenta a fonte ou a razão da manutenção da meta operacional.",
        "  Emprego (feedback ponto 1): novas linhas LAB028-LAB031 com as séries INE/IEA (originais e harmonizadas pela OIT); folha 12_Emprego_INE com leitura complementar 2019-2025. O índice-mãe mantém as séries OIT 2015-2025 por exigência de cobertura.",
        "  Inclusão (feedback ponto 2): pobreza a 2,15 USD (PDN) e 3,00 USD (novo padrão BM) separadas (INC001/INC031); nova linha INC032 (cobertura do Kwenda). Continuam inelegíveis por cobertura temporal — documentadas para incorporação futura.",
        "  Correcções: séries 'Compilação interna BDA' de dívida e saldo orçamental substituídas pelo FMI WEO Abr-2026; mortalidade materna 2025 sem suporte removida (IIMS 2023-24 = 170 em 2024); valores em cache desactualizados da v7 (v6) eliminados — o livro recalcula ao abrir.",
        f"  Registo completo em 13_Alteracoes_{version_label}.",
    ]
    for n in notes:
        ws.cell(r, 2, n).alignment = Alignment(wrap_text=True, vertical="top"); r += 1
    # versão no rodapé
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for c in row:
            if isinstance(c.value, str) and c.value.strip() == "IGDA-BDA · Selecção Final":
                c.value = f"IGDA-BDA · Selecção Final · {version_label} ({date_label})"
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
