"""
Constrói IGDA_BDA_Construtor_v8.xlsx a partir da v7, aplicando v8_plan.py:
  1. copia o livro v7 (fórmulas intactas)
  2. alarga os intervalos $5:$335 -> $5:$<nova última linha>
  3. actualiza séries, metas, metadados e notas nas linhas existentes
  4. acrescenta novas linhas de candidatos (fórmulas traduzidas da linha-modelo)
  5. acrescenta coluna "Origem da Meta 2027" (AT) e preenche
  6. acrescenta linhas em 10_Fontes para os novos indicadores
  7. cria folhas 12_Emprego_INE (leitura complementar) e 13_Alteracoes_v8 (changelog)
  8. actualiza textos de versão (00_Painel, 09_Metodologia)
  9. grava; devolve o log de alterações (lista de dicts) para o relatório
"""
from __future__ import annotations
import os
import json, sys, shutil, datetime
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_tools import *  # noqa
import v8_plan as P

SRC = "/home/user/bda-indice/Versão Final/Finalíssima/files/IGDA_BDA_Construtor_v7.xlsx"
YEARS = list(range(2015, 2026))

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
SUB_FONT = Font(italic=True, size=9, color="595959")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
EDIT_FILL = PatternFill("solid", fgColor="FFF2CC")
NOTE_FONT = Font(size=9, color="404040")


def build(out_path: str, extra_series: dict | None = None, extra_rows: list | None = None, extra_metas: dict | None = None,
          extra_rename: dict | None = None, changelog_extra: list | None = None, version_label="v8", data_label="Setembro de 2026"):
    log = []
    wb = openpyxl.load_workbook(SRC)
    ws = wb[BASE]
    cols = sheet_columns(ws)
    old_last = last_data_row(ws)

    series = dict(P.SERIES); series.update(extra_series or {})
    rows_new = list(P.NEW_ROWS) + list(extra_rows or [])
    metas = dict(P.METAS); metas.update(extra_metas or {})
    rename = dict(P.RENAME); rename.update(extra_rename or {})

    new_last = old_last + len(rows_new)
    n = extend_ranges(wb, old_last, new_last)
    log.append(dict(tipo="estrutura", id="", descricao=f"Intervalos da Base_Potencial alargados de $5:${old_last} para $5:${new_last} ({n} fórmulas actualizadas)"))

    # ---- coluna Origem da Meta 2027 (AT) ----
    meta_col = ws.max_column + 1
    hdr_src = ws.cell(HDR_ROW, ws.max_column)
    c = ws.cell(HDR_ROW, meta_col, "Origem da Meta 2027")
    copy_style(hdr_src, c)
    ws.column_dimensions[get_column_letter(meta_col)].width = 70
    for r in range(FIRST, new_last + 1):
        copy_style(ws.cell(r, ws.max_column - 1), ws.cell(r, meta_col))
    log.append(dict(tipo="estrutura", id="", descricao="Nova coluna 'Origem da Meta 2027' (AT) na Base_Potencial, documentando a fonte de cada meta"))

    # ---- séries / metadados existentes ----
    def read_series(r):
        return {y: ws.cell(r, cols[y]).value for y in YEARS}

    for ind_id, spec in series.items():
        r = find_row_by_id(ws, ind_id)
        before = read_series(r)
        if spec.get("values"):
            for y, v in spec["values"].items():
                ws.cell(r, cols[y]).value = v
            changed = [f"{y}: {before[y]}→{v}" for y, v in spec["values"].items() if (before[y] is None) != (v is None) or (v is not None and before[y] is not None and abs(float(before[y]) - float(v)) > 1e-9)]
        else:
            changed = []
        if spec.get("origem"):
            ws.cell(r, cols["Origem"]).value = spec["origem"]
        if spec.get("fonte"):
            ws.cell(r, cols["Fonte"]).value = spec["fonte"]
        if spec.get("codigo"):
            ws.cell(r, cols["Código/API"]).value = spec["codigo"]
        if spec.get("nacional") is not None:
            ws.cell(r, cols["Fonte nacional"]).value = spec["nacional"]
        if spec.get("nota_add"):
            old = ws.cell(r, cols["Nota"]).value or ""
            ws.cell(r, cols["Nota"]).value = (old + " | " if old else "") + spec["nota_add"]
        log.append(dict(tipo="série", id=ind_id, descricao=(spec.get("nota_add") or "") + (" Alterações: " + "; ".join(changed) if changed else " (sem alteração de valores)")))

    for ind_id, spec in rename.items():
        r = find_row_by_id(ws, ind_id)
        for k, colname in (("indicador", "Indicador"), ("codigo", "Código/API"), ("unidade", "Unidade"), ("fonte", "Fonte"), ("origem", "Origem")):
            if spec.get(k):
                ws.cell(r, cols[colname]).value = spec[k]
        if spec.get("values"):
            for y, v in spec["values"].items():
                ws.cell(r, cols[y]).value = v
        if spec.get("minimo") is not None: ws.cell(r, cols["Mínimo"]).value = spec["minimo"]
        if spec.get("maximo") is not None: ws.cell(r, cols["Máximo"]).value = spec["maximo"]
        if spec.get("nota_add"):
            old = ws.cell(r, cols["Nota"]).value or ""
            ws.cell(r, cols["Nota"]).value = (old + " | " if old else "") + spec["nota_add"]
        log.append(dict(tipo="metadados", id=ind_id, descricao=spec.get("nota_add", "metadados actualizados")))

    # ---- metas ----
    for ind_id, (meta, origem) in metas.items():
        try:
            r = find_row_by_id(ws, ind_id)
        except KeyError:
            continue
        before = ws.cell(r, cols["Meta 2027"]).value
        ws.cell(r, cols["Meta 2027"]).value = meta
        ws.cell(r, meta_col).value = origem
        if (before != meta) and not (before is None and meta is None):
            log.append(dict(tipo="meta", id=ind_id, descricao=f"Meta 2027: {before} → {meta}. {origem}"))
    # metas não alteradas: documentar origem genérica
    for r in range(FIRST, old_last + 1):
        if ws.cell(r, meta_col).value is None:
            m = ws.cell(r, cols["Meta 2027"]).value
            ws.cell(r, meta_col).value = ("Meta operacional da v7 mantida (sem meta comparável no PDN 2023-2027)" if m is not None else "Sem meta definida")

    # ---- novas linhas ----
    template_row = old_last
    for i, spec in enumerate(rows_new):
        rr = old_last + 1 + i
        data = {"A": spec["id"], "B": spec["dim"], "C": spec["subtema"], "D": spec["indicador"], "E": spec["unidade"], "F": spec["sentido"],
                "G": spec["minimo"], "H": spec["maximo"], "I": spec.get("meta"), "J": spec.get("peso", 1), "K": spec["fonte"], "L": spec.get("codigo"),
                "M": spec["origem"], "N": spec.get("nacional", 0), "O": spec["grupo"], "P": spec["prioridade"], "AP": "Auto", "AS": spec.get("nota", "")}
        append_candidate_row(ws, template_row, rr, data)
        for y, v in spec["values"].items():
            ws.cell(rr, cols[y]).value = v
        ws.cell(rr, meta_col).value = spec.get("meta_origem", "")
        log.append(dict(tipo="novo indicador", id=spec["id"], descricao=f"{spec['indicador']} — {spec['origem']}. {spec.get('nota','')}"))

    # ---- 10_Fontes: novas linhas ----
    wf = wb["10_Fontes"]
    lf = wf.max_row
    while lf > 4 and wf.cell(lf, 1).value in (None, ""):
        lf -= 1
    for i, spec in enumerate(rows_new):
        rr = lf + 1 + i
        for col in range(1, 9):
            copy_style(wf.cell(lf, col), wf.cell(rr, col))
        wf.cell(rr, 1, spec["id"]); wf.cell(rr, 2, spec["indicador"]); wf.cell(rr, 3, spec["dim"]); wf.cell(rr, 4, spec["fonte"])
        wf.cell(rr, 5, spec.get("codigo")); wf.cell(rr, 6, spec.get("nacional", 0)); wf.cell(rr, 7, spec.get("url", "")); wf.cell(rr, 8, spec["origem"])

    # ---- textos de versão ----
    ws0 = wb["00_Painel"]
    ws0["B2"].value = f"ÍNDICE GLOBAL DE DESENVOLVIMENTO DE ANGOLA (IGDA-BDA) — Selecção Final · {version_label}"
    ws2 = wb[BASE]
    ws2["B3"].value = (f"{new_last - FIRST + 1} indicadores candidatos, valores anuais como extraídos das fontes de origem (INE, BNA, MINFIN/FMI, IIMS, Banco Mundial/WDI, "
                       "OIT, OMS, UNCTAD, benchmarks internacionais). Colunas de avaliação e selecção calculadas por fórmulas. Editável: Peso (J), Meta 2027 (I) e Ajuste manual (AP). "
                       f"A coluna 'Origem da Meta 2027' documenta a fonte de cada meta ({version_label}).")
    return wb, log, new_last, meta_col


if __name__ == "__main__":
    wb, log, new_last, meta_col = build("/tmp/test_v8.xlsx")
    wb.save(os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_v8.xlsx"))
    print(json.dumps(log, ensure_ascii=False, indent=1)[:3000])
    print("new_last", new_last)
