"""
Utilitários para editar o construtor IGDA (openpyxl) preservando fórmulas:
  - extend_ranges: alarga todas as referências "$5:$<last>" (e variantes) de 02_Base_Potencial para uma nova última linha
  - append_candidate_row: acrescenta uma linha de candidato em 02_Base_Potencial copiando/traduzindo as fórmulas
  - set_year_values: escreve a série anual de um indicador
"""
from __future__ import annotations
import re
from copy import copy

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, column_index_from_string

YEARS = list(range(2015, 2026))
BASE = "02_Base_Potencial"
HDR_ROW = 4
FIRST = 5


def sheet_columns(ws, hdr_row=HDR_ROW):
    return {c.value: c.column for c in ws[hdr_row] if c.value is not None}


def last_data_row(ws):
    r = ws.max_row
    while r > FIRST and ws.cell(r, 1).value in (None, ""):
        r -= 1
    return r


def extend_ranges(wb, old_last: int, new_last: int):
    """Substitui referências absolutas à última linha da Base_Potencial em todas as fórmulas do livro."""
    if old_last == new_last:
        return 0
    pat = re.compile(r"\$(%d)(?![0-9])" % old_last)
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("=") and ("$%d" % old_last) in v:
                    # só substituir quando a referência pertence à Base_Potencial ou a intervalos internos da própria folha Base
                    nv = pat.sub("$%d" % new_last, v)
                    if nv != v:
                        c.value = nv; n += 1
    return n


def copy_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font); dst.border = copy(src.border); dst.fill = copy(src.fill)
        dst.number_format = src.number_format; dst.protection = copy(src.protection); dst.alignment = copy(src.alignment)


def append_candidate_row(ws, template_row: int, new_row: int, data: dict):
    """Cria new_row copiando fórmulas (traduzidas) e estilos de template_row; escreve valores de `data` (col letter -> valor)."""
    for col in range(1, ws.max_column + 1):
        src = ws.cell(template_row, col)
        dst = ws.cell(new_row, col)
        copy_style(src, dst)
        v = src.value
        if isinstance(v, str) and v.startswith("="):
            dst.value = Translator(v, origin=f"{get_column_letter(col)}{template_row}").translate_formula(f"{get_column_letter(col)}{new_row}")
        else:
            dst.value = None
    for letter, val in data.items():
        ws[f"{letter}{new_row}"] = val
    ws.row_dimensions[new_row].height = ws.row_dimensions[template_row].height


def set_year_values(ws, row: int, cols: dict, series: dict):
    """series: {ano: valor|None}; escreve apenas os anos presentes em series (None apaga)."""
    for y, v in series.items():
        c = cols[y]
        ws.cell(row, c).value = v


def find_row_by_id(ws, ind_id: str):
    for r in range(FIRST, ws.max_row + 1):
        if ws.cell(r, 1).value == ind_id:
            return r
    raise KeyError(ind_id)
