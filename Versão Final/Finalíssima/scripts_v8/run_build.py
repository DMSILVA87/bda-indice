"""
Constrói o IGDA_BDA_Construtor_v8.xlsx completo e grava v8_results.json (log, metas alteradas, categorias, resultados v7/v8).
Uso: python3 run_build.py <saida.xlsx>
"""
import sys, json, os
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_v8, build_v8_sheets, v8_plan, igda_engine
from igda_engine import Model, load_inputs_from_workbook, DIM_ORDER, DIM_SHORT

REPO = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
V7 = os.path.join(REPO, "Versão Final/Finalíssima/files/IGDA_BDA_Construtor_v7.xlsx")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "test_v8.xlsx")

res = build_v8.build(OUT)
wb = res["wb"]
labour = json.load(open(os.path.join(HERE, "labour_data.json")))
web_new = {"2026T1": {"desemprego": 21.3, "emprego": 42.3, "desemprego_jovem": 40.7, "informal": 79.3, "actividade": 53.8},
           "2026T2": {"desemprego": 21.5, "emprego": 44.2, "desemprego_jovem": 40.5, "informal": 80.1, "actividade": 56.3}}
# resultados: v7 (recálculo integral) e v8 (motor sobre o livro novo, gravado primeiro num temporário)
tmp = OUT.replace(".xlsx", "_pre.xlsx"); wb.save(tmp)
m7 = Model(load_inputs_from_workbook(V7)).run()
m8 = Model(load_inputs_from_workbook(tmp)).run()


def pack(m):
    dims = {s: m.subindices[d] for d, s in zip(DIM_ORDER, DIM_SHORT) if any(v is not None for v in m.subindices[d])}
    return {"dims": dims, "igda": m.igda}


n_set, n_clear = build_v8_sheets.refresh_chart_caches(wb, m8)
res["log"].append(dict(tipo="estrutura", id="", descricao=f"Caches dos gráficos (00_Painel, 07_Subindices, 08_Visualizacoes) regeneradas com os resultados v8 ({n_set} séries; {n_clear} sem intervalo reconhecido, cache removida)"))
ws12, emprego_diffs = build_v8_sheets.add_emprego_ine(wb, labour, web_new, m8)
build_v8_sheets.add_changelog(wb, res["log"], pack(m7), pack(m8), res["meta_changes"], res["transposicoes"])
build_v8_sheets.add_version_notes(wb, res["n_cand"])
wb.save(OUT)
os.remove(tmp)
out = {"log": res["log"], "meta_changes": res["meta_changes"], "meta_cat": res["meta_cat"], "transposicoes": res["transposicoes"], "n_cand": res["n_cand"],
       "v7": pack(m7), "v8": pack(m8), "selected_v8": [r.id for r in m8.selected()], "selected_v7": [r.id for r in m7.selected()],
       "emprego_diffs": emprego_diffs, "chart_caches": [n_set, n_clear],
       "wgi": {k: dict(nome=v[0], p2022=v[1], p2027=v[2], e2022=v[3], meta=v8_plan.METAS[k][0], sens_pp=v8_plan.wgi_meta_pp(v[3], v[1], v[2])) for k, v in v8_plan.WGI_PDN.items()}}
json.dump(out, open(OUT.replace(".xlsx", "_results.json"), "w"), ensure_ascii=False, indent=1, default=str)
print("saved", OUT, "| rows", len(m8.rows), "| selected", len(m8.selected()), "| same selection as v7:", out["selected_v8"] == out["selected_v7"], "| sheets", wb.sheetnames)
print(m8.summary())
print("metas alteradas:", len(res["meta_changes"]), "| charts:", n_set, n_clear)
