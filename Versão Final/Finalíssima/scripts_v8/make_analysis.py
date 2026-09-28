"""
Gera analysis.json (drivers por indicador, cobertura, subíndices, escala comum) para os geradores de documentos.
Uso: python3 make_analysis.py <v8.xlsx> <v7.xlsx> <analysis.json>
"""
import sys, json, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from igda_engine import Model, load_inputs_from_workbook, DIM_ORDER, DIM_SHORT


def analyse(path):
    m = Model(load_inputs_from_workbook(path)).run()
    out = {"dims": {}, "indicators": [], "coverage": {}}
    for dim, short in zip(DIM_ORDER, DIM_SHORT):
        for s in range(1, 11):
            r = m.base_indice[dim][s - 1]
            if r is None:
                continue
            sc = m.scores[(dim, s)]; fl = m.flags[(dim, s)]
            nobs = sum(1 for f in fl if f == "OBS")
            out["indicators"].append(dict(dim=short, dim_full=dim, slot=s, id=r.id, nome=r.indicador, unidade=r.unidade, sentido=r.sentido, fonte=r.fonte, nacional=r.nacional,
                                          minimo=r.minimo, maximo=r.maximo, meta=r.meta, ultimo=r.ultimo, cobertura=nobs / 11, score_2015=sc[0], score_2023=sc[8], score_2025=sc[10],
                                          d_15_25=(sc[10] - sc[0]) if sc[0] is not None and sc[10] is not None else None,
                                          d_23_25=(sc[10] - sc[8]) if sc[8] is not None and sc[10] is not None else None,
                                          flags=fl, values=r.values, adj=m.adjusted[(dim, s)]))
    out["subindices"] = {short: m.subindices[dim] for dim, short in zip(DIM_ORDER, DIM_SHORT)}
    out["igda"] = m.igda
    out["escala"] = {(DIM_SHORT[DIM_ORDER.index(k)] if k in DIM_ORDER else k): v for k, v in m.escala.items()}
    sel = m.selected()
    out["coverage"] = {"n_sel": len(sel), "last2025": sum(1 for r in sel if r.ultimo == 2025), "last2024": sum(1 for r in sel if r.ultimo == 2024),
                       "last2023": sum(1 for r in sel if r.ultimo == 2023), "nacional": sum(1 for r in sel if r.nacional == 1), "n_rows": len(m.rows),
                       "estimated_cells": sum(1 for k, f in m.flags.items() for x in f if x.startswith("EST")), "total_cells": sum(len(f) for f in m.flags.values())}
    out["elegiveis_por_dim"] = {short: sum(1 for r in m.rows if r.dim == dim and r.elegivel == 1) for dim, short in zip(DIM_ORDER, DIM_SHORT)}
    out["candidatos_por_dim"] = {short: sum(1 for r in m.rows if r.dim == dim) for dim, short in zip(DIM_ORDER, DIM_SHORT)}
    out["selected_ids"] = [r.id for r in sel]
    # escala comum a partir dos valores não arredondados (evita dupla arredondagem 1 decimal -> inteiro nos documentos)
    def esc_raw(vals):
        s0 = vals[0]
        return [None if (s0 in (None, 0) or v is None) else 100 * v / s0 for v in vals]
    out["escala_raw"] = {short: esc_raw(m.subindices[dim]) for dim, short in zip(DIM_ORDER, DIM_SHORT)}
    out["escala_raw"]["IGDA"] = esc_raw(m.igda)
    # metas 2027 já cumpridas (em 2015 e na última observação) entre os seleccionados
    def cumprida(r, v):
        if r.meta is None or v is None or not isinstance(r.meta, (int, float)):
            return None
        return (v >= r.meta) if r.sentido == "+" else (v <= r.meta)
    out["metas_cumpridas"] = []
    for r in sel:
        obs = [(y, v) for y, v in zip(range(2015, 2026), r.values) if v is not None]
        if not obs:
            continue
        out["metas_cumpridas"].append(dict(id=r.id, nome=r.indicador, meta=r.meta, sentido=r.sentido, v2015=obs[0][1] if obs[0][0] == 2015 else None,
                                           ultimo_ano=obs[-1][0], ultimo=obs[-1][1], em_2015=cumprida(r, obs[0][1] if obs[0][0] == 2015 else None), no_ultimo=cumprida(r, obs[-1][1])))
    return out


def sensitivity_nacional(path, ids=("MAC004", "MAC005")):
    """Recalcula o índice com o marcador de fonte nacional a 0 nos indicadores indicados (sensibilidade da selecção ao bónus nacional)."""
    inp = load_inputs_from_workbook(path)
    for r in inp.rows:
        if r.id in ids:
            r.nacional = 0
    m = Model(inp).run()
    return {"igda_2025": m.igda[-1], "igda": m.igda, "selected": [r.id for r in m.selected()],
            "macro_2025": m.subindices["2. Estabilidade Macroeconómica"][-1]}


if __name__ == "__main__":
    v8, v7, out = sys.argv[1], sys.argv[2], sys.argv[3]
    a8 = analyse(v8); a7 = analyse(v7)
    sens = sensitivity_nacional(v8)
    base_sel = set(a8["selected_ids"]); alt_sel = set(sens["selected"])
    sens["saem"] = sorted(base_sel - alt_sel); sens["entram"] = sorted(alt_sel - base_sel)
    a8["sens_nacional"] = sens
    json.dump({"v8": a8, "v7": a7}, open(out, "w"), ensure_ascii=False, indent=1, default=str)
    print("sensibilidade nac=0 em MAC004/MAC005:", round(sens["igda_2025"], 2), "saem", sens["saem"], "entram", sens["entram"])
    print("metas cumpridas no último ano:", [d["id"] for d in a8["metas_cumpridas"] if d["no_ultimo"]])
    print("metas cumpridas em 2015:", [d["id"] for d in a8["metas_cumpridas"] if d["em_2015"]])
    print("coverage v8:", a8["coverage"])
    print("igda v8:", [round(v, 2) for v in a8["igda"]])
    print("igda v7:", [round(v, 2) for v in a7["igda"]])
    print("sub 2025 v8 vs v7:", {d: (round(a8["subindices"][d][-1], 2), round(a7["subindices"][d][-1], 2)) for d in DIM_SHORT[:8]})
    print("same selection:", a8["selected_ids"] == a7["selected_ids"])
