# Scripts da revisão v8 do IGDA-BDA

Scripts Python (openpyxl, python-docx, python-pptx) que constroem o `IGDA_BDA_Construtor_v8.xlsx` a partir do v7 e geram os três suportes (Nota Metodológica v8, Justificação de mínimos e máximos v8, apresentação CEX v10). Os caminhos são relativos à raiz do repositório (a pasta `scripts_v8` está três níveis abaixo).

| Ficheiro | Função |
|---|---|
| `igda_engine.py` | Motor de cálculo que replica a cadeia de fórmulas do construtor (scoring, elegibilidade, selecção, imputação, normalização, agregação, escala comum). Reproduz exactamente a v6 e a v7. |
| `xlsx_tools.py` | Utilitários openpyxl: alargar intervalos `$5:$335 → $5:$342`, copiar linhas-modelo com tradução de fórmulas, estilos. |
| `v8_plan.py` | Plano de alterações, fonte única do registo: metas 2027 por categoria (PDN directa, transposta, convertida, MINPLAN, operacional, sem meta), regra de transposição (`TRANSPOR`), conversão WGI (`WGI_PDN`), séries actualizadas (`SERIES`), renomeações (`RENAME`), novas linhas (`NEW_ROWS`), uniformização de grafia, nomes curtos e contexto para os documentos. |
| `build_v8.py` | Aplica o plano ao livro v7: séries, metas (incluindo o cálculo das metas transpostas), novas linhas, coluna "Origem da Meta 2027", sincronização de 10_Fontes, filtros/formatação condicional/validação alargados, grafia, textos de 09_Metodologia, navegação do painel, título do gráfico do IGDA. |
| `build_v8_sheets.py` | Folhas novas `12_Emprego_INE` (blocos A–D) e `13_Alteracoes_v8` (comparação v7/v8, registo linha a linha, metas alteradas por categoria, conversão WGI), nota de versão em 09_Metodologia, regeneração das caches dos gráficos, recálculo integral ao abrir. |
| `run_build.py` | Orquestra a construção: `python3 run_build.py <saida.xlsx>` grava o livro e `<saida>_results.json` (log, metas alteradas, categorias, transposições, resultados v7/v8, diferenças do bloco D de 12_Emprego_INE). |
| `make_analysis.py` | `python3 make_analysis.py <v8.xlsx> <v7.xlsx> <analysis.json>`: drivers por indicador (Δ 2015–25 e 2023–25), cobertura, subíndices, escala comum a partir dos valores exactos, metas já cumpridas e sensibilidade da selecção ao marcador de fonte nacional. |
| `gen_nota.py` | `python3 gen_nota.py <saida.docx> [<v8.xlsx>] [<analysis.json>] [<v8_results.json>]` — Nota Metodológica v8 a partir do modelo v7. |
| `gen_justificacao.py` | `python3 gen_justificacao.py <v8.xlsx> <saida.docx> [<v8_results.json>]` — Justificação de mínimos e máximos v8. |
| `gen_cex.py` | `python3 gen_cex.py <saida.pptx> [<analysis.json>] [<v8_results.json>]` — apresentação CEX v10 a partir da v9. |
| `docx_tools.py`, `pptx_tools.py` | Utilitários de geração Word (estilos do modelo v7, arredondamento "half away from zero") e PowerPoint (substituição de texto preservando formatação, dados de gráficos por cache XML). |
| `labour_data.json` | Séries do IEA (13.ª CIET, nova metodologia, ILOSTAT nacional) extraídas dos ficheiros INE do repositório e da recolha web. |
| `analysis.json`, `v8_results.json` | Resultados intermédios usados pelos geradores de documentos. |
| `sweep_partial.json` | 242 registos da recolha web (INE, MINFIN/FMI, MINPLAN/INSS/Kwenda) com citações e URLs. |
| `verify_round1.json` | 46 achados da verificação adversarial da primeira entrega v8 (42 confirmados; 4 refutados), com os veredictos. |
| `verify_round2.json` | 28 achados da segunda ronda de verificação, sobre os ficheiros corrigidos, com os veredictos de refutação; todos tratados na versão final. |

Sequência completa (a partir da raiz do repositório):

```
S="Versão Final/Finalíssima/scripts_v8"; F="Versão Final/Finalíssima/files"
python3 "$S/run_build.py" "$F/IGDA_BDA_Construtor_v8.xlsx" && mv "$F/IGDA_BDA_Construtor_v8_results.json" "$S/v8_results.json"
python3 "$S/make_analysis.py" "$F/IGDA_BDA_Construtor_v8.xlsx" "$F/IGDA_BDA_Construtor_v7.xlsx" "$S/analysis.json"
python3 "$S/gen_nota.py" "$F/IGDA_BDA_Nota_Metodologica_v8.docx"
python3 "$S/gen_justificacao.py" "$F/IGDA_BDA_Construtor_v8.xlsx" "$F/IGDA_BDA_Justificacao_Minimos_Maximos_v8.docx"
python3 "$S/gen_cex.py" "$F/IGDA_Evolucao_CEX_v10.pptx"
```

Nota: o livro v8 é gravado sem valores em cache nas células (o Excel recalcula tudo ao abrir; se necessário, Ctrl+Alt+F9). Os valores apresentados nos documentos são os do motor de cálculo, que replica as fórmulas.
