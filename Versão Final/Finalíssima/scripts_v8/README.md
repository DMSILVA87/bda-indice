# Scripts de construção da v8 (IGDA-BDA)

Ferramentas Python usadas para gerar os entregáveis da pasta `../files/` a partir da v7. Executar a partir desta pasta
(os caminhos para o repositório estão em `build_v8.py`, `run_build.py`, `gen_*.py` e podem ser adaptados).

| Ficheiro | Função |
|---|---|
| `igda_engine.py` | Motor de cálculo que replica todas as fórmulas do construtor (selecção, imputação, normalização, agregação). Reproduz exactamente os valores da v6 e da v7; usado para validar a v8 e para gerar os números dos documentos. |
| `xlsx_tools.py` | Alargamento de intervalos (`$5:$335` → `$5:$342`), cópia/tradução de fórmulas para novas linhas. |
| `v8_plan.py` | Plano de alterações: metas 2027 (PDN), novas linhas, séries actualizadas, notas e contexto. Fonte única do registo de alterações. |
| `build_v8.py`, `build_v8_sheets.py`, `run_build.py` | Constroem `IGDA_BDA_Construtor_v8.xlsx` (folhas novas `12_Emprego_INE` e `13_Alteracoes_v8`) e gravam `v8_results.json`. |
| `gen_nota.py`, `gen_justificacao.py`, `gen_cex.py` | Geram a Nota Metodológica v8, a Justificação de mínimos/máximos v8 e a apresentação CEX v10 (a partir dos modelos v7/v9, mantendo estilos). |
| `docx_tools.py`, `pptx_tools.py` | Utilitários python-docx / python-pptx. |
| `labour_data.json` | Séries INE/IEA (13.ª CIET, nova metodologia) e séries OIT harmonizadas usadas na folha 12. |
| `analysis.json`, `v8_results.json` | Resultados calculados (subíndices, IGDA, drivers, cobertura; registo de alterações). |
| `sweep_partial.json` | Dados recolhidos na web (INE, MINFIN/FMI, MINPLAN/INSS/Kwenda), com citações e URLs. |

Dependências: `openpyxl`, `python-docx`, `python-pptx`, `pandas`, `lxml`.

Ordem de execução: `python3 run_build.py <saída.xlsx>` → `python3 gen_justificacao.py <v8.xlsx> <saída.docx>` → `python3 gen_nota.py <saída.docx>` → `python3 gen_cex.py <saída.pptx>`.
