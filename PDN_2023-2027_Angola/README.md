# PDN Angola 2023–2027

Pasta de trabalho para analisar o PDN 2023–2027 e compilar a execução publicada pelo MINPLAN.

## Entregas

- `01_Execucao_PDN_2025.xlsx` — indicadores quantitativos publicados, cobertura dos sete domínios, acções narrativas de 2025, estrutura programática e resumo de execução.
- `02_Execucao_vs_Metas_PDN_2027.xlsx` — correspondência entre os indicadores de execução e as metas do PDN, sem inventar metas numéricas inexistentes.
- `03_Series_Meta_Indices_Compositos.xlsx` — séries meta por indicador (último dado verificado → Meta 2027, interpolação linear ou geométrica consoante a especificação da série), metas anuais interpoladas, índice compósito de execução (observados 2022-2025, 2022=100, média geométrica) e índice de execução percentual (observado vs meta anual, tecto 100%). Metas 2027 derivadas são identificadas como tal na coluna de origem.
- `ANALISE_PDN.md` — leitura executiva, cobertura, limitações e regras de comparação.
- `fontes/` — cópias locais do PDF, do HTML e do bundle JavaScript lazy usados na extração.
- `dados/` — CSVs normalizados e texto paginado do PDF.
- `scripts/compilar_pdn.py` — scraper/compilador reexecutável.
- `scripts/construir_series_meta_indices.py` — gera o `03_...xlsx` e os CSVs `series_meta_convergencia.csv`, `metas_anuais_interpoladas.csv`, `indice_composito_execucao.csv`, `indice_execucao_percentual.csv` e `metas_pdn_trajectorias_2022_2027.csv`.

## Execução

Na raiz do repositório:

```bash
python3 -m pip install -r PDN_2023-2027_Angola/requirements.txt
python3 PDN_2023-2027_Angola/scripts/compilar_pdn.py
```

Para renovar as fontes:

```bash
python3 PDN_2023-2027_Angola/scripts/compilar_pdn.py --refresh
```

Para regenerar as séries meta e os índices compósitos:

```bash
python3 PDN_2023-2027_Angola/scripts/construir_series_meta_indices.py
```

## Fontes

- MINPLAN: <https://www.minplan.gov.ao/en/publicacoes/relatorios-balanco-pdn>
- URL do PDN indicada: <https://www.nepad.org/sites/default/files/2024-07/20231030%283%29_layout_Final_Angola_PDN%202023-2027-1.pdf>
- Espelho acessível usado quando a URL NEPAD devolve 403: <https://www.mpla.ao/wp-content/uploads/2023/12/PDN_Angola_2023-2027.pdf>

## Cobertura

Os sete domínios do componente **Implementation - Key Indicators by Domain** são extraídos do bundle JavaScript oficial associado à página. O HTML inicial só materializa o separador **Saúde**; os restantes são lazy-loaded. A extração completa contém 51 indicadores e preserva a meta anual 2025 publicada no bundle.
