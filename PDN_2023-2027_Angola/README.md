# PDN Angola 2023–2027

Pasta de trabalho para analisar o PDN 2023–2027 e compilar a execução publicada pelo MINPLAN.

## Entregas

- `01_Execucao_PDN_2025.xlsx` — indicadores quantitativos publicados, cobertura dos sete domínios, acções narrativas de 2025, estrutura programática e resumo de execução.
- `02_Execucao_vs_Metas_PDN_2027.xlsx` — correspondência entre os indicadores de execução e as metas do PDN, sem inventar metas numéricas inexistentes.
- `ANALISE_PDN.md` — leitura executiva, cobertura, limitações e regras de comparação.
- `fontes/` — cópias locais do PDF e do HTML usados na extração.
- `dados/` — CSVs normalizados e texto paginado do PDF.
- `scripts/compilar_pdn.py` — scraper/compilador reexecutável.

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

## Fontes

- MINPLAN: <https://www.minplan.gov.ao/en/publicacoes/relatorios-balanco-pdn>
- URL do PDN indicada: <https://www.nepad.org/sites/default/files/2024-07/20231030%283%29_layout_Final_Angola_PDN%202023-2027-1.pdf>
- Espelho acessível usado quando a URL NEPAD devolve 403: <https://www.mpla.ao/wp-content/uploads/2023/12/PDN_Angola_2023-2027.pdf>

## Limitação de cobertura

No HTML capturado, apenas o separador **Saúde** contém uma tabela de indicadores (7 linhas). Os outros seis separadores existem, mas estão vazios. A pasta regista essa situação como “Sem dados publicados no HTML”; não a trata como execução zero.
