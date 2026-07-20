# Análise do PDN 2023-2027 e do painel de execução

Data de extração: 2026-07-17

## 1. Plano de referência

O PDN 2023-2027 é o instrumento de médio prazo que implementa a Estratégia de Longo Prazo Angola 2050. A estrutura publicada organiza o plano em sete eixos, 16 políticas, 50 programas, 171 objectivos, 284 prioridades e 398 indicadores monitorizados.

A metodologia do próprio PDN separa indicadores de impacto, indicadores de resultado e execução orçamental programática. O plano prevê monitorização permanente, relatórios mensais e trimestrais, balanços anuais, avaliação intercalar em 2025 e avaliação final em 2028.

Os dois motores de desenvolvimento são o desenvolvimento do capital humano e a segurança alimentar. Os filtros transversais incluem receita fiscal, juventude, igualdade de género, emprego, sustentabilidade ambiental, comunidades vulneráveis e ambiente de negócios.

## 2. O que foi publicado no painel do MINPLAN

O painel foi capturado na página `https://www.minplan.gov.ao/en/publicacoes/relatorios-balanco-pdn` e identifica o período `Anual 2025`. O inventário verificou os sete domínios de `Implementation - Key Indicators by Domain`.

Há dados numéricos em 7 domínio(s): Saúde, Protecção Social, Comunicação, Educação e Formação Profissional, Infraestruturas, Habitação e Serviços Comunitários, Fomento à Produção, Económico. O bundle lazy contém 51 indicadores, com ano base 2022, valores para 2023, 2024, meta anual 2025 e execução 2025.

Todos os sete domínios têm dados no bundle JavaScript, embora os seis separadores inactivos não apareçam no HTML inicial por serem lazy-loaded.

A secção `Main Actions` contém resultados narrativos de 2025 para os domínios Social, Fomento da produção nacional, Infra-estruturas e Construção e Obras públicas. Esses resultados foram preservados numa folha separada, sem os misturar com indicadores quantitativos.

## 3. Comparação com metas

Foram catalogadas 118 linhas de metas das páginas `Metas da Política` do PDF. A Política de Saúde (página 74) define metas para esperança de vida, mortalidade, despesa, profissionais e densidade de unidades; as prioridades de imunização, pré-natal, malária e tuberculose aparecem nas páginas 76–77.

O segundo Excel preserva a meta anual 2025 publicada pelo MINPLAN e procura equivalências com as metas 2027 do PDN. Quando o conceito, escala ou período não é comparável, a referência é mantida sem cálculo automático; quando é comparável, a diferença e a percentagem do alvo são calculadas.

## 4. Proveniência e limitação de acesso

A URL NEPAD indicada foi preservada como fonte solicitada. Como o ficheiro devolveu erro 403 no ambiente de extração, foi guardado um espelho acessível do documento no ficheiro `fontes/PDN_Angola_2023-2027.pdf`. A URL usada e o fallback estão registados em `fontes/00_manifesto_fontes.csv` e nas folhas `Manifesto_Fontes`.

Os seis domínios inactivos não tinham dados no HTML inicial porque os componentes são lazy-loaded. A extração usa o bundle JavaScript oficial associado à página; se o site alterar o hash do bundle, o manifesto e o script devem ser actualizados.

## 5. Séries meta e índices compósitos

O ficheiro `03_Series_Meta_Indices_Compositos.xlsx` acrescenta três camadas analíticas sobre os 51 indicadores de execução (46 com meta; excluídos: novas camas, coberturas de água urbana/rural e empregos formais bruto/líquido, sem meta publicada):

1. **Séries Meta** — para cada indicador, uma trajectória do último dado verificado (2025; 2023 nos três indicadores de professores) até à Meta 2027. A Meta 2027 é a publicada no PDN quando existe correspondência directa (7 casos); nos restantes é derivada da trajectória planeada 2022→Meta MINPLAN 2025 (prolongamento ao mesmo ritmo, ou manutenção da meta 2025 quando o plano não previa melhoria), sempre identificada como derivada. Interpolação linear para taxas, percentagens e stocks cumulativos; geométrica (CAGR) para níveis de produção.
2. **Índice compósito de execução** (observados 2022-2025): normalização a 2022=100 com inversão dos indicadores "menor é melhor" (malária, inflação, desemprego) e média geométrica por domínio e entre os sete domínios. Resultado global: 100 → 104,0 (2023) → 123,4 (2024) → 131,6 (2025). Protecção Social é o domínio com maior progressão (219,4) e Infraestruturas o mais estagnado (101,8).
3. **Índice de execução percentual** (observado vs meta anual interpolada, tecto de 100%): global de 87,1% (2023), 87,4% (2024) e 86,7% (2025) — o ritmo de execução mantém-se ~13 p.p. aquém da trajectória de metas. A coluna de validação face à meta MINPLAN 2025 publicada reproduz exactamente o "% do alvo" do painel oficial nos indicadores "maior é melhor"; nos "menor é melhor" o atingimento é invertido intencionalmente (ex.: desemprego 92,9% em vez dos 107,6% do painel).

A folha `05_Metas_PDN_Trajectorias` interpola ainda 108 das 118 metas 2027 do PDN a partir da base 2022 (10 sem base ou sem meta numérica).

## Ficheiros entregues

- `01_Execucao_PDN_2025.xlsx`: execução publicada, cobertura dos domínios, acções 2025, estrutura e resumo.
- `02_Execucao_vs_Metas_PDN_2027.xlsx`: comparação sem inferência indevida, catálogo completo das metas transcritas e metas de saúde em separado.
- `fontes/`: cópias locais das fontes usadas.
- `dados/`: CSVs normalizados e texto extraído do PDF.
- `scripts/compilar_pdn.py`: scraper/compilador reexecutável.
