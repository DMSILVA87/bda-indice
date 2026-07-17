# Análise do PDN 2023-2027 e do painel de execução

Data de extração: 2026-07-17

## 1. Plano de referência

O PDN 2023-2027 é o instrumento de médio prazo que implementa a Estratégia de Longo Prazo Angola 2050. A estrutura publicada organiza o plano em sete eixos, 16 políticas, 50 programas, 171 objectivos, 284 prioridades e 398 indicadores monitorizados.

A metodologia do próprio PDN separa indicadores de impacto, indicadores de resultado e execução orçamental programática. O plano prevê monitorização permanente, relatórios mensais e trimestrais, balanços anuais, avaliação intercalar em 2025 e avaliação final em 2028.

Os dois motores de desenvolvimento são o desenvolvimento do capital humano e a segurança alimentar. Os filtros transversais incluem receita fiscal, juventude, igualdade de género, emprego, sustentabilidade ambiental, comunidades vulneráveis e ambiente de negócios.

## 2. O que foi publicado no painel do MINPLAN

O painel foi capturado na página `https://www.minplan.gov.ao/en/publicacoes/relatorios-balanco-pdn` e identifica o período `Anual 2025`. O inventário verificou os sete domínios de `Implementation - Key Indicators by Domain`.

Há dados numéricos em 1 domínio(s): Saúde. A tabela contém 7 indicadores, com ano base 2022 e valores para 2023, 2024 e 2025.

Os painéis sem tabela no HTML são: Protecção Social, Comunicação, Educação e Formação Profissional, Infraestruturas, Habitação e Serviços Comunitários, Fomento à Produção, Económico. Eles foram mantidos no inventário com estado `Sem dados publicados no HTML`; não foram convertidos em zero.

A secção `Main Actions` contém resultados narrativos de 2025 para os domínios Social, Fomento da produção nacional, Infra-estruturas e Construção e Obras públicas. Esses resultados foram preservados numa folha separada, sem os misturar com indicadores quantitativos.

## 3. Comparação com metas

Foram catalogadas 115 linhas de metas das páginas `Metas da Política` do PDF. A Política de Saúde (página 74) define metas para esperança de vida, mortalidade, despesa, profissionais e densidade de unidades; as prioridades de imunização, pré-natal, malária e tuberculose aparecem nas páginas 76–77.

Os sete indicadores publicados pelo MINPLAN não têm, no quadro de metas da Política de Saúde, um valor quantitativo equivalente que permita calcular automaticamente o progresso para 2027. Por isso, o segundo Excel distingue correspondência temática de comparabilidade numérica e deixa diferença/progresso vazios quando não há meta equivalente.

## 4. Proveniência e limitação de acesso

A URL NEPAD indicada foi preservada como fonte solicitada. Como o ficheiro devolveu erro 403 no ambiente de extração, foi guardado um espelho acessível do documento no ficheiro `fontes/PDN_Angola_2023-2027.pdf`. A URL usada e o fallback estão registados em `fontes/00_manifesto_fontes.csv` e nas folhas `Manifesto_Fontes`.

A ausência de tabelas nos seis domínios não prova ausência de execução; prova apenas que esses dados não estavam presentes no HTML recebido da página consultada. Para completar esses domínios será necessário obter os relatórios/documentos sectoriais que o MINPLAN ainda assinala como em desenvolvimento.

## Ficheiros entregues

- `01_Execucao_PDN_2025.xlsx`: execução publicada, cobertura dos domínios, acções 2025, estrutura e resumo.
- `02_Execucao_vs_Metas_PDN_2027.xlsx`: comparação sem inferência indevida, catálogo completo das metas transcritas e metas de saúde em separado.
- `fontes/`: cópias locais das fontes usadas.
- `dados/`: CSVs normalizados e texto extraído do PDF.
- `scripts/compilar_pdn.py`: scraper/compilador reexecutável.
