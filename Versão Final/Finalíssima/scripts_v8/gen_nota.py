"""
Gera IGDA_BDA_Nota_Metodologica_v8.docx (a partir do modelo v7, mesmos estilos) com os resultados recalculados da v8.
"""
from __future__ import annotations
import os
import json, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx.shared import Cm
from docx_tools import DocBuilder, fmt
import v8_plan as P

SC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "")
A = json.load(open(SC + "analysis.json"))
R = json.load(open(SC + "v8_results.json"))
V8, V7 = A["v8"], A["v7"]
YEARS = list(range(2015, 2026))
DIMS8 = ["Governança", "Macroeconomia", "Capital Humano", "Inclusão Social", "Infraestruturas", "Mercado Trabalho", "Saúde/Alimentar", "Diversificação"]
DIM_LABEL = {"Governança": "Governança e Estado de Direito", "Macroeconomia": "Estabilidade Macroeconómica", "Capital Humano": "Capital Humano", "Inclusão Social": "Inclusão Social e Protecção",
             "Infraestruturas": "Infraestruturas e Serviços", "Mercado Trabalho": "Mercado de Trabalho", "Saúde/Alimentar": "Segurança Alimentar e Saúde", "Diversificação": "Diversificação Produtiva e Setor Privado"}


def f1(x, sign=False):
    return fmt(x, 1, sign=sign)


def pct1(x):
    return fmt(x, 1, pct=True)


def build(template, out):
    igda = V8["igda"]; igda7 = V7["igda"]
    sub = V8["subindices"]; sub7 = V7["subindices"]
    esc = V8["escala"]
    cov = V8["coverage"]
    n_rows = cov["n_rows"]
    ind = V8["indicators"]
    b = DocBuilder(template)

    b.title("Índice Global de Desenvolvimento de Angola — IGDA-BDA")
    b.subtitle(f"Nota Metodológica (v8) · Selecção Final · 2015–2025 · 8 dimensões activas (de 11) · {cov['n_sel']} indicadores")
    b.author("Banco de Desenvolvimento de Angola · Gabinete de Planeamento e Controlo")
    b.date("Setembro de 2026")

    # 1
    b.h1("1. Introdução e objectivo")
    b.h2("1.1 Propósito do índice")
    b.para("O Índice Global de Desenvolvimento de Angola (IGDA-BDA) é uma medida compósita que sintetiza, num único valor anual numa escala de 0 a 100, o estado multidimensional do desenvolvimento do país. Foi concebido para apoiar a função de estudos e planeamento do Banco de Desenvolvimento de Angola, oferecendo um instrumento de leitura rápida da trajectória de desenvolvimento, de comparação entre dimensões e de fundamentação de decisões de afectação de recursos e de avaliação de políticas.")
    b.para("O índice não substitui a análise sectorial detalhada; o seu valor está em integrar, de forma transparente e replicável, informação dispersa por dezenas de fontes e indicadores num quadro coerente, permitindo distinguir onde o progresso é robusto e onde persistem défices estruturais.")
    b.h2("1.2 Âmbito e cobertura")
    b.para(f"A unidade de análise é Angola no seu conjunto. A cobertura temporal abrange o período de 2015 a 2025, com periodicidade anual. O ano de 2025 constitui o ano de referência (headline), embora deva ser lido como provisório pelas razões expostas na Secção 9. O índice é composto por oito dimensões activas — de um total de onze definidas — e por {cov['n_sel']} indicadores, resultantes de uma selecção automática por ranking dentro de cada dimensão, complementada por uma curadoria de não-redundância. A versão v8 actualiza os dados até Setembro de 2026, alinha as metas 2027 com o Plano de Desenvolvimento Nacional 2023-2027 (PDN) e incorpora o feedback externo recebido sobre a v7 (Secção 8.6).")
    b.h2("1.3 Enquadramento metodológico")
    b.para("A construção segue as melhores práticas internacionais para índices compósitos, em particular o quadro de dez passos do Handbook on Constructing Composite Indicators da OCDE e do Joint Research Centre da Comissão Europeia (OCDE/JRC, 2008). Esse quadro estrutura o processo desde a definição de um enquadramento teórico, passando pela selecção e tratamento de dados, normalização, ponderação e agregação, até à análise de robustez e à comunicação dos resultados.")
    b.para("Os dez passos são: (1) quadro teórico; (2) selecção de dados; (3) tratamento de dados em falta; (4) análise multivariada; (5) normalização; (6) ponderação e agregação; (7) análise de incerteza e sensibilidade; (8) regresso aos dados (desconstrução); (9) ligação a outros indicadores; e (10) visualização e comunicação. As secções seguintes mapeiam directamente estes passos; o passo 9 é desenvolvido na Secção 10.5, com o cruzamento sistemático entre o índice e as metas do PDN.")
    b.h2("1.4 Organização do documento")
    b.para("A Secção 2 estabelece o quadro conceptual e a arquitectura do índice. A Secção 3 descreve a base de dados, as fontes e os vintages. A Secção 4 detalha o processo de selecção de indicadores. As Secções 5 a 7 cobrem, respectivamente, o tratamento de dados em falta, a normalização e a agregação. A Secção 8 descreve a configuração final do índice e a revisão v7 → v8. As Secções 9 e 10 tratam dos pressupostos e limitações e da interpretação dos resultados, incluindo o cruzamento com o PDN. A Secção 11 lista recomendações de desenvolvimento futuro e a Secção 12 descreve a implementação técnica. Os Anexos apresentam a estrutura completa dos indicadores, as fontes, siglas e vintages, e o registo de alterações.")

    # 2
    b.h1("2. Quadro conceptual")
    b.h2("2.1 Conceito de desenvolvimento adoptado")
    b.para("O IGDA-BDA adopta uma concepção multidimensional de desenvolvimento, alinhada com a tradição do desenvolvimento humano e com a agenda dos Objectivos de Desenvolvimento Sustentável. O desenvolvimento não é reduzido ao rendimento por habitante; é entendido como a expansão simultânea das capacidades económicas, sociais, institucionais e ambientais que sustentam o bem-estar da população de forma duradoura. Esta opção determina a escolha das dimensões e a recusa de qualquer indicador único como medida-resumo.")
    b.h2("2.2 Arquitectura hierárquica")
    b.para("O índice está organizado em três níveis. No nível inferior estão os indicadores individuais, cada um proveniente de uma fonte e expresso na sua unidade própria. No nível intermédio, os indicadores agrupam-se em dimensões temáticas, das quais resulta um subíndice por dimensão. No nível superior, os subíndices dimensionais combinam-se no índice global, o IGDA. Esta estrutura espelha a do Índice de Desenvolvimento Humano e tem a vantagem de tornar legível o contributo de cada dimensão para o resultado agregado.")
    b.h2("2.3 As dimensões do índice")
    b.para("A versão actual define onze dimensões. Estão activas oito delas, que reúnem os quarenta e um indicadores seleccionados; as três restantes — Ambiente, Clima e Resiliência; Demografia, Território e Urbanização; e Transformação Digital e Inovação — estão definidas mas inactivas, por não terem, nesta versão, indicadores seleccionados que cumpram os critérios. As dimensões são:")
    b.bullets([
        ("Governança e Estado de Direito (activa) — ", "percepção da corrupção (Transparency International) e quatro indicadores de governança do Banco Mundial (estabilidade política, Estado de direito, eficiência do governo e qualidade regulatória)."),
        ("Estabilidade Macroeconómica (activa) — ", "crescimento do PIB total e não petrolífero, rendimento por habitante, inflação, saldo orçamental, dívida pública e reservas internacionais."),
        ("Capital Humano (activa) — ", "esperança de vida, mortalidade infantil, escolaridade obrigatória e despesa pública em educação."),
        ("Inclusão Social e Protecção (activa) — ", "rendimento nacional bruto por habitante, disparidade de género (Global Gender Gap) e representação das mulheres no parlamento."),
        ("Infraestruturas e Serviços (activa) — ", "acesso a água potável, densidade rodoviária, electrificação, energia limpa para cozinhar e tráfego portuário de contentores."),
        ("Mercado de Trabalho (activa) — ", "emprego, desemprego total e juvenil, participação feminina na força de trabalho e produtividade do trabalho."),
        ("Segurança Alimentar e Saúde (activa) — ", "mortalidade materna e de menores de cinco anos, cobertura vacinal, desnutrição, incidência de malária, despesa em saúde e despesa directa das famílias."),
        ("Diversificação Produtiva e Setor Privado (activa) — ", "o crescimento do PIB não petrolífero, o índice de capacidades produtivas da UNCTAD, a indústria transformadora no PIB, o peso dos combustíveis nas exportações e o crédito ao sector privado."),
        ("Ambiente, Clima e Resiliência (inactiva) — ", "dimensão definida mas sem indicadores seleccionados nesta entrega."),
        ("Demografia, Território e Urbanização (inactiva) — ", "dimensão definida mas sem indicadores seleccionados nesta entrega."),
        ("Transformação Digital e Inovação (inactiva) — ", "dimensão definida mas sem indicadores seleccionados nesta entrega."),
    ])
    b.para("A composição de cada dimensão activa, indicador a indicador, consta do Anexo A. A Secção 8 descreve a configuração final do índice.")
    b.h2("2.4 Princípios de desenho")
    b.para("Três princípios orientaram o desenho e a calibração do índice:")
    b.bullets([
        ("Parcimónia. ", "Cada dimensão deve conter o número mínimo de indicadores necessário para representar o conceito, evitando inflar a estrutura com variáveis marginais que apenas adicionam ruído."),
        ("Não-redundância. ", "Conceitos sobrepostos ou indicadores fortemente correlacionados não devem entrar em duplicado, sob pena de sobreponderar implicitamente esses conceitos. A selecção final respeita este princípio (Secção 8)."),
        ("Transparência. ", "As regras de selecção, normalização e agregação devem ser explícitas, reproduzíveis e auditáveis, privilegiando opções simples e defensáveis em detrimento de esquemas estatísticos opacos."),
    ])

    # 3
    b.h1("3. Base de dados e fontes")
    b.h2("3.1 Catálogo de indicadores potenciais")
    b.para(f"A construção parte de um catálogo abrangente de {n_rows} indicadores candidatos (331 na v7, mais sete acrescentados na v8), organizados por dimensão e mantido na base potencial do construtor. Para cada candidato, o catálogo regista a unidade de medida, o sentido (se valores mais altos representam melhor ou pior desempenho), as fronteiras de normalização, a meta 2027 e — novidade da v8 — a origem documentada dessa meta, a fonte, um marcador de fonte nacional, o grupo temático para controlo de duplicação, a série anual de 2015 a 2025, a cobertura observada e o último ano com dados. É a partir deste catálogo que o motor de selecção escolhe, de forma automática e recalculável, os indicadores que entram no índice.")
    b.h2("3.2 Fontes")
    b.para("Os indicadores provêm de uma combinação de fontes nacionais e internacionais. Entre as fontes nacionais contam-se o Instituto Nacional de Estatística (Contas Nacionais anuais e trimestrais, Índice de Preços no Consumidor Nacional, Inquérito sobre o Emprego em Angola, IIMS 2023-24, Censo 2024), o Banco Nacional de Angola, o Ministério das Finanças, o Ministério do Planeamento (Balanço do PDN), o Ministério da Saúde, o Ministério da Energia e Águas, o Ministério dos Transportes e o INSS/MAPTSS. Entre as fontes internacionais figuram o Banco Mundial (World Development Indicators e Worldwide Governance Indicators), o FMI (World Economic Outlook), a Organização Mundial da Saúde, a UNESCO, a OIT (ILOSTAT), a UNCTAD, a Transparency International, a União Interparlamentar, o Fórum Económico Mundial, a UNICEF e a FAO. A lista completa consta do Anexo B.")
    b.para(f"Dos {cov['n_sel']} indicadores seleccionados, {cov['nacional']} envolvem uma fonte nacional. O processo de selecção atribui uma bonificação explícita aos indicadores de origem nacional, reconhecendo a sua maior pertinência e actualidade para o contexto angolano (ver Secção 4).")
    b.h2("3.3 Período e estrutura temporal")
    b.para(f"Todas as séries são organizadas numa grelha anual fixa de 2015 a 2025, o que assegura a comparabilidade entre indicadores e ao longo do tempo. A cobertura efectiva varia entre indicadores: {cov['last2025']} dos {cov['n_sel']} indicadores têm observações até 2025, {cov['last2024']} têm 2024 como última observação e {cov['last2023']} têm 2023; por efeito do limiar de recência, nenhum indicador seleccionado tem última observação anterior a 2023. Esta heterogeneidade de cobertura é tratada pelas regras descritas na Secção 5.")
    b.h2("3.4 Anualização de séries infra-anuais")
    b.para("Quando uma fonte disponibiliza dados com periodicidade superior à anual (por exemplo, mensal ou trimestral), a série é convertida para frequência anual segundo uma regra de agregação coerente com a natureza da variável: variáveis de fluxo são somadas ou anualizadas (o PIB não petrolífero resulta da soma anual das medidas de volume trimestrais encadeadas do INE), variáveis de stock e índices são tomados em fim de período ou em média anual (a inflação usa a média anual do IPC; a variação homóloga de Dezembro é registada em linha separada). Esta regra é aplicada de forma uniforme para evitar descontinuidades artificiais na série.")
    b.h2("3.5 Vintages e revisões (v8)")
    b.para("Cada série está associada ao vintage da fonte de que foi extraída, registado na coluna 'Origem' da base potencial e no Anexo B. Na v8 foram actualizadas, com fontes primárias nacionais, as séries de crescimento do PIB (Contas Nacionais Anuais Preliminares 2025 do INE, Maio de 2026), do PIB não petrolífero (Contas Nacionais Trimestrais até ao IV trimestre de 2025), da inflação (IPCN até Dezembro de 2025 e, para leitura, até Agosto de 2026), do emprego (Anuário do IEA 2025 e primeiras publicações da nova metodologia), da dívida pública e do saldo orçamental (FMI WEO Abril 2026, coerente com o PIB rebaseado pelo INE e com os rácios do Ministério das Finanças) e da mortalidade materna (IIMS 2023-24). As séries do Banco Mundial (WDI Junho/Julho 2026), da OIT, da OMS/UNICEF, da UNCTAD, da Transparency International, do WEF e da IPU mantêm o vintage da v7 (Junho-Agosto de 2026), que é o mais recente publicado por essas instituições à data desta revisão.")

    # 4
    b.h1("4. Selecção de indicadores")
    b.h2("4.1 Critérios de selecção")
    b.para("A selecção de indicadores assenta em três critérios, que traduzem as recomendações da literatura sobre a qualidade de indicadores para índices compósitos:")
    b.bullets([("Relevância. ", "O grau em que o indicador capta de forma directa e teoricamente fundamentada o conceito da dimensão a que pertence."),
               ("Completude. ", "A extensão da cobertura temporal observada, premiando séries com menos lacunas e dados mais recentes."),
               ("Consistência. ", "A estabilidade e fiabilidade da série, penalizando rupturas metodológicas e valores anómalos não explicados.")])
    b.h2("4.2 Pontuação e ranking")
    b.para("Cada indicador candidato recebe uma pontuação total que combina os três critérios com os pesos abaixo, acrescida de uma bonificação fixa quando a fonte é nacional. Os parâmetros são configuráveis na folha de critérios do construtor:")
    b.table(["Parâmetro de selecção", "Símbolo", "Valor"], [["Peso do critério Relevância", "w₁", "0,40"], ["Peso do critério Completude", "w₂", "0,35"], ["Peso do critério Consistência", "w₃", "0,25"], ["Bonificação de fonte nacional (pontos)", "+b", "12"]], col_widths=[Cm(8), Cm(2.5), Cm(2.5)])
    b.para("Os indicadores são ordenados pela pontuação total dentro de cada dimensão. Um indicador só é elegível se tiver sinal e fronteiras de normalização válidos, dados suficientes (cobertura mínima de 80% dos anos) e um último ano observado não anterior ao limiar de recência (fixado em 2023).")
    b.h2("4.3 Mecanismo anti-duplicação")
    b.para("Para impedir que o mesmo conceito entre múltiplas vezes — por exemplo, o mesmo fenómeno medido por duas fontes, ou um indicador e um seu subcomponente — cada candidato é atribuído a um grupo temático. Dentro de cada grupo, apenas o indicador mais bem classificado é elegível para selecção, em toda a base e não apenas dentro de uma dimensão. Este mecanismo é decisivo: garante, por exemplo, que a taxa de desemprego da OIT e as taxas de desemprego do INE (novas linhas da v8) não sejam contabilizadas em simultâneo.")
    b.h2("4.4 Modo automático e modo manual")
    b.para("O construtor admite dois modos de operação. No modo automático, a selecção resulta do ranking: por dimensão, são escolhidos os melhores indicadores por Score Total até ao número-alvo, com anti-duplicação por grupo temático. No modo manual, entram apenas os indicadores marcados «Incluir». A selecção final opera em modo automático, com ajustes pontuais («Excluir») para retirar sobreposições residuais; o modo de referência é «Automática».")
    b.h2("4.5 Resultado da selecção")
    b.para(f"A aplicação destes critérios, com curadoria de não-redundância, resulta em {cov['n_sel']} indicadores distribuídos pelas oito dimensões activas, conforme o quadro seguinte. A selecção da v8 é idêntica à da v7: os sete novos candidatos (emprego INE, pobreza a 3,00 USD, cobertura do Kwenda e segurados do INSS) ficam documentados no catálogo mas não cumprem o limiar de cobertura temporal, por razões explicadas na Secção 8.6. O detalhe indicador a indicador consta do Anexo A.")
    counts = {d: sum(1 for i in ind if i["dim"] == d) for d in DIMS8}
    b.table(["Dimensão", "Candidatos", "Elegíveis", "N.º de indicadores"], [[f"{k+1}. {DIM_LABEL[d]}", V8["candidatos_por_dim"][d], V8["elegiveis_por_dim"][d], counts[d]] for k, d in enumerate(DIMS8)] + [["Total (8 dimensões activas)", sum(V8["candidatos_por_dim"][d] for d in DIMS8), sum(V8["elegiveis_por_dim"][d] for d in DIMS8), cov["n_sel"]]], col_widths=[Cm(9), Cm(2.5), Cm(2.5), Cm(3)])

    # 5
    b.h1("5. Tratamento de dados em falta")
    b.para("Nenhuma base de indicadores para um país em desenvolvimento está completa. A estratégia de tratamento de lacunas procura preservar o máximo de informação observada sem introduzir distorções, e é deliberadamente conservadora. Combina técnicas de imputação temporal e de harmonização de frequência, todas sujeitas a limites estritos e a uma salvaguarda de cobertura mínima.")
    b.h2("5.1 Técnicas de complemento de dados")
    b.table(["Técnica", "Quando se aplica", "Mecanismo de cálculo"], [
        ["Interpolação linear", "Lacunas internas da série, entre duas observações conhecidas.", "Estima o valor em falta sobre a recta que une as observações conhecidas imediatamente anterior e posterior."],
        ["Repetição do último valor (carry-forward)", "Anos no fim da série sem observação.", "Mantém constante o último valor observado, projectando-o para os anos seguintes até surgir nova observação."],
        ["Repetição do primeiro valor (carry-backward)", "Anos no início da série sem observação.", "Recua o primeiro valor observado para os anos anteriores em falta, mantendo-o constante."],
        ["Anualização de séries infra-anuais", "Indicadores com periodicidade mensal ou trimestral.", "Converte a série para frequência anual de modo coerente com a variável: soma ou anualização para fluxos; fim de período ou média anual para stocks e índices."],
    ], col_widths=[Cm(4.5), Cm(5.5), Cm(7)])
    b.para("Não são utilizadas técnicas de imputação por média, por regressão ou por modelos mais complexos: a opção por métodos simples e locais reduz o risco de introduzir estrutura artificial nos dados e mantém o processo transparente e auditável.")
    b.h2("5.2 Parâmetros e limites de imputação")
    b.table(["Regra de imputação", "Limite", "Unidade"], [["Lacuna máxima para interpolação linear", "2", "anos"], ["Carry-forward máximo no fim da série", "2", "anos"], ["Carry-backward máximo no início da série", "2", "anos"]], col_widths=[Cm(8), Cm(2), Cm(2)])
    b.para("Os limites garantem que nenhum valor é projectado para além de uma janela curta, evitando que séries com dados antigos contaminem os anos recentes com valores artificiais demasiado distantes da última observação real.")
    b.h2("5.3 Regra de cobertura mínima")
    b.para("A imputação opera ao nível do indicador; a robustez ao nível da dimensão é assegurada por uma regra de cobertura mínima. Um subíndice dimensional só é calculado, num dado ano, se pelo menos 75% dos seus indicadores tiverem valor observado ou validamente imputado nesse ano. De igual modo, o índice global só é calculado se pelo menos quatro dimensões cumprirem a sua cobertura mínima.")
    b.h2("5.4 Implicações para os anos recentes")
    b.para(f"Como {cov['last2025']} dos {cov['n_sel']} indicadores têm observação até 2025 e {cov['last2024']} até 2024, o último ano da série assenta ainda, em parte, em extrapolação por carry-forward ({cov['estimated_cells']} das {cov['total_cells']} células da base do índice são estimadas). Em consequência, 2025 deve ser interpretado como provisório e revisto à medida que as fontes publiquem dados definitivos. Esta limitação é tratada explicitamente nas Secções 9 e 11.")

    # 6
    b.h1("6. Normalização")
    b.h2("6.1 Método min-máx")
    b.para("Os indicadores são expressos em unidades heterogéneas (percentagens, dólares, anos, taxas por cem mil habitantes), pelo que têm de ser convertidos para uma escala comum antes de serem combinados. Adopta-se a normalização min-máx, que transforma cada indicador para o intervalo de 0 a 100 a partir de um valor mínimo e de um valor máximo de referência. Para um indicador em que valores mais altos são melhores, a fórmula é a diferença entre o valor e o mínimo, dividida pela amplitude entre máximo e mínimo, multiplicada por cem.")
    b.h2("6.2 Fronteiras de normalização (goalposts)")
    b.para("A escolha das fronteiras é determinante para o significado dos scores. Todas as fronteiras são definidas por referência a valores externos e normativos: a escala oficial publicada pela fonte, o limite natural da variável ou um tecto operacional documentado por referência a metas, pares internacionais e ordem de grandeza observada. Desde a v7 nenhum indicador do índice usa fronteiras derivadas da amplitude histórica; a v8 não altera nenhuma fronteira dos 41 indicadores seleccionados. A justificação linha a linha consta do documento anexo 'Justificação dos limites mínimo e máximo (v8)'.")
    b.h2("6.3 Sentido dos indicadores")
    b.para("Cada indicador tem um sentido, que indica se valores mais altos representam melhor desempenho (sentido positivo) ou pior desempenho (sentido negativo). Para indicadores de sentido negativo — como a inflação, o desemprego, a mortalidade ou a dívida — a normalização é invertida. A correcta atribuição do sentido é crítica; a v8 reverificou os sentidos dos 41 indicadores e dos sete novos candidatos.")

    # 7
    b.h1("7. Ponderação e agregação")
    b.h2("7.1 Ponderação")
    b.para("Tanto os indicadores dentro de cada dimensão como as dimensões dentro do índice recebem pesos iguais. Esta é a opção por defeito recomendada pelo Handbook da OCDE/JRC quando não existe uma base teórica ou empírica inequívoca para diferenciar pesos, e tem a vantagem decisiva da transparência. Pesos iguais não significam contributos iguais quando existem indicadores correlacionados; por essa razão, a sobreponderação implícita é evitada pela ausência de redundância entre indicadores (Secção 8), e não pela manipulação ad-hoc dos pesos nominais.")
    b.h2("7.2 Agregação intra-dimensional")
    b.para("Dentro de cada dimensão, o subíndice é a média aritmética ponderada dos scores normalizados dos indicadores disponíveis nesse ano, condicionada ao cumprimento da cobertura mínima.")
    b.h2("7.3 Agregação inter-dimensional")
    b.para("Entre dimensões, o índice global é a média geométrica ponderada dos subíndices das dimensões incluídas. A média geométrica é deliberadamente menos compensatória do que a aritmética: penaliza o desequilíbrio entre dimensões, de modo que um valor muito baixo numa dimensão não pode ser plenamente compensado por valores altos noutras.")
    b.h2("7.4 Justificação da escolha aritmética e geométrica")
    b.para("A combinação de média aritmética dentro das dimensões e média geométrica entre dimensões segue o desenho adoptado pelo Índice de Desenvolvimento Humano desde 2010. Reflecte um juízo substantivo: a substituibilidade é razoável entre indicadores de um mesmo tema, mas não entre pilares distintos do desenvolvimento.")
    b.h2("7.5 Regras de inclusão")
    b.para("A agregação respeita as regras de cobertura descritas na Secção 5. Na configuração actual, as oito dimensões activas são calculáveis em todos os anos do período, pelo que o IGDA é estimado de forma consistente ao longo de todo o intervalo 2015–2025.")

    # 8
    b.h1("8. Configuração final do índice")
    b.h2("8.1 Composição")
    b.para(f"O índice final reúne {cov['n_sel']} indicadores distribuídos pelas oito dimensões activas: cinco em Governança, sete em Macroeconomia, quatro em Capital Humano, três em Inclusão Social, cinco em Infraestruturas, cinco em Mercado de Trabalho, sete em Segurança Alimentar e Saúde e cinco em Diversificação. A selecção resulta do ranking automático por dimensão, complementado por uma curadoria que assegura a não-redundância e a coerência dos sentidos.")
    b.h2("8.2 Não-redundância")
    b.bullets([("Compósitos e componentes. ", "Não se incluem simultaneamente um índice compósito e as suas próprias componentes."),
               ("Mortalidade infantil. ", "A mortalidade infantil é representada em Capital Humano e a de menores de cinco anos em Saúde, sem duplicar séries quase idênticas."),
               ("Capacidades produtivas. ", "Utiliza-se um único índice de capacidades produtivas da UNCTAD (o índice geral)."),
               ("Governança. ", "O combate à corrupção é captado pelo Índice de Percepção da Corrupção, complementado por indicadores institucionais distintos."),
               ("Emprego. ", "As séries INE/IEA acrescentadas na v8 partilham grupo temático com as séries OIT correspondentes: só uma medida de cada conceito pode entrar no índice."),
               ("Sobreposições ligeiras. ", "Algumas medidas próximas coexistem em dimensões distintas (rendimento por habitante; crescimento não petrolífero), por captarem vertentes diferentes.")])
    b.h2("8.3 Coerência dos sentidos")
    b.para("Cada indicador tem o sentido coerente com a sua escala, para que melhorias reais sejam lidas como progresso. Por exemplo, o Índice de Percepção da Corrupção (0 = muito corrupto, 100 = muito íntegro) tem sentido positivo; a subida do CPI de Angola de 15, em 2015, para 32, em 2025, traduz-se numa melhoria da dimensão de Governança.")
    b.h2("8.4 Comparabilidade ao longo do tempo")
    b.para("As fronteiras de normalização são fixas por indicador e os sentidos são coerentes com cada escala. Em consequência, um mesmo valor tem o mesmo significado em qualquer ano, e o índice pode ser lido com segurança em termos de evolução. As substituições de séries efectuadas na v8 (Secção 8.6) aplicam-se a todo o período 2015–2025, preservando a comparabilidade intertemporal.")
    b.h2("8.5 Revisão v6 → v7: fronteiras normativas")
    b.para("A revisão v7 substituiu as três últimas fronteiras de base histórica (desemprego juvenil, produto por trabalhador e tráfego portuário) por fronteiras normativas, sem alterar indicadores, pesos ou regras de agregação. O Mercado de Trabalho corrigiu de 66,0 para 56,2 em 2025 e as Infraestruturas passaram de 48,0 para 49,9.")
    b.h2("8.6 Revisão v7 → v8: dados, metas do PDN e resposta ao feedback externo")
    b.para("A v8 resulta de uma auditoria de qualidade ao construtor v7 e do cruzamento do índice com o PDN 2023-2027 solicitado a um avaliador externo (Fable 5.1). Incidiu em cinco frentes, sem alterar a lista dos 41 indicadores, as fronteiras, os pesos ou as regras de agregação:")
    b.numbered([
        ("Actualização de séries com fontes primárias. ", "Crescimento do PIB 2025 confirmado pelas Contas Nacionais Anuais Preliminares do INE (+3,13%); PIB não petrolífero recalculado a partir das Contas Nacionais Trimestrais mais recentes (IV trimestre de 2025, com revisões de até 0,2 p.p.); inflação validada com o IPCN (média anual 2025 de 20,2%; homóloga de Dezembro 15,7%); emprego informal e formalização actualizados com o Anuário do IEA 2025."),
        ("Correcção de séries sem rastreabilidade. ", "A dívida pública e o saldo orçamental usavam valores arredondados de 'compilação interna' (o saldo de 2017 estava registado em +1,5% do PIB quando o valor oficial foi −5,7%). Foram substituídos pela série do FMI WEO de Abril de 2026, já coerente com o PIB rebaseado pelo INE em 2025 e com os rácios do Ministério das Finanças (dívida governamental de 46,6% em 2025; FMI 51,3%; Banco Mundial ≈52%). A mortalidade materna de 2025 (170) não tinha suporte documental e foi removida; o valor do IIMS 2023-24 (170 por 100 mil, IC 99–242) passa a figurar em 2024."),
        ("Alinhamento das metas 2027 com o PDN. ", f"{len(R['meta_changes'])} metas foram alteradas para as metas oficiais do PDN 2023-2027 (por exemplo, mortalidade materna 300 → 165; esperança de vida 65 → 63; electrificação 60 → 49; água 70 → 61; desemprego 20 → 25; dívida pública 60; IPC 35 → 34). As metas do PDN expressas em percentis dos Worldwide Governance Indicators foram convertidas para a escala de estimativas por deslocamento equivalente na normal padrão a partir do valor de 2022; as metas expressas em % do PIB não petrolífero foram convertidas para % do PIB. Onde o PDN não fixa meta comparável, a meta operacional da v7 foi mantida e identificada como tal. Uma nova coluna 'Origem da Meta 2027' documenta cada caso; a coluna deixa assim de conter as treze metas inconsistentes assinaladas na avaliação externa."),
        ("Emprego — séries do INE (feedback, ponto 1). ", "Foram acrescentadas ao catálogo as séries do Inquérito sobre o Emprego em Angola, originais (13.ª CIET) e harmonizadas pela OIT para a definição internacional, bem como a primeira observação da nova metodologia (19.ª–21.ª CIET, IV trimestre de 2025). Como o IEA só existe desde 2019, estas séries não cumprem a cobertura mínima de 80% e não substituem as estimativas modeladas da OIT no índice-mãe; a nova folha 12_Emprego_INE calcula um subíndice alternativo 2019–2025 com as séries INE, que isola o efeito do desfasamento de vintage (a série INE harmonizada regista a descida do desemprego de 13,9% para 10,4% em 2025, que a estimativa modelada ainda não incorpora)."),
        ("Inclusão — pobreza e protecção social (feedback, ponto 2). ", "A taxa de pobreza foi desdobrada nas linhas de 2,15 USD (PPC 2017, referência do PDN: 31% → 28%) e de 3,00 USD (PPC 2021, novo padrão do Banco Mundial: 39,3% em 2018), e foram criados dois indicadores de protecção social: cobertura do programa Kwenda (agregados beneficiários acumulados em % dos agregados familiares do Censo 2024: 6,7% em 2022 → 14,8% em 2025) e segurados inscritos no INSS (1,97 milhões em 2020 → 3,34 milhões em 2025; meta PDN 4,3 milhões). Sem inquérito de despesas posterior ao IDREA 2018-19 e com programas iniciados em 2020, nenhum destes indicadores atinge ainda a cobertura temporal exigida; ficam documentados para incorporação automática quando a cumprirem."),
    ])
    b.para(f"O efeito agregado das alterações é contido: o IGDA 2025 passa de {f1(igda7[-1])} para {f1(igda[-1])} e a variação 2015–2025 mantém-se em {f1(igda[-1]-igda[0], sign=True)} pontos. A dimensão mais afectada é a Estabilidade Macroeconómica, onde a série oficial do saldo orçamental revela um défice em 2017 e em 2025 que a série anterior escondia; a Diversificação ajusta-se marginalmente pela revisão do PIB não petrolífero. O quadro seguinte compara as duas versões.")
    rows = []
    for d in DIMS8:
        rows.append([DIM_LABEL[d], f1(sub7[d][0]), f1(sub7[d][-1]), f1(sub[d][0]), f1(sub[d][-1]), f1(sub[d][-1] - sub7[d][-1], sign=True)])
    rows.append(["IGDA-BDA", f1(igda7[0]), f1(igda7[-1]), f1(igda[0]), f1(igda[-1]), f1(igda[-1] - igda7[-1], sign=True)])
    b.table(["Dimensão", "v7 · 2015", "v7 · 2025", "v8 · 2015", "v8 · 2025", "Δ v8−v7 (2025)"], rows, col_widths=[Cm(6.5), Cm(2), Cm(2), Cm(2), Cm(2), Cm(2.5)])
    b.para("Nota sobre o ficheiro v7: o livro Excel distribuído com a v7 tinha sido gravado sem recálculo e conservava em cache os valores da v6 (IGDA 2025 = 46,4; Mercado de Trabalho = 66,0), embora as fórmulas e os documentos correspondessem à v7 (45,7; 56,2). Os valores 'v7' deste quadro resultam do recálculo integral das fórmulas da v7. O livro v8 é gravado com recálculo integral forçado na abertura.")

    # 9
    b.h1("9. Pressupostos e limitações")
    b.h2("9.1 Pressupostos centrais")
    b.bullets(["A escolha das oito dimensões activas representa adequadamente o conceito multidimensional de desenvolvimento adoptado.",
               "Pesos iguais são uma aproximação aceitável na ausência de uma base inequívoca para diferenciar a importância de indicadores e dimensões.",
               "A substituibilidade é razoável entre indicadores de uma mesma dimensão, mas limitada entre dimensões — pressuposto materializado na escolha das funções de agregação.",
               "A imputação conservadora de lacunas, dentro dos limites fixados, não distorce materialmente os resultados."])
    b.h2("9.2 Limitações de dados")
    b.bullets([
        ("Anos recentes provisórios. ", f"Pouco mais de metade dos indicadores ({cov['last2025']} em {cov['n_sel']}) chega a 2025; o último ano assenta em parte em extrapolação e deve ser lido como provisório."),
        ("Lacunas em Inclusão. ", "A dimensão de Inclusão Social e Protecção conta apenas com três indicadores. Os candidatos de pobreza e protecção social acrescentados na v8 documentam a lacuna mas não a colmatam: não existe inquérito de despesas posterior a 2018-19 e as séries do Kwenda e do INSS começam em 2020."),
        ("Dependência de estimativas internacionais. ", "Vários indicadores provêm de modelações de organismos internacionais e não de registos administrativos nacionais. O caso mais relevante é o emprego: as estimativas modeladas da OIT para 2025 ainda não incorporam a queda do desemprego medida pelo INE (Secção 10.5 e folha 12_Emprego_INE)."),
        ("Quebra de série no IEA. ", "Desde o IV trimestre de 2025 o INE aplica a nova metodologia (19.ª–21.ª CIET), que exclui a produção para autoconsumo do emprego: a taxa de desemprego passa de 26,9% para 20,1% e a taxa de emprego de ~63% para ~40% sem alteração real do mercado. Não existe retropolação; as metas do PDN (25% em 2027) foram fixadas na definição antiga."),
        ("Rebasing do PIB. ", "O INE rebaseou o PIB em 2025; rácios em % do PIB de vintages anteriores (dívida, saldo, crédito) não são comparáveis com os actuais. A v8 usa um único vintage (FMI WEO Abril 2026) para dívida e saldo."),
        ("Fontes internacionais não reverificadas nesta revisão. ", "Os valores WGI, WUENIC, SOFI, UNCTAD, CPI 2025, GGGI e IPU mantêm o vintage da v7 (Junho–Agosto de 2026); a sua reverificação directa nas fontes não foi possível no ambiente desta revisão e deve ser feita no ciclo seguinte."),
    ])
    b.h2("9.3 Limitações metodológicas")
    b.bullets([
        ("Níveis não comparáveis entre dimensões. ", "Cada indicador é normalizado contra as suas próprias fronteiras, cuja amplitude útil varia substancialmente. Um subíndice mais alto do que outro não significa melhor desempenho relativo; os subíndices devem ser lidos face à sua própria trajectória no tempo (Secção 10.4)."),
        ("Variáveis de política vs. resultado. ", "O indicador de anos de escolaridade obrigatória (Capital Humano) é uma variável legal, não de resultado: o salto de 2016 reflecte a reforma que alargou a obrigatoriedade. A avaliação externa recomenda a sua substituição por um indicador de resultado (por exemplo, taxa de conclusão do ensino primário) quando existir série elegível; a alteração da composição foi deliberadamente adiada para não introduzir uma quebra nesta revisão (Secção 11)."),
        ("Sensibilidade às escolhas de desenho. ", "Como qualquer índice compósito, o IGDA é sensível às opções de selecção, normalização, ponderação e agregação. A quantificação dessa sensibilidade — passo 7 do quadro da OCDE — mantém-se como desenvolvimento prioritário."),
    ])

    # 10
    b.h1("10. Interpretação dos resultados")
    b.h2("10.1 Leitura da escala")
    b.para("O IGDA e cada subíndice variam entre 0 e 100. O valor não tem uma unidade material; deve ser lido em termos relativos, face às fronteiras de normalização, à evolução temporal e à comparação entre dimensões. Subidas e descidas ao longo do tempo são mais informativas do que o nível absoluto isolado.")
    b.h2("10.2 Bandas de classificação")
    b.table(["Intervalo do score", "Classificação"], [["80–100", "Muito Elevado"], ["60–80", "Elevado"], ["40–60", "Médio"], ["20–40", "Baixo"], ["0–20", "Muito Baixo"]], col_widths=[Cm(4), Cm(5)])
    b.h2("10.3 Série de referência (2015–2025)")
    b.para("Nesta versão, e tratando 2025 como provisório, o índice global apresenta a seguinte trajectória — ascendente no conjunto do período —, com as oito dimensões activas a contribuir em todos os anos:")
    b.table([str(y) for y in YEARS], [[f1(v) for v in igda]], font_size=9)
    b.para(f"No conjunto do período, o IGDA sobe de {f1(igda[0])} (2015) para {f1(igda[-1])} (2025), uma melhoria de {f1(igda[-1]-igda[0])} pontos. O mínimo da série ocorre em 2020 ({f1(igda[5])}, choque pandémico e petrolífero, único ano na banda Baixo); a recuperação é contínua desde então, com estabilização em 2025 ({f1(igda[-1]-igda[-2], sign=True)} p.p., provisório).")
    deltas = {d: sub[d][-1] - sub[d][0] for d in DIMS8}
    up = sorted([d for d in DIMS8 if deltas[d] > 0], key=lambda d: -deltas[d]); down = sorted([d for d in DIMS8 if deltas[d] <= 0], key=lambda d: deltas[d])
    b.para("Por dimensão, entre 2015 e 2025: melhoraram " + ", ".join(f"{DIM_LABEL[d]} ({f1(deltas[d], sign=True)})" for d in up) + "; recuaram " + ", ".join(f"{DIM_LABEL[d]} ({f1(deltas[d], sign=True)})" for d in down) + ". Estas variações são medidas em pontos de score, na escala própria de cada dimensão; a leitura comparável faz-se pela via da Secção 10.4.")
    b.table(["Dimensão"] + [str(y) for y in YEARS] + ["Δ 2015–25"], [[DIM_LABEL[d]] + [f1(v) for v in sub[d]] + [f1(deltas[d], sign=True)] for d in DIMS8] + [["IGDA-BDA"] + [f1(v) for v in igda] + [f1(igda[-1]-igda[0], sign=True)]], font_size=8)
    b.h2("10.4 Escala comum de evolução (2015 = 100)")
    b.para("Para permitir uma comparação legítima entre dimensões, o construtor inclui uma folha que reexprime cada subíndice face ao seu próprio valor de 2015 (Índice(t) = 100 × Subíndice(t) ÷ Subíndice(2015)). Por partirem todas do mesmo valor, as trajectórias tornam-se directamente comparáveis, o que os níveis não permitem.")
    order = sorted(DIMS8, key=lambda d: -esc[d][-1])
    b.para("Nesta escala, " + ", ".join(f"{DIM_LABEL[d]} ({fmt(esc[d][-1]-100, 1, pct=True, sign=True)})" for d in order[:3]) + " lideram o progresso acumulado; recuam " + ", ".join(f"{DIM_LABEL[d]} ({fmt(esc[d][-1]-100, 1, pct=True, sign=True)})" for d in order if esc[d][-1] < 100) + f". O IGDA global progride {fmt(esc['IGDA'][-1]-100, 1, pct=True, sign=True)}. Esta escala compara ritmos de progresso, não patamares de desenvolvimento.")
    b.table(["Dimensão"] + [str(y) for y in YEARS], [[DIM_LABEL[d]] + [fmt(v, 0) for v in esc[d]] for d in order] + [["IGDA-BDA"] + [fmt(v, 0) for v in esc["IGDA"]]], font_size=8)
    b.h2("10.5 Cruzamento com o PDN 2023-2027 (passo 9 do quadro OCDE/JRC)")
    b.para("A avaliação externa da v7 cruzou a trajectória recente de cada dimensão (2023–2025) com as metas e o Balanço oficial do PDN 2023-2027. A v8 retoma e actualiza esse cruzamento com os dados desta revisão. Em seis das oito dimensões o índice e o plano contam a mesma história, o que constitui uma validação externa da construção; nos dois pilares do PDN as leituras divergem — o capital humano avança devagar e a segurança alimentar regride.")
    d23 = {d: sub[d][-1] - sub[d][8] for d in DIMS8}
    cruz = [
        ["Governança", "Metas modestas (percentis WGI +4 a +7 p.p.; IPC 33 → 34); IPC estagnado em 32 (2024 e 2025)", f"{f1(d23['Governança'], sign=True)} p.p.; estagnação após os ganhos de 2018–23 (estabilidade política −5,2)", "Consistente"],
        ["Macroeconomia", "Dívida 66 → 60% do PIB largamente superada (FMI 51,3% em 2025); reservas >6 meses; inflação desviou (27,5% em Dez-2024) e corrigiu (15,7% em Dez-2025; 8,8% em Ago-2026); défice de 4,1% do PIB em 2025", f"{f1(d23['Macroeconomia'], sign=True)} p.p.; dívida (+16,3) e crescimento (+7,2) puxam; inflação (−13,0) e saldo orçamental (−6,3) pesam", "Consistente"],
        ["Capital Humano", "Incrementos pequenos (esperança de vida 62 → 63; alfabetização 76 → 78%); compromisso de subir a educação para 11,8% da despesa", f"{f1(d23['Capital Humano'], sign=True)} p.p.; despesa em educação em queda (dotações OGE: 2,0% do PIB em 2024, 1,8% em 2025, 1,7% em 2026)", "Consistente, com alerta"],
        ["Inclusão Social", "Pobreza 31 → 28%; Kwenda com 1,35 milhões de agregados acumulados (meta anual 2025: 1,8 milhões)", f"{f1(d23['Inclusão Social'], sign=True)} p.p., quase só por representação política (mulheres no parlamento +9,1)", "Parcial — o índice não mede pobreza (candidatos criados, sem série elegível)"],
        ["Infraestruturas", "Electrificação 43 → 49%: 48% em 2025 (MINPLAN); água tratada estagnada; PIP com 846 projectos por iniciar (2024)", f"{f1(d23['Infraestruturas'], sign=True)} p.p., lento; 2025 por carry-forward em electrificação e energia limpa", "Consistente"],
        ["Mercado de Trabalho", "Desemprego 30 → 25%: 28,3% em 2025 (Anuário IEA); 20,1% no IV trim 2025 na nova metodologia", f"{f1(d23['Mercado Trabalho'], sign=True)} p.p., estagnado; leitura INE (folha 12) mostra melhoria em 2025 — divergência de vintage; produtividade em queda é a divergência substantiva", "Inconsistente — desfasamento + produtividade"],
        ["Saúde e Seg. Alimentar", "Metas ambiciosas (mortalidade <5 anos 69 → 51; materna 199 → 165); Balanço com vacinação (76% vs 80%) e produção alimentar abaixo da meta", f"{f1(d23['Saúde/Alimentar'], sign=True)} p.p.; malária (−0,7), cobertura vacinal (−10,0) e desnutrição pesam; indicadores de resultado (mortalidade) ainda melhoram", "Consistente — alerta principal"],
        ["Diversificação", "Não petrolífero +4,6%/ano em média (INE: +5,2% em 2025); IDE e exportações não petrolíferas aquém", f"{f1(d23['Diversificação'], sign=True)} p.p.; PIB não petrolífero (+10,8) puxa; crédito ao sector privado (−1,1) e concentração das exportações sem sinal", "Parcial"],
    ]
    b.table(["Dimensão", "PDN 2023-2027 e Balanço", "IGDA 2023–2025 (v8)", "Veredicto"], cruz, col_widths=[Cm(3), Cm(6.5), Cm(6), Cm(3)], font_size=8)
    b.para("Três conclusões para o Comité Executivo. Primeira: o cruzamento com o plano oficial não desmente o índice. Segunda: o pilar da segurança alimentar do PDN está a regredir nos indicadores de processo (vacinação, desnutrição, malária) antes de as metas de mortalidade o reflectirem — é a leitura de maior valor do índice. Terceira: o Balanço do PDN reconhece que, dos 1 036 indicadores definidos, 856 não apresentavam execução no I trimestre de 2025 e que os reportados são sobretudo de produto; o IGDA mede resultados e complementa o PDN no que este não mede.")

    # 11
    b.h1("11. Recomendações para desenvolvimento futuro")
    b.para("O índice está operacional e auditável. As recomendações seguintes constituem o programa de aperfeiçoamento futuro:")
    b.bullets([
        ("Concluído na v8 — metas do PDN. ", "Coluna de metas alinhada com as metas oficiais e documentada linha a linha; as metas operacionais remanescentes estão identificadas."),
        ("Concluído na v8 — leitura INE do emprego. ", "Séries INE/IEA no catálogo e folha 12_Emprego_INE. Passo seguinte: quando a OIT publicar estimativas modeladas que incorporem o IEA 2025 (Novembro de 2026), actualizar LAB001–LAB003; a partir de 2028 (≥9 anos de IEA), avaliar a substituição directa pelas séries INE harmonizadas."),
        ("Em curso — Inclusão. ", "Candidatos de pobreza (2,15 e 3,00 USD), Kwenda e INSS criados; obter do INSS a série anual de segurados 2015–2025 (tornaria INC033 elegível de imediato) e acompanhar o próximo inquérito de despesas do INE."),
        ("Substituir a variável de política em Capital Humano. ", "Avaliar a troca dos anos de escolaridade obrigatória por um indicador de resultado com série elegível (taxa de conclusão do ensino primário: 58% em 2025 vs. meta anual 69%, MINPLAN)."),
        ("Estabelecer um referencial externo de comparação. ", "Reescalar cada indicador contra a distribuição observada num grupo de países pares, à maneira da medida de distância à fronteira."),
        ("Consolidar os anos recentes. ", "Rever 2024 e 2025 à medida que as fontes publiquem dados definitivos (WDI, WGI 2025, WUENIC, IGME, SOFI 2026, UNCTAD PCI) e levantar progressivamente o estatuto de provisório."),
        ("Activar as dimensões inactivas. ", "Ambiente (10 elegíveis), Demografia (7) e Digital (9) têm já candidatos elegíveis; a sua activação é uma decisão de desenho a fundamentar (número de indicadores alvo e não-redundância)."),
        ("Realizar análise de sensibilidade. ", "Executar o passo 7 do quadro da OCDE, testando a robustez do IGDA a pesos alternativos, à escolha entre agregação aritmética e geométrica e à inclusão ou exclusão de indicadores marginais."),
    ])

    # 12
    b.h1("12. Implementação técnica")
    b.h2("12.1 Arquitectura do construtor")
    b.para("O índice é implementado num livro de cálculo integralmente ligado por fórmulas, organizado em folhas sequenciais: um painel de síntese (00); os critérios de selecção e parâmetros (01); a base potencial (02, o catálogo completo de candidatos, onde opera o motor de selecção, agora com a coluna 'Origem da Meta 2027'); a base do índice (03); os dados ajustados (04); as bandeiras de qualidade (05); a normalização (06); os subíndices e o IGDA (07); as visualizações (08); a metodologia (09); as fontes (10); a escala comum de evolução (11); a leitura complementar do emprego com dados do INE (12, nova na v8); e o registo de alterações v7 → v8 com a comparação de resultados (13, nova na v8). Toda a cadeia é dinâmica.")
    b.h2("12.2 Cadeia de cálculo")
    b.para("A sequência de cálculo parte da base potencial, onde o motor de selecção pontua e ordena os candidatos; os seleccionados alimentam a base do índice; seguem-se os dados ajustados, onde se aplica a imputação; a normalização, que converte tudo para a escala de 0 a 100 com os sentidos correctos; e os subíndices e o IGDA, calculados com as regras de cobertura e as funções de agregação descritas. O painel, as visualizações e as folhas 11 a 13 lêem o resultado final.")
    b.h2("12.3 Recalibração e reprodutibilidade")
    b.para("A recalibração faz-se através de dois pontos de controlo: a folha de critérios e a base potencial. Qualquer alteração, seguida de recálculo, reconstrói o índice de forma reproduzível. A v8 foi verificada com um motor de cálculo independente, que replica todas as fórmulas do construtor e reproduziu exactamente os valores da v6 e da v7; o livro é gravado com a opção de recálculo integral na abertura, para evitar a repetição do problema de valores em cache detectado na v7.")

    # Anexo A
    b.h1(f"Anexo A — Estrutura completa: {cov['n_sel']} indicadores")
    b.para("O quadro seguinte lista os indicadores do índice, agrupados pelas oito dimensões activas, com a respectiva unidade, sentido, cobertura temporal observada (percentagem de anos com observação em 2015–2025, antes de imputação), fonte e meta 2027 (alinhada com o PDN quando existe correspondência; 'op.' identifica metas operacionais mantidas).")
    rows = []
    wsmeta = {}
    import openpyxl
    wb = openpyxl.load_workbook("/home/user/bda-indice/Versão Final/Finalíssima/files/IGDA_BDA_Construtor_v8.xlsx", data_only=True)
    ws = wb["02_Base_Potencial"]; hdr = [c.value for c in ws[4]]; ci = {h: i for i, h in enumerate(hdr) if h is not None}
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[0]:
            wsmeta[r[0]] = (r[ci["Meta 2027"]], r[ci["Origem da Meta 2027"]])
    for d in DIMS8:
        rows.append([f"{DIMS8.index(d)+1}. {DIM_LABEL[d]}", "", "", "", "", "", ""])
        for i in sorted([i for i in ind if i["dim"] == d], key=lambda x: x["slot"]):
            meta, orig = wsmeta.get(i["id"], (None, ""))
            mtxt = "–" if meta is None else (fmt(meta, 2) if abs(meta) < 10 and not float(meta).is_integer() else fmt(meta, 0))
            if orig and not str(orig).startswith("PDN") and meta is not None:
                mtxt += " (op.)"
            rows.append([i["id"], i["nome"], i["unidade"], i["sentido"].replace("-", "−"), pct1(100 * i["cobertura"]), i["fonte"], mtxt])
    b.table(["Código", "Indicador", "Unidade", "Sentido", "Cobertura", "Fonte", "Meta 2027"], rows, col_widths=[Cm(1.8), Cm(6), Cm(2.6), Cm(1.3), Cm(1.8), Cm(3.5), Cm(1.8)], font_size=8)

    # Anexo B
    b.h1("Anexo B — Fontes, vintages e siglas")
    b.para("B.1 Fontes de dados e vintages usados na v8")
    b.table(["Fonte", "Séries", "Vintage / publicação", "Estado na v8"], [
        ["INE — Contas Nacionais Anuais", "Crescimento do PIB (MAC001)", "CN Anuais Preliminares 2025 (05-05-2026); série 2002–2024 (24-03-2026)", "Actualizado"],
        ["INE — Contas Nacionais Trimestrais", "PIB não petrolífero (MAC014/DIV017)", "IV trimestre 2025 (ficheiro 09-05-2026)", "Recalculado"],
        ["INE — IPCN", "Inflação média anual e homóloga (MAC003/MAC027)", "Dezembro 2025; leitura até Agosto 2026", "Validado"],
        ["INE — IEA", "Emprego informal, formalização, séries INE (LAB004/005/028-031)", "Anuário 2025 (Abril 2026); I e II trim 2026 (nova metodologia)", "Actualizado / novo"],
        ["INE — IIMS 2023-24", "Mortalidade materna, infantil, <5 anos", "Relatório final (2025)", "Corrigido"],
        ["INE — Censo 2024", "População, agregados familiares (denominador do Kwenda)", "Resultados definitivos (Nov-2025)", "Novo"],
        ["FMI — WEO", "Dívida e saldo do governo geral (MAC004/005/016/019/020)", "Abril 2026 (PIB rebaseado)", "Substituído"],
        ["MINFIN/UGD; MINPLAN; MAPTSS/INSS; FAS", "Rácios de dívida (validação), Balanço do PDN, Kwenda, INSS", "PAE 2026; RBPDN I trim 2025; Boletim 2025", "Novo / contexto"],
        ["Banco Mundial — WDI/WGI", "Restantes séries macro, sociais e de governança", "WDI Junho/Julho 2026; WGI 2025 (ano 2024)", "Mantido da v7"],
        ["OIT — ILOSTAT", "Emprego, desemprego, produtividade (modeladas) e estimativas nacionais harmonizadas", "Modeladas Nov-2025; nacionais Abr-2026", "Mantido / novo"],
        ["OMS, UNICEF, FAO, UNCTAD, TI, WEF, IPU", "Saúde, nutrição, capacidades produtivas, portos, índices", "Edições 2025–2026 (v7)", "Mantido da v7"],
        ["PDN 2023-2027 (Diário da República / AUDA-NEPAD)", "Metas 2027", "Outubro 2023", "Novo (metas)"],
    ], col_widths=[Cm(4.5), Cm(5.5), Cm(5), Cm(2.5)], font_size=8)
    b.para("Referência metodológica: OCDE e Joint Research Centre da Comissão Europeia (2008), Handbook on Constructing Composite Indicators: Methodology and User Guide.")
    b.para("B.2 Lista de siglas")
    b.table(["Sigla", "Designação"], [
        ["BDA", "Banco de Desenvolvimento de Angola"], ["BNA", "Banco Nacional de Angola"], ["CIET", "Conferência Internacional de Estatísticos do Trabalho (OIT)"], ["CPI", "Corruption Perceptions Index (Índice de Percepção da Corrupção)"],
        ["FAO", "Organização das Nações Unidas para a Alimentação e a Agricultura"], ["FMI", "Fundo Monetário Internacional"], ["IDH", "Índice de Desenvolvimento Humano"], ["IEA", "Inquérito sobre o Emprego em Angola (INE)"],
        ["IGDA", "Índice Global de Desenvolvimento de Angola"], ["IIMS", "Inquérito de Indicadores Múltiplos e de Saúde"], ["INE", "Instituto Nacional de Estatística"], ["INSS", "Instituto Nacional de Segurança Social"],
        ["IPC", "Índice de Preços no Consumidor"], ["IPU", "Inter-Parliamentary Union"], ["JRC", "Joint Research Centre (Comissão Europeia)"], ["MAPTSS", "Ministério da Administração Pública, Trabalho e Segurança Social"],
        ["MINEA", "Ministério da Energia e Águas"], ["MINFIN", "Ministério das Finanças"], ["MINPLAN", "Ministério do Planeamento"], ["MINSA", "Ministério da Saúde"], ["OCDE", "Organização para a Cooperação e Desenvolvimento Económico"],
        ["OIT", "Organização Internacional do Trabalho"], ["OMS", "Organização Mundial da Saúde"], ["PDN", "Plano de Desenvolvimento Nacional 2023-2027"], ["PIB", "Produto Interno Bruto"], ["PPC", "Paridade de Poder de Compra"],
        ["RNB", "Rendimento Nacional Bruto"], ["TEU", "Twenty-foot Equivalent Unit"], ["UGD", "Unidade de Gestão da Dívida (MINFIN)"], ["UNCTAD", "Conferência das Nações Unidas sobre Comércio e Desenvolvimento"],
        ["UNICEF", "Fundo das Nações Unidas para a Infância"], ["WDI", "World Development Indicators (Banco Mundial)"], ["WEF", "Fórum Económico Mundial"], ["WEO", "World Economic Outlook (FMI)"], ["WGI", "Worldwide Governance Indicators (Banco Mundial)"],
    ], col_widths=[Cm(2.5), Cm(13)], font_size=8)

    # Anexo C
    b.h1("Anexo C — Registo resumido de alterações v7 → v8")
    b.para("O registo completo, linha a linha, consta da folha 13_Alteracoes_v8 do construtor. Resumo:")
    tipos = {}
    for e in R["log"]:
        tipos.setdefault(e["tipo"], []).append(e)
    rows = []
    for t, es in tipos.items():
        rows.append([t.capitalize(), len(es), "; ".join(sorted(set(e["id"] for e in es if e["id"])))[:400] or "—"])
    rows.append(["Metas 2027 alteradas", len(R["meta_changes"]), "; ".join(m[0] for m in R["meta_changes"])])
    b.table(["Tipo de alteração", "N.º", "Indicadores"], rows, col_widths=[Cm(4), Cm(1.5), Cm(11.5)], font_size=8)
    b.p("Banco de Desenvolvimento de Angola", "Source Code")
    b.save(out)


if __name__ == "__main__":
    build("/home/user/bda-indice/Versão Final/Finalíssima/files/IGDA_BDA_Nota_Metodologica_v7.docx", sys.argv[1])
    print("ok")
