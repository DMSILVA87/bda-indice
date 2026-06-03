# BNA - Estatísticas / Consultar dados

Compilado em 2026-05-09 19:14 a partir do portal oficial do BNA: <https://bna.ao/#/pt>.

- Ficheiros descarregados: 74
- Datasets exportados: 13
- Páginas HTML preservadas: 1

Cobertura por área:
- Central de Balanços: 7 item(ns)
- Documentos Metodológicos: 2 item(ns)
- Estatísticas Externas: 33 item(ns)
- Estatísticas Monetárias e Financeiras: 17 item(ns)
- Evolução das Reservas Internacionais: 20 item(ns)
- Preços e Contas Nacionais: 9 item(ns)

Notas:
- Os ficheiros do BNA são obtidos pela API pública `file/getPDF/v2`, que devolve PDF ou Excel conforme o documento original.
- As séries tabulares foram exportadas com o maior período devolvido pela API, usando `$top` elevado e sem restringir datas, salvo no caso das reservas diárias onde foi usado um intervalo aberto 1900-2099.
- O ficheiro `00_manifesto_bna.csv` contém fonte, periodicidade, intervalo temporal e caminho local de cada item.
