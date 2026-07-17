#!/usr/bin/env python3
"""
Compilação reprodutível do PDN 2023-2027 e da página de balanço do MINPLAN.

Entradas locais:
  fontes/minplan_relatorios_balanco_pdn.html
  fontes/PDN_Angola_2023-2027.pdf

Saídas:
  dados/*.csv
  01_Execucao_PDN_2025.xlsx
  02_Execucao_vs_Metas_PDN_2027.xlsx

O script não preenche os separadores de domínio que estão vazios no HTML
publicado pelo MINPLAN. Essa ausência é registada explicitamente nos
ficheiros para evitar confundir "sem dados publicados" com zero execução.
"""

from __future__ import annotations

import argparse
import csv
import re
import unicodedata
from datetime import date
from pathlib import Path
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from pypdf import PdfReader


ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / "fontes"
DATA_DIR = ROOT / "dados"

HTML_URL = "https://www.minplan.gov.ao/en/publicacoes/relatorios-balanco-pdn"
PDN_URL_REQUESTED = (
    "https://www.nepad.org/sites/default/files/2024-07/"
    "20231030%283%29_layout_Final_Angola_PDN%202023-2027-1.pdf"
)
PDN_URL_MIRROR = "https://www.mpla.ao/wp-content/uploads/2023/12/PDN_Angola_2023-2027.pdf"

HTML_FILE = SOURCE_DIR / "minplan_relatorios_balanco_pdn.html"
PDF_FILE = SOURCE_DIR / "PDN_Angola_2023-2027.pdf"
PDF_TEXT_FILE = SOURCE_DIR / "PDN_Angola_2023-2027_texto.txt"

DOMAINS = [
    "Saúde",
    "Protecção Social",
    "Comunicação",
    "Educação e Formação Profissional",
    "Infraestruturas, Habitação e Serviços Comunitários",
    "Fomento à Produção",
    "Económico",
]

YEAR_COLUMNS = ["2022_Base", "2023", "2024", "2025"]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/125.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Language": "pt-PT,pt;q=0.9,en;q=0.8",
}


def clean_text(value: Any) -> str:
    return " ".join(str(value or "").replace("\xa0", " ").split())


def slug(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9]+", "_", ascii_text).strip("_").lower()


def parse_web_number(value: str) -> int | float | None:
    """Converte valores do quadro HTML sem transformar percentagens em fracções."""
    raw = clean_text(value)
    if not raw or raw in {"-", "—", "–"}:
        return None
    raw = raw.replace(" ", "").replace("%", "")
    if not re.fullmatch(r"[-+]?\d+(?:[,.]\d+)?", raw):
        return None
    if "," in raw:
        raw = raw.replace(".", "").replace(",", ".")
    try:
        number = float(raw)
    except ValueError:
        return None
    return int(number) if number.is_integer() else number


def format_pt_number(value: int | float | None, unit: str = "") -> str:
    if value is None:
        return "-"
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    elif isinstance(value, float):
        text = f"{value:g}".replace(".", ",")
    else:
        text = str(value)
    if "%" in unit:
        return f"{text}%"
    return text


def fetch_url(url: str, timeout: int = 180) -> bytes:
    request = Request(url, headers=HEADERS)
    with urlopen(request, timeout=timeout) as response:
        return response.read()


def ensure_sources(refresh: bool = False) -> list[dict[str, str]]:
    """Mantém cópias locais das fontes e regista fallback do PDF."""
    SOURCE_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    manifest: list[dict[str, str]] = []

    if refresh or not HTML_FILE.exists() or HTML_FILE.stat().st_size == 0:
        try:
            HTML_FILE.write_bytes(fetch_url(HTML_URL))
            html_status = "descarregado"
        except (HTTPError, URLError, TimeoutError, OSError) as exc:
            html_status = f"falhou: {exc}"
    else:
        html_status = "já existente localmente"
    manifest.append(
        {
            "Fonte": "MINPLAN — página de relatórios de balanço",
            "URL solicitada": HTML_URL,
            "URL usada": HTML_URL,
            "Ficheiro local": str(HTML_FILE.relative_to(ROOT)),
            "Estado": html_status,
        }
    )

    pdf_url_used = PDN_URL_REQUESTED
    if refresh or not PDF_FILE.exists() or PDF_FILE.stat().st_size == 0:
        try:
            PDF_FILE.write_bytes(fetch_url(PDN_URL_REQUESTED, timeout=300))
            pdf_status = "descarregado da URL solicitada"
        except (HTTPError, URLError, TimeoutError, OSError) as first_exc:
            try:
                PDF_FILE.write_bytes(fetch_url(PDN_URL_MIRROR, timeout=300))
                pdf_url_used = PDN_URL_MIRROR
                pdf_status = f"fallback de espelho acessível; URL primária falhou: {first_exc}"
            except (HTTPError, URLError, TimeoutError, OSError) as second_exc:
                pdf_status = f"falhou: primária={first_exc}; espelho={second_exc}"
    else:
        pdf_status = "já existente localmente"
    manifest.append(
        {
            "Fonte": "PDN Angola 2023-2027 — AUDA-NEPAD / espelho",
            "URL solicitada": PDN_URL_REQUESTED,
            "URL usada": pdf_url_used,
            "Ficheiro local": str(PDF_FILE.relative_to(ROOT)),
            "Estado": pdf_status,
        }
    )

    write_csv(DATA_DIR / "00_manifesto_fontes.csv", manifest, list(manifest[0]))
    return manifest


def extract_pdn_text() -> dict[str, Any]:
    """Extrai texto paginado para auditoria e devolve metadados do PDF."""
    if not PDF_FILE.exists() or PDF_FILE.stat().st_size == 0:
        return {"Estado": "PDF local indisponível"}
    reader = PdfReader(str(PDF_FILE))
    pages: list[str] = []
    for number, page in enumerate(reader.pages, start=1):
        pages.append(f"===== PDF PAGE {number} =====\n{page.extract_text() or ''}")
    PDF_TEXT_FILE.write_text("\n\n".join(pages), encoding="utf-8")
    return {
        "Estado": "OK",
        "Páginas PDF": len(reader.pages),
        "Título": "Plano de Desenvolvimento Nacional 2023-2027",
        "Data de extração": date.today().isoformat(),
    }


def parse_html_source() -> dict[str, Any]:
    if not HTML_FILE.exists() or HTML_FILE.stat().st_size == 0:
        raise FileNotFoundError(f"Fonte HTML não encontrada: {HTML_FILE}")

    soup = BeautifulSoup(HTML_FILE.read_text(encoding="utf-8", errors="replace"), "html.parser")
    panels = soup.select('[role="tabpanel"]')
    indicators: list[dict[str, Any]] = []
    domain_rows: list[dict[str, Any]] = []

    for panel in panels:
        trigger = soup.find(id=panel.get("aria-labelledby", ""))
        domain = clean_text(trigger.get_text(" ", strip=True) if trigger else panel.get("id", ""))
        if domain not in DOMAINS:
            continue
        table = panel.find("table")
        if table is None:
            domain_rows.append(
                {
                    "Domínio": domain,
                    "Estado": "Sem dados publicados no HTML",
                    "Nº indicadores": 0,
                    "Observação": (
                        "O separador existe no site, mas o painel recebido não contém "
                        "tabela nem valores."
                    ),
                    "Fonte URL": HTML_URL,
                }
            )
            continue

        table_rows = table.find_all("tr")
        if not table_rows:
            continue
        headers = [clean_text(cell.get_text(" ", strip=True)) for cell in table_rows[0].find_all(["th", "td"])]
        normalized_headers = ["2022_Base" if h == "2022*" else h for h in headers]
        count = 0
        for row in table_rows[1:]:
            cells = [clean_text(cell.get_text(" ", strip=True)) for cell in row.find_all(["th", "td"])]
            if len(cells) != len(normalized_headers):
                continue
            record: dict[str, Any] = {
                "ID": f"IMP-{slug(domain)}-{slug(cells[0])}",
                "Domínio": domain,
                "Indicador": cells[0],
                "Unidade": cells[1],
                "Fonte": cells[-1],
                "Fonte URL": HTML_URL,
                "Estado": "Extraído",
                "Observação": "* 2022 é o ano base indicado na fonte.",
            }
            for header, value in zip(normalized_headers[2:-1], cells[2:-1]):
                record[header] = parse_web_number(value)
                record[f"{header}_Original"] = value
            indicators.append(record)
            count += 1
        domain_rows.append(
            {
                "Domínio": domain,
                "Estado": "Dados extraídos",
                "Nº indicadores": count,
                "Observação": "Tabela de indicadores presente no HTML renderizado.",
                "Fonte URL": HTML_URL,
            }
        )

    # Caso a página altere a estrutura e algum separador não seja encontrado,
    # mantemos o inventário esperado para tornar a lacuna visível.
    seen = {row["Domínio"] for row in domain_rows}
    for domain in DOMAINS:
        if domain not in seen:
            domain_rows.append(
                {
                    "Domínio": domain,
                    "Estado": "Separador não localizado",
                    "Nº indicadores": 0,
                    "Observação": "Verificar alteração da estrutura HTML do site.",
                    "Fonte URL": HTML_URL,
                }
            )
    domain_rows.sort(key=lambda row: DOMAINS.index(row["Domínio"]))

    # Tabela de estrutura programática — primeira tabela fora dos separadores.
    tables = soup.find_all("table")
    structure: list[dict[str, Any]] = []
    if tables:
        rows = tables[0].find_all("tr")
        if rows:
            headers = [clean_text(c.get_text(" ", strip=True)) for c in rows[0].find_all(["th", "td"])]
            for row in rows[1:]:
                cells = [clean_text(c.get_text(" ", strip=True)) for c in row.find_all(["th", "td"])]
                if len(cells) != len(headers):
                    continue
                item: dict[str, Any] = {}
                for header, value in zip(headers, cells):
                    item[header] = parse_web_number(value) if header != "Designation" else value
                structure.append(item)

    main_actions: list[dict[str, Any]] = []
    actions_heading = next(
        (h for h in soup.find_all("h2") if clean_text(h.get_text()) == "Main Actions"),
        None,
    )
    if actions_heading:
        section = actions_heading.find_parent("section")
        for card_header in section.select("div.bg-primary-800") if section else []:
            labels = [
                clean_text(p.get_text(" ", strip=True))
                for p in card_header.find_all("p", recursive=False)
            ]
            if not labels:
                continue
            action_list = card_header.find_next_sibling("ul")
            if action_list is None:
                continue
            for position, item in enumerate(action_list.find_all("li", recursive=False), start=1):
                paragraphs = item.find_all("p")
                text = clean_text(paragraphs[-1].get_text(" ", strip=True) if paragraphs else item.get_text(" ", strip=True))
                main_actions.append(
                    {
                        "Domínio": labels[0],
                        "Subdomínio": labels[1] if len(labels) > 1 else "",
                        "Nº": position,
                        "Resultado/Ação 2025": text,
                        "Período": "Anual 2025",
                        "Fonte URL": HTML_URL,
                    }
                )

    summary: list[dict[str, Any]] = []
    summary_heading = next(
        (h for h in soup.find_all("h2") if clean_text(h.get_text()) == "Implementation Summary"),
        None,
    )
    if summary_heading:
        section = summary_heading.find_parent("section")
        for card in section.select("div.rounded-xl") if section else []:
            spans = [clean_text(s.get_text(" ", strip=True)) for s in card.find_all("span")]
            title = spans[0] if spans else ""
            value = next((parse_web_number(s) for s in spans[1:] if parse_web_number(s) is not None), None)
            paragraphs = [clean_text(p.get_text(" ", strip=True)) for p in card.find_all("p")]
            if title and value is not None:
                summary.append(
                    {
                        "KPI": title,
                        "Valor": value,
                        "Unidade": "N.º",
                        "Descrição": paragraphs[-1] if paragraphs else "",
                        "Fonte URL": HTML_URL,
                    }
                )
        summary_text = clean_text(section.get_text(" ", strip=True) if section else "")
        indicator_bands = [
            ("Indicadores >100%", r"(\d[\d ]*)\s+Indicators recorded execution levels above 100%"),
            ("Indicadores 50–100%", r"(\d[\d ]*)\s+Indicators recorded execution levels between 50% and 100%"),
            ("Indicadores <50%", r"(\d[\d ]*)\s+Indicators recorded execution levels below 50%"),
            ("Indicadores sem execução", r"(\d[\d ]*)\s+Indicators recorded no execution in the period"),
        ]
        for kpi, pattern in indicator_bands:
            match = re.search(pattern, summary_text)
            if match:
                summary.append(
                    {
                        "KPI": kpi,
                        "Valor": parse_web_number(match.group(1)),
                        "Unidade": "N.º",
                        "Descrição": "Distribuição dos 398 indicadores monitorizados.",
                        "Fonte URL": HTML_URL,
                    }
                )
        priority_match = re.search(r"Priorities\s+(\d+)", summary_text)
        if priority_match:
            total_priorities = next(
                (
                    row.get("Priorities")
                    for row in structure
                    if row.get("Designation") == "Total"
                ),
                None,
            )
            planned = total_priorities or 0
            achieved = parse_web_number(priority_match.group(1)) or 0
            summary.append(
                {
                    "KPI": "Prioridades com implementação reportada",
                    "Valor": achieved,
                    "Unidade": f"N.º de {planned} planeadas",
                    "Descrição": f"{(achieved / planned * 100):.2f}% das prioridades planeadas." if planned else "",
                    "Fonte URL": HTML_URL,
                }
            )

    period_match = re.search(r"Anual\s+(\d{4})", soup.get_text(" ", strip=True))
    period = f"Anual {period_match.group(1)}" if period_match else "Não identificado"

    return {
        "soup": soup,
        "indicators": indicators,
        "domains": domain_rows,
        "structure": structure,
        "actions": main_actions,
        "summary": summary,
        "period": period,
    }


def add_meta(
    rows: list[dict[str, Any]],
    policy: str,
    group: str,
    indicator: str,
    unit: str,
    base: int | float | None,
    target: int | float | None,
    long_term: int | float | None,
    page: int,
    notes: str = "",
    raw_base: str | None = None,
    raw_target: str | None = None,
    raw_long_term: str | None = None,
) -> None:
    index = len(rows) + 1
    rows.append(
        {
            "ID_Meta": f"PDN-META-{index:03d}",
            "Eixo/Política": policy,
            "Domínio/Grupo": group,
            "Indicador PDN": indicator,
            "Unidade": unit,
            "Valor_2022_Base": base,
            "Valor_2022_Original": raw_base if raw_base is not None else format_pt_number(base, unit),
            "Meta_2027": target,
            "Meta_2027_Original": raw_target if raw_target is not None else format_pt_number(target, unit),
            "Meta_2050": long_term,
            "Meta_2050_Original": raw_long_term if raw_long_term is not None else format_pt_number(long_term, unit),
            "Página PDF": page,
            "Notas": notes,
        }
    )


def build_pdn_targets() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    policy = "Modernização do Estado"
    add_meta(rows, policy, "Administração Pública", "% de funcionários públicos que beneficiaram de formação nos últimos 5 anos", "%", 11.2, 36.2, None, 30)
    add_meta(rows, policy, "Administração Pública", "% de capacidade de atendimento ao cidadão da rede SIAC", "%", 13.5, 19.3, None, 30)
    add_meta(rows, policy, "Reforma do Estado", "Classificação no indicador Direitos Fundamentais — Projecto de Justiça Mundial", "Índice (0–1)", 0.38, 0.40, 0.50, 30)
    add_meta(rows, policy, "Reforma do Estado", "Classificação no indicador Governo Aberto — Projecto de Justiça Mundial", "Índice (0–1)", 0.38, 0.41, 0.51, 30)
    add_meta(rows, policy, "Reforma do Estado", "Classificação no indicador Democraticidade Eleitoral — Índice Ibrahim", "%", 27.7, 33.7, 71.8, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Estado de Direito — Banco Mundial", "%", 17.3, 22.0, 54.8, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Estabilidade Política — Banco Mundial", "%", 20.8, 25.8, 59.2, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Eficiência do Governo — Banco Mundial", "%", 13.0, 17.5, 55.2, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Qualidade Regulatória — Banco Mundial", "%", 27.4, 33.5, 72.4, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Voz e Participação — Banco Mundial", "%", 24.7, 29.5, 60.6, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Controlo da Corrupção — Banco Mundial", "%", 27.9, 35.3, 63.3, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Capacidade de Gestão — Bertelsmann", "Pontos (0–10)", 4.3, 4.7, 6.5, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Eficiência de Recursos — BTI Project", "Pontos (0–10)", 4.0, 4.3, 6.5, 30)
    add_meta(rows, policy, "Reforma do Estado", "Pontuação no indicador Serviços Públicos Online — UNDESA", "Índice (0–1)", 0.47, 0.49, 0.77, 30)

    policy = "Ordenamento do Território"
    add_meta(rows, policy, "Habitação e território", "Défice habitacional", "% de agregados familiares", 32.0, 31.4, 23.0, 48)
    add_meta(rows, policy, "Habitação e território", "% da população urbana a viver em assentamentos informais", "%", 49.0, 44.5, 33.0, 48)
    add_meta(rows, policy, "Habitação e território", "% de títulos de propriedade para uso habitacional", "% do total de propriedades", 10.0, 35.0, 77.0, 48)
    add_meta(rows, policy, "Habitação e território", "Volume de investimento na reabilitação e manutenção da rede principal", "Biliões de Kz, cumulativo", None, 1.0, 10.2, 48)
    add_meta(rows, policy, "Habitação e território", "Volume de investimento na reabilitação e manutenção da rede complementar", "Biliões de Kz, cumulativo", None, 1.2, 12.2, 48)

    policy = "Educação, Juventude, Emprego e Inovação"
    add_meta(rows, policy, "Educação", "Anos de escolaridade ajustados à aprendizagem", "Anos", 4.2, 4.6, 6.3, 55)
    add_meta(rows, policy, "Educação", "Desempenho em Testes Harmonizados — componente Educação do Índice de Capital Humano", "Pontuação", 326, 337, 380, 55, notes="2022: 4.º quartil; 2050: 3.º quartil.")
    add_meta(rows, policy, "Educação", "Taxa de alfabetização (>=15 anos)", "%", 76.0, 78.0, 90.0, 55)
    add_meta(rows, policy, "Educação", "Taxa líquida de escolarização da 2.ª etapa do pré-escolar (3–5 anos)", "%", 19.0, 23.0, 46.0, 55)
    add_meta(rows, policy, "Educação", "Taxa líquida de escolarização na classe de iniciação (5 anos)", "%", 48.0, 58.0, 90.0, 55)
    add_meta(rows, policy, "Educação", "Taxa líquida de escolarização no ensino primário (1.ª–6.ª classes)", "%", 64.0, 70.0, 90.0, 55)
    add_meta(rows, policy, "Educação", "Taxa líquida de escolarização no 1.º ciclo do ensino secundário (7.ª–9.ª classes)", "%", 23.0, 27.0, 75.0, 55)
    add_meta(rows, policy, "Educação", "% da população entre 15–18 anos a frequentar o 2.º ciclo do ensino secundário", "%", 25.0, 29.0, 50.0, 55)

    policy = "Saúde"
    add_meta(rows, policy, "Saúde", "Esperança média de vida", "Anos", 62, 63, 68, 74)
    add_meta(rows, policy, "Saúde", "Esperança de vida saudável (HALE)", "Anos", 56, 57, 60, 74)
    add_meta(rows, policy, "Saúde", "Taxa de fecundidade", "Filhos por mulher", 5.4, 4.6, 3.2, 74)
    add_meta(rows, policy, "Saúde", "Taxa de mortalidade de menores de 5 anos", "Por 1.000 nados-vivos", 69, 51, 17, 74)
    add_meta(rows, policy, "Saúde", "Taxa de mortalidade de menores de 1 ano", "Por 1.000 nados-vivos", 47, 37, 14, 74)
    add_meta(rows, policy, "Saúde", "Taxa de mortalidade materna", "Por 100.000 nascimentos", 199, 165, 70, 74)
    add_meta(rows, policy, "Saúde", "Gastos correntes com a saúde", "% do PIB", 3.0, 4.0, 7.0, 74)
    add_meta(rows, policy, "Saúde", "Número de médicos por 10.000 habitantes", "Por 10.000 habitantes", 1.8, 2.8, 10.0, 74)
    add_meta(rows, policy, "Saúde", "Número de enfermeiros e parteiras por 10.000 habitantes", "Por 10.000 habitantes", 14.1, 20.6, 22.0, 74)
    add_meta(rows, policy, "Saúde", "Número de hospitais por 100.000 habitantes", "Por 100.000 habitantes", 0.7, 0.8, 0.9, 74)
    add_meta(rows, policy, "Saúde", "Número de centros médicos por 100.000 habitantes", "Por 100.000 habitantes", 2.3, 2.5, 3.5, 74)

    policy = "Cultura"
    add_meta(rows, policy, "Cultura", "% de exportações de bens culturais e criativos", "% do total das exportações", 0.010, 0.016, 0.110, 82, raw_base="0,010%", raw_target="0,016%", raw_long_term="0,110%")
    add_meta(rows, policy, "Cultura", "% da população entre 15–64 anos empregada no sector cultural e criativo", "%", None, 0.3, 2.5, 82)

    policy = "Desporto"
    add_meta(rows, policy, "Desporto", "Classificação no Ranking Mundial de Desporto de Alta Competição", "Posição no ranking", 123, 118, 100, 87)
    add_meta(rows, policy, "Desporto", "% da população envolvida na prática de desporto de recreação", "%", 36.0, 37.8, 47.0, 87)
    add_meta(rows, policy, "Desporto", "Número de medalhas olímpicas ganhas", "N.º", 0, 0, 7, 87)
    add_meta(rows, policy, "Desporto", "Número de medalhas de Jogos Paralímpicos ganhas", "N.º, cumulativo", 8, 9, 16, 87)
    add_meta(rows, policy, "Desporto", "Classificação nos Jogos Africanos", "Posição no ranking", 16, 13, 5, 87)
    add_meta(rows, policy, "Desporto", "Classificação no ranking da SADC (Região 5)", "Posição no ranking", 2, 2, 1, 87)
    add_meta(rows, policy, "Desporto", "Classificação no ranking da CPLP", "Posição no ranking", 3, 3, 3, 87)
    add_meta(rows, policy, "Desporto", "Número de monitores desportivos", "N.º, anual", 7067, 8800, 15000, 87, raw_base="7.067", raw_target="8.800", raw_long_term="15.000")
    add_meta(rows, policy, "Desporto", "Número de atletas federados", "N.º, anual", 63091, 69500, 110000, 87, raw_base="63.091", raw_target="69.500", raw_long_term="110.000")
    add_meta(rows, policy, "Desporto", "% de atletas federados do sexo feminino", "%", 20.0, 28.0, 50.0, 87)

    policy = "População e Promoção das Comunidades Vulneráveis"
    add_meta(rows, policy, "Protecção Social", "% da população que vive abaixo do limiar de pobreza (<2,15 USD/dia)", "%", 31.0, 28.0, 18.0, 92)
    add_meta(rows, policy, "Protecção Social", "Volume de investimento em assistência social", "Biliões de Kz, anual", 0.16, 0.33, 1.54, 92)
    add_meta(rows, policy, "Protecção Social", "Número de segurados registados na protecção social obrigatória", "Milhões, anual", 2.5, 4.3, 13.6, 92)
    add_meta(rows, policy, "Protecção Social", "% de excedente do sistema de segurança social", "%", 0.25, 0.26, 0.33, 92, raw_base="0,25%", raw_target="0,26%", raw_long_term="0,33%")
    add_meta(rows, policy, "Idosos", "Número de instalações públicas de cuidados a terceira idade construídas e reabilitadas", "N.º, cumulativo", 11, 12, 13, 92)

    policy = "Energética"
    add_meta(rows, policy, "Energia", "Taxa de electrificação (on-grid)", "%", 43.0, 49.0, 72.0, 104)
    add_meta(rows, policy, "Energia", "Capacidade instalada de energia eléctrica", "GW", 6, 10, 33, 104)
    add_meta(rows, policy, "Energia", "Produção de energias renováveis", "% da capacidade instalada", 64.0, 73.0, 94.0, 104)
    policy = "Energética"
    add_meta(rows, policy, "Petróleo e Gás", "Contribuição do petróleo e do gás para o PIB", "Biliões de Kz, anual", 15.4, 13.8, 5.5, 104)
    add_meta(rows, policy, "Petróleo e Gás", "Contribuição do petróleo e do gás para o PIB", "% do PIB", 29.0, 22.0, 4.0, 104)
    add_meta(rows, policy, "Petróleo e Gás", "Contribuição do sector do petróleo para o PIB", "Biliões de Kz, anual", 13.2, 11.4, 4.6, 104)
    add_meta(rows, policy, "Petróleo e Gás", "Contribuição do sector do petróleo para o PIB", "% do PIB", 27.0, 21.0, 3.0, 104)
    add_meta(rows, policy, "Petróleo e Gás", "Contribuição do sector do gás para o PIB", "Biliões de Kz, anual", 0.9, 1.4, 0.9, 104)
    add_meta(rows, policy, "Petróleo e Gás", "Contribuição do sector do gás para o PIB", "% do PIB", 2.0, 1.0, 1.0, 104)

    policy = "Comunicações e Aceleração Digital"
    add_meta(rows, policy, "Comunicações", "Contribuição do sector das comunicações e tecnologias de informação para o PIB", "Biliões de Kz, anual", 0.31, 0.50, 4.30, 113)
    add_meta(rows, policy, "Comunicações", "Contribuição do sector das comunicações e tecnologias de informação para o PIB", "% do PIB", 0.6, 0.8, 3.2, 113)
    add_meta(rows, policy, "Comunicações", "Cobertura de rede 3G", "% da população", 76.0, 93.0, 100.0, 113)
    add_meta(rows, policy, "Comunicações", "Cobertura de rede 4G", "% da população", 23.0, 32.0, 100.0, 113)
    add_meta(rows, policy, "Comunicações", "Cobertura de rede 5G", "% da população", 0.0, 21.0, 100.0, 113)

    policy = "Transportes e Logística"
    add_meta(rows, policy, "Transportes", "Contribuição do sector dos transportes para o PIB", "Biliões de Kz, anual", 0.85, 1.11, 2.0, 117)
    add_meta(rows, policy, "Transportes", "Subsector da aviação civil — número de passageiros", "Milhões, anual", 2.2, 4.8, 10.1, 117)
    add_meta(rows, policy, "Transportes", "Subsector ferroviário — toneladas transportadas", "Milhões de toneladas, anual", 0.5, 3.3, 20.9, 117)
    add_meta(rows, policy, "Transportes", "Subsector marítimo e portuário — toneladas de carga processadas nos portos", "Milhões de toneladas, anual", 17.4, 29.8, 48.4, 117)

    policy = "Águas e Saneamento"
    add_meta(rows, policy, "Água", "Abastecimento de água para responder às necessidades da população e economia angolanas", "hm³", 1344, 2469, 13139, 125, raw_base="1.344", raw_target="2.469", raw_long_term="13.139")
    add_meta(rows, policy, "Água", "Capacidade de armazenamento de água", "m³ per capita", 317, 350, 400, 125)
    add_meta(rows, policy, "Água", "% da população que utiliza serviços básicos de água potável", "%", 57.0, 61.0, 89.0, 125)
    add_meta(rows, policy, "Água", "% da população que utiliza serviços básicos de saneamento", "%", 52.0, 55.0, 66.0, 125)
    add_meta(rows, policy, "Água", "Densidade das estações hidrométricas", "km²/estação", 21500, 13508, 3750, 125, raw_base="21.500", raw_target="13.508", raw_long_term="3.750")

    policy = "Sustentabilidade Ambiental"
    add_meta(rows, policy, "Ambiente", "Pontuação no Índice de Desempenho Ambiental (EPI)", "Pontuação", 30.5, 33.6, 52.6, 129)
    add_meta(rows, policy, "Ambiente", "Emissões de gases de efeito estufa", "KtCO₂e", 99992.2, 84694.3, None, 129, raw_base="99.992,2", raw_target="84.694,3", raw_long_term="-")
    add_meta(rows, policy, "Ambiente", "Pontuação na categoria Qualidade do Ar do EPI", "Pontuação", 23.1, 25.7, 42.2, 129)
    add_meta(rows, policy, "Ambiente", "Pontuação na categoria Biodiversidade e Habitat do EPI", "Pontuação", 30.1, 34.6, 65.3, 129)
    add_meta(rows, policy, "Ambiente", "Pontuação no Índice Mundial de Risco — Bündnis Entwicklung Hilft", "Pontuação", 11.0, 10.2, 7.0, 129)

    policy = "Apoio à Produção, Diversificação das Exportações e Substituição das Importações"
    add_meta(rows, policy, "Agricultura e Pecuária", "Contribuição da agricultura para o PIB", "Biliões de Kz, anual", 5.2, 7.5, 18.4, 138, raw_target="7,5")
    add_meta(rows, policy, "Agricultura e Pecuária", "Contribuição da agricultura para o PIB", "% do PIB", 9.7, 12.1, 14.1, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Contribuição da pecuária para o PIB", "Biliões de Kz, anual", 0.4, 0.6, 2.2, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Contribuição da pecuária para o PIB", "% do PIB", 0.7, 0.9, 1.7, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Número de explorações agrícolas familiares", "Milhões, anual", 3.10, 3.41, 4.70, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Área média de terreno agrícola por exploração familiar", "Hectares", 1.81, 1.93, 2.30, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Número de explorações agrícolas comerciais", "Milhares, anual", 5.88, 7.25, 18.95, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Área média de terreno agrícola por exploração comercial", "Hectares", 88.2, 91.8, 110.2, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Área total de terra agrícola para agricultura familiar", "Milhões de hectares", 5.61, 6.58, 10.94, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Área total de terra agrícola para agricultura empresarial", "Milhões de hectares", 0.50, 0.62, 2.10, 138)
    add_meta(rows, policy, "Agricultura e Pecuária", "Toneladas de produção de carne", "Milhares de toneladas", 320, 788, 2415, 138, raw_long_term="2.415")

    policy = "Estabilidade e Crescimento Económico"
    add_meta(rows, policy, "Formalização da economia", "Taxa de formalização", "% de todos os empregos", 22.0, 31.0, 55.0, 163)
    add_meta(rows, policy, "Formalização da economia", "Número total de trabalhadores registados no INSS", "Milhões, anual", 2.5, 4.3, 13.6, 163)
    add_meta(rows, policy, "Apoio ao empresariado nacional e financiamento da economia", "Volume total de investimento", "Mil milhões de USD, cumulativo", None, 126, 948, 163)
    add_meta(rows, policy, "Apoio ao empresariado nacional e financiamento da economia", "Stock de crédito ao sector privado", "% do PIB não-petrolífero", 10.9, 12.5, 36.0, 163)
    add_meta(rows, policy, "Apoio ao empresariado nacional e financiamento da economia", "IDE não-petrolífero", "% do PIB não-petrolífero", None, 9.0, 12.0, 163, raw_base="<1%", raw_target="9%", raw_long_term="12%")
    add_meta(rows, policy, "Ambiente de negócios", "Classificação no ranking do Índice de Competitividade Global — Fórum Económico Mundial", "Posição no ranking", 136, 112, 80, 163, raw_long_term="Top 80")
    add_meta(rows, policy, "Ambiente de negócios", "Investimento privado per capita", "Milhões de USD, anual", 280, 338, 630, 163)

    policy = "Defesa e Segurança"
    add_meta(rows, policy, "Defesa", "Classificação no Índice de Paz Global — Instituto para Economia e Paz", "Posição no ranking", 78, 72, 60, 179, raw_long_term="Top 60")
    add_meta(rows, policy, "Defesa", "Classificação no Índice de Segurança e Protecção — Índice Ibrahim", "Posição no ranking", 28, 24, 20, 179, raw_long_term="Top 20")
    add_meta(rows, policy, "Defesa", "Despesa militar", "% do PIB", 1.4, 1.0, 1.0, 179)
    add_meta(rows, policy, "Defesa", "Número de militares por 1.000 habitantes", "Por 1.000 habitantes", 3.5, 3.1, 1.7, 179)
    add_meta(rows, policy, "Interior e Segurança Pública", "Número de crimes por 100.000 habitantes", "Por 100.000 habitantes", 207, 197, 165, 179)
    add_meta(rows, policy, "Interior e Segurança Pública", "Taxa de ocupação do sistema penitenciário", "%", 113.0, 108.0, 100.0, 179)
    add_meta(rows, policy, "Interior e Segurança Pública", "Despesa com ordem e segurança", "% do PIB", 1.4, 1.3, 1.2, 179)
    add_meta(rows, policy, "Interior e Segurança Pública", "Rácio polícia/cidadão", "Por 100.000 habitantes", 404, 367, 336, 179)

    policy = "Política Externa"
    add_meta(rows, policy, "Política Externa", "Volume de investimento directo estrangeiro, em especial no sector do agronegócio", "Mil milhões de USD, anual", 6.0, 14.0, 28.0, 193)
    add_meta(rows, policy, "Política Externa", "Volume de exportações não-petrolíferas", "Mil milhões de USD, anual", 4.0, 7.3, 64.0, 193, raw_target="7,3")
    add_meta(rows, policy, "Política Externa", "Número de turistas internacionais recebidos por Angola", "Milhões, anual", 0.12, 0.16, 2.0, 193)
    add_meta(rows, policy, "Política Externa", "Cumprimento dos limiares da primeira avaliação para graduação de Angola para País de Rendimento Médio", "Marco", None, None, None, 193, raw_base="-", raw_target="X", raw_long_term="-")
    add_meta(rows, policy, "Política Externa", "Adesão efectiva à Zona de Livre Comércio da SADC, CEEAC, Tripartida e Continental concluída", "Marco", None, None, None, 193, raw_base="-", raw_target="X", raw_long_term="-")
    add_meta(rows, policy, "Política Externa", "Número de novos quadros angolanos que trabalham em organizações internacionais", "N.º, anual", None, 16, 100, 193, raw_base="-")

    return rows


def build_comparison(indicators: list[dict[str, Any]]) -> list[dict[str, Any]]:
    mapping = {
        "Novas Camas em Hospitais (Cumulativo)": {
            "tipo": "Não comparável diretamente",
            "referencia": "Número de hospitais por 100.000 habitantes",
            "paginas": "74",
            "observacao": "Camas hospitalares não são o mesmo que número/densidade de hospitais; o PDN não define meta numérica de camas.",
        },
        "Taxa de cobertura da 4.ª consulta pré-natal": {
            "tipo": "Correspondência temática sem meta numérica",
            "referencia": "Prioridade 18.2.3 — expansão do acesso a serviços de saúde sexual e reprodutiva",
            "paginas": "76–77",
            "observacao": "O PDN menciona a melhoria da oferta de consultas pré-natais, mas não publica meta quantitativa equivalente no quadro da Política de Saúde.",
        },
        "Taxa de imunização BCG": {
            "tipo": "Correspondência temática sem meta numérica",
            "referencia": "Prioridade 18.2.1 — reforço da imunização a nível nacional",
            "paginas": "76",
            "observacao": "O PDN prevê fortalecer vacinação de rotina, mas não publica uma meta numérica de BCG.",
        },
        "Taxa de vacinação contra o sarampo": {
            "tipo": "Correspondência temática sem meta numérica",
            "referencia": "Prioridade 18.2.1 — reforço da imunização a nível nacional",
            "paginas": "76",
            "observacao": "O PDN prevê fortalecer vacinação de rotina, mas não publica uma meta numérica de sarampo.",
        },
        "Taxa de incidência da malária": {
            "tipo": "Correspondência temática sem meta numérica",
            "referencia": "Objectivo 18.3 / Redução da incidência da malária",
            "paginas": "77",
            "observacao": "A redução da incidência é explicitamente prevista, mas sem valor-base/meta numérica no quadro de metas.",
        },
        "Taxa de sucesso do tratamento para a tuberculose": {
            "tipo": "Correspondência temática sem meta numérica",
            "referencia": "Objectivo 18.3 / Redução da incidência da tuberculose",
            "paginas": "77",
            "observacao": "O PDN prevê diagnóstico e tratamento completos, mas não publica meta quantitativa de sucesso terapêutico.",
        },
        "Partos institucionais realizados por profissionais qualificados": {
            "tipo": "Correspondência temática sem meta numérica",
            "referencia": "Prioridade 18.2.3 — serviços de saúde sexual e reprodutiva",
            "paginas": "76–77",
            "observacao": "Partos institucionais são mencionados como intervenção, mas não existe meta numérica equivalente no quadro da Política de Saúde.",
        },
    }
    result: list[dict[str, Any]] = []
    for item in indicators:
        info = mapping.get(
            item["Indicador"],
            {
                "tipo": "Sem correspondência",
                "referencia": "",
                "paginas": "",
                "observacao": "Não foi localizado indicador PDN equivalente.",
            },
        )
        result.append(
            {
                "ID_Implementação": item["ID"],
                "Domínio": item["Domínio"],
                "Indicador de execução": item["Indicador"],
                "Unidade execução": item["Unidade"],
                "2022_Base": item.get("2022_Base"),
                "2023": item.get("2023"),
                "2024": item.get("2024"),
                "2025": item.get("2025"),
                "Meta PDN 2027": None,
                "Meta PDN 2027 original": "Sem valor numérico equivalente publicado",
                "Tipo de correspondência": info["tipo"],
                "Referência/metas PDN": info["referencia"],
                "Páginas PDN": info["paginas"],
                "Diferença 2025 vs meta": None,
                "Progresso 2025 vs meta": None,
                "Observação": info["observacao"],
            }
        )
    return result


def write_csv(path: Path, rows: list[dict[str, Any]], columns: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if columns is None:
        columns = list(rows[0]) if rows else []
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({column: "" if row.get(column) is None else row.get(column) for column in columns})


NAVY = "153A5B"
BLUE = "1F6E8C"
GOLD = "D8A126"
PALE_BLUE = "EAF3F8"
PALE_GOLD = "FFF4D6"
GREY = "F2F4F6"
WHITE = "FFFFFF"
THIN_GREY = Side(style="thin", color="D9E1E8")


def style_header(ws: Any, row: int = 1) -> None:
    for cell in ws[row]:
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="medium", color=GOLD))
    ws.row_dimensions[row].height = 34


def write_table_sheet(
    wb: Workbook,
    name: str,
    rows: list[dict[str, Any]],
    columns: list[str],
    widths: dict[str, int] | None = None,
) -> Any:
    ws = wb.create_sheet(name[:31])
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])
    style_header(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"
    if ws.max_row > 1:
        table_name = f"T_{slug(name)[:20] or 'dados'}"
        table = Table(displayName=table_name, ref=f"A1:{get_column_letter(len(columns))}{ws.max_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        )
        ws.add_table(table)
    for column_index, column in enumerate(columns, start=1):
        letter = get_column_letter(column_index)
        width = (widths or {}).get(column)
        if width is None:
            maximum = max(
                [len(str(column))]
                + [len(str(ws.cell(row=r, column=column_index).value or "")) for r in range(2, min(ws.max_row, 250) + 1)]
            )
            width = min(max(maximum + 2, 12), 52)
        ws.column_dimensions[letter].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'
    return ws


def write_readme_sheet(wb: Workbook, name: str, rows: list[tuple[str, str]]) -> Any:
    ws = wb.create_sheet(name[:31])
    ws.append(["Campo", "Valor"])
    for key, value in rows:
        ws.append([key, value])
    style_header(ws)
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 31
    ws.column_dimensions["B"].width = 115
    for row in ws.iter_rows(min_row=2):
        row[0].font = Font(bold=True, color=NAVY)
        row[0].fill = PatternFill("solid", fgColor=PALE_BLUE)
        row[0].alignment = Alignment(vertical="top", wrap_text=True)
        row[1].alignment = Alignment(vertical="top", wrap_text=True)
        if isinstance(row[1].value, str) and row[1].value.startswith("http"):
            row[1].hyperlink = row[1].value
            row[1].style = "Hyperlink"
    ws.auto_filter.ref = f"A1:B{ws.max_row}"
    return ws


def finish_workbook(wb: Workbook, path: Path) -> None:
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.oddFooter.center.text = "PDN Angola 2023-2027 | fonte oficial preservada na pasta"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def make_execution_workbook(
    parsed: dict[str, Any],
    pdn_info: dict[str, Any],
    manifest: list[dict[str, str]],
) -> Path:
    wb = Workbook()
    readme_rows = [
        ("Título", "Execução do PDN Angola 2023-2027 — indicadores e acções reportadas"),
        ("Período da página", parsed["period"]),
        ("Data da extração", date.today().isoformat()),
        ("Fonte de execução", HTML_URL),
        ("Fonte do plano", PDN_URL_REQUESTED),
        ("Ficheiro PDF usado", str(PDF_FILE.relative_to(ROOT))),
        ("Páginas PDF", str(pdn_info.get("Páginas PDF", ""))),
        ("Cobertura", "Todos os sete separadores de domínio foram verificados no HTML."),
        ("Limitação principal", "Só Saúde contém tabela de indicadores no HTML publicado; os outros seis separadores estão vazios."),
        ("Tratamento de ausência", "Sem dados publicados não foi convertido em zero."),
        ("Ano base", "2022* conforme a nota da página do MINPLAN."),
        ("Conteúdo", "Indicadores extraídos, acções de 2025, estrutura programática e resumo de implementação."),
    ]
    write_readme_sheet(wb, "Leia-me", readme_rows)

    indicator_columns = [
        "ID", "Domínio", "Indicador", "Unidade", "2022_Base", "2023", "2024", "2025",
        "2022_Base_Original", "2023_Original", "2024_Original", "2025_Original",
        "Fonte", "Fonte URL", "Estado", "Observação",
    ]
    write_table_sheet(
        wb, "Indicadores_Execução", parsed["indicators"], indicator_columns,
        {"ID": 25, "Domínio": 28, "Indicador": 54, "Unidade": 24, "Fonte URL": 45, "Observação": 48},
    )
    write_table_sheet(
        wb, "Índice_Domínios", parsed["domains"],
        ["Domínio", "Estado", "Nº indicadores", "Observação", "Fonte URL"],
        {"Domínio": 42, "Estado": 30, "Observação": 70, "Fonte URL": 45},
    )
    write_table_sheet(
        wb, "Ações_2025", parsed["actions"],
        ["Domínio", "Subdomínio", "Nº", "Resultado/Ação 2025", "Período", "Fonte URL"],
        {"Domínio": 30, "Subdomínio": 70, "Resultado/Ação 2025": 110, "Fonte URL": 45},
    )
    write_table_sheet(
        wb, "Estrutura_Programa", parsed["structure"],
        ["Designation", "Pol.", "Programmes", "Objectives", "Priorities", "Monitored Indicators"],
        {"Designation": 115},
    )
    write_table_sheet(
        wb, "Resumo_Execução", parsed["summary"],
        ["KPI", "Valor", "Unidade", "Descrição", "Fonte URL"],
        {"KPI": 38, "Unidade": 24, "Descrição": 75, "Fonte URL": 45},
    )
    write_table_sheet(
        wb, "Manifesto_Fontes", manifest,
        ["Fonte", "URL solicitada", "URL usada", "Ficheiro local", "Estado"],
        {"Fonte": 42, "URL solicitada": 65, "URL usada": 65, "Ficheiro local": 55, "Estado": 90},
    )
    path = ROOT / "01_Execucao_PDN_2025.xlsx"
    finish_workbook(wb, path)
    return path


def make_comparison_workbook(
    parsed: dict[str, Any],
    targets: list[dict[str, Any]],
    comparison: list[dict[str, Any]],
    pdn_info: dict[str, Any],
) -> Path:
    wb = Workbook()
    readme_rows = [
        ("Título", "Execução do PDN vs metas do PDN Angola 2023-2027"),
        ("Data da extração", date.today().isoformat()),
        ("Fonte de execução", HTML_URL),
        ("Fonte das metas", PDN_URL_REQUESTED),
        ("Metas catalogadas", f"{len(targets)} indicadores/metas das páginas 'Metas da Política' do PDN."),
        ("Indicadores comparados", f"{len(comparison)} indicadores publicados no painel de execução do MINPLAN."),
        ("Resultado de comparabilidade", "Os sete indicadores publicados na execução não têm valor-meta quantitativo equivalente no quadro de metas da Política de Saúde."),
        ("Regra", "Não foi inferida uma meta a partir de um indicador diferente. Correspondências temáticas são identificadas, mas ficam sem cálculo de progresso."),
        ("Implicação", "As colunas de diferença e progresso ficam vazias quando o PDN não publica meta numérica equivalente."),
        ("Páginas de saúde", "A Política de Saúde e as suas metas principais estão na página 74; prioridades de imunização, pré-natal, malária e tuberculose nas páginas 76–77."),
        ("Páginas PDF", str(pdn_info.get("Páginas PDF", ""))),
    ]
    write_readme_sheet(wb, "Leia-me", readme_rows)

    comparison_ws = write_table_sheet(
        wb, "Comparação", comparison,
        [
            "ID_Implementação", "Domínio", "Indicador de execução", "Unidade execução",
            "2022_Base", "2023", "2024", "2025", "Meta PDN 2027",
            "Meta PDN 2027 original", "Tipo de correspondência", "Referência/metas PDN",
            "Páginas PDN", "Diferença 2025 vs meta", "Progresso 2025 vs meta", "Observação",
        ],
        {"ID_Implementação": 30, "Domínio": 27, "Indicador de execução": 55, "Unidade execução": 24, "Tipo de correspondência": 38, "Referência/metas PDN": 68, "Observação": 75},
    )
    comparison_ws.conditional_formatting.add(
        f"K2:K{max(2, comparison_ws.max_row)}",
        CellIsRule(operator="equal", formula=['"Não comparável diretamente"'], fill=PatternFill("solid", fgColor=PALE_GOLD)),
    )
    write_table_sheet(
        wb, "Metas_PDN_Originais", targets,
        [
            "ID_Meta", "Eixo/Política", "Domínio/Grupo", "Indicador PDN", "Unidade",
            "Valor_2022_Base", "Valor_2022_Original", "Meta_2027", "Meta_2027_Original",
            "Meta_2050", "Meta_2050_Original", "Página PDF", "Notas",
        ],
        {"Eixo/Política": 60, "Domínio/Grupo": 42, "Indicador PDN": 78, "Unidade": 32, "Notas": 55},
    )
    health_targets = [row for row in targets if row["Eixo/Política"] == "Saúde"]
    write_table_sheet(
        wb, "Metas_Saúde_PDN", health_targets,
        [
            "ID_Meta", "Eixo/Política", "Domínio/Grupo", "Indicador PDN", "Unidade",
            "Valor_2022_Base", "Valor_2022_Original", "Meta_2027", "Meta_2027_Original",
            "Meta_2050", "Meta_2050_Original", "Página PDF", "Notas",
        ],
        {"Indicador PDN": 78, "Unidade": 32, "Notas": 55},
    )
    write_table_sheet(
        wb, "Cobertura_Domínios", parsed["domains"],
        ["Domínio", "Estado", "Nº indicadores", "Observação", "Fonte URL"],
        {"Domínio": 42, "Estado": 30, "Observação": 70, "Fonte URL": 45},
    )
    path = ROOT / "02_Execucao_vs_Metas_PDN_2027.xlsx"
    finish_workbook(wb, path)
    return path


def write_analysis_markdown(parsed: dict[str, Any], targets: list[dict[str, Any]], pdn_info: dict[str, Any]) -> Path:
    total_structure = next(
        (row for row in parsed["structure"] if row.get("Designation") == "Total"), {}
    )
    populated = [row for row in parsed["domains"] if row["Nº indicadores"]]
    empty = [row for row in parsed["domains"] if not row["Nº indicadores"]]
    lines = [
        "# Análise do PDN 2023-2027 e do painel de execução",
        "",
        f"Data de extração: {date.today().isoformat()}",
        "",
        "## 1. Plano de referência",
        "",
        "O PDN 2023-2027 é o instrumento de médio prazo que implementa a Estratégia de Longo Prazo Angola 2050. A estrutura publicada organiza o plano em sete eixos, 16 políticas, 50 programas, 171 objectivos, 284 prioridades e 398 indicadores monitorizados.",
        "",
        "A metodologia do próprio PDN separa indicadores de impacto, indicadores de resultado e execução orçamental programática. O plano prevê monitorização permanente, relatórios mensais e trimestrais, balanços anuais, avaliação intercalar em 2025 e avaliação final em 2028.",
        "",
        "Os dois motores de desenvolvimento são o desenvolvimento do capital humano e a segurança alimentar. Os filtros transversais incluem receita fiscal, juventude, igualdade de género, emprego, sustentabilidade ambiental, comunidades vulneráveis e ambiente de negócios.",
        "",
        "## 2. O que foi publicado no painel do MINPLAN",
        "",
        f"O painel foi capturado na página `{HTML_URL}` e identifica o período `{parsed['period']}`. O inventário verificou os sete domínios de `Implementation - Key Indicators by Domain`.",
        "",
        f"Há dados numéricos em {len(populated)} domínio(s): {', '.join(row['Domínio'] for row in populated)}. A tabela contém {len(parsed['indicators'])} indicadores, com ano base 2022 e valores para 2023, 2024 e 2025.",
        "",
        f"Os painéis sem tabela no HTML são: {', '.join(row['Domínio'] for row in empty)}. Eles foram mantidos no inventário com estado `Sem dados publicados no HTML`; não foram convertidos em zero.",
        "",
        "A secção `Main Actions` contém resultados narrativos de 2025 para os domínios Social, Fomento da produção nacional, Infra-estruturas e Construção e Obras públicas. Esses resultados foram preservados numa folha separada, sem os misturar com indicadores quantitativos.",
        "",
        "## 3. Comparação com metas",
        "",
        f"Foram catalogadas {len(targets)} linhas de metas das páginas `Metas da Política` do PDF. A Política de Saúde (página 74) define metas para esperança de vida, mortalidade, despesa, profissionais e densidade de unidades; as prioridades de imunização, pré-natal, malária e tuberculose aparecem nas páginas 76–77.",
        "",
        "Os sete indicadores publicados pelo MINPLAN não têm, no quadro de metas da Política de Saúde, um valor quantitativo equivalente que permita calcular automaticamente o progresso para 2027. Por isso, o segundo Excel distingue correspondência temática de comparabilidade numérica e deixa diferença/progresso vazios quando não há meta equivalente.",
        "",
        "## 4. Proveniência e limitação de acesso",
        "",
        f"A URL NEPAD indicada foi preservada como fonte solicitada. Como o ficheiro devolveu erro 403 no ambiente de extração, foi guardado um espelho acessível do documento no ficheiro `{PDF_FILE.relative_to(ROOT)}`. A URL usada e o fallback estão registados em `fontes/00_manifesto_fontes.csv` e nas folhas `Manifesto_Fontes`.",
        "",
        "A ausência de tabelas nos seis domínios não prova ausência de execução; prova apenas que esses dados não estavam presentes no HTML recebido da página consultada. Para completar esses domínios será necessário obter os relatórios/documentos sectoriais que o MINPLAN ainda assinala como em desenvolvimento.",
        "",
        "## Ficheiros entregues",
        "",
        "- `01_Execucao_PDN_2025.xlsx`: execução publicada, cobertura dos domínios, acções 2025, estrutura e resumo.",
        "- `02_Execucao_vs_Metas_PDN_2027.xlsx`: comparação sem inferência indevida, catálogo completo das metas transcritas e metas de saúde em separado.",
        "- `fontes/`: cópias locais das fontes usadas.",
        "- `dados/`: CSVs normalizados e texto extraído do PDF.",
        "- `scripts/compilar_pdn.py`: scraper/compilador reexecutável.",
        "",
    ]
    path = ROOT / "ANALISE_PDN.md"
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--refresh", action="store_true", help="Volta a descarregar as fontes.")
    args = parser.parse_args()

    manifest = ensure_sources(refresh=args.refresh)
    pdn_info = extract_pdn_text()
    parsed = parse_html_source()
    targets = build_pdn_targets()
    comparison = build_comparison(parsed["indicators"])

    indicator_columns = [
        "ID", "Domínio", "Indicador", "Unidade", *YEAR_COLUMNS,
        "2022_Base_Original", "2023_Original", "2024_Original", "2025_Original",
        "Fonte", "Fonte URL", "Estado", "Observação",
    ]
    write_csv(DATA_DIR / "indicadores_execucao.csv", parsed["indicators"], indicator_columns)
    write_csv(DATA_DIR / "dominios_cobertura.csv", parsed["domains"])
    write_csv(DATA_DIR / "acoes_2025.csv", parsed["actions"])
    write_csv(DATA_DIR / "estrutura_programa.csv", parsed["structure"])
    write_csv(DATA_DIR / "resumo_execucao.csv", parsed["summary"])
    write_csv(DATA_DIR / "metas_pdn_2027.csv", targets)
    write_csv(DATA_DIR / "comparacao_execucao_metas.csv", comparison)

    execution_path = make_execution_workbook(parsed, pdn_info, manifest)
    comparison_path = make_comparison_workbook(parsed, targets, comparison, pdn_info)
    analysis_path = write_analysis_markdown(parsed, targets, pdn_info)

    print(f"Indicadores de execução: {len(parsed['indicators'])}")
    print(f"Domínios verificados: {len(parsed['domains'])}")
    print(f"Acções 2025 extraídas: {len(parsed['actions'])}")
    print(f"Metas PDN catalogadas: {len(targets)}")
    print(f"Ficheiro 1: {execution_path}")
    print(f"Ficheiro 2: {comparison_path}")
    print(f"Análise: {analysis_path}")


if __name__ == "__main__":
    main()
