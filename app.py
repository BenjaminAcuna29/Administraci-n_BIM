from __future__ import annotations

import io
import html
import re
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException
import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st


st.set_page_config(
    page_title="Control Presupuestario | Construcción",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

PALETTE = {
    "navy": "#12253F",
    "navy_light": "#203B5D",
    "orange": "#F28C28",
    "gray": "#64748B",
    "light": "#F3F6FA",
    "white": "#FFFFFF",
    "green": "#15966B",
    "red": "#C54141",
}
CATEGORY_COLORS = {
    "Materiales": "#F28C28",
    "Mano de obra": "#203B5D",
    "Maquinaria": "#5582B0",
    "Subcontratos": "#9AA9BA",
    "Sin clasificar": "#C8D1DC",
}

st.markdown(
    """
    <style>
    .stApp { background: #F3F6FA; color: #17283D; }
    [data-testid="stSidebar"] { background: #12253F; }
    [data-testid="stSidebar"] * { color: #F5F7FA; }
    [data-testid="stMetric"] {
        background: #FFFFFF; border: 1px solid #E1E7EF; border-radius: 12px;
        padding: 16px 18px; box-shadow: 0 3px 12px rgba(18,37,63,.05);
    }
    h1, h2, h3 { color: #12253F; }
    .hero { background: linear-gradient(110deg,#12253F,#203B5D); color: white;
        padding: 24px 28px; border-radius: 16px; margin-bottom: 18px; }
    .hero h1 { color: white; margin: 0; font-size: 1.8rem; }
    .hero p { color: #DCE5EF; margin: 7px 0 0; }
    .note { border-left: 4px solid #F28C28; padding: 10px 14px;
        background: white; border-radius: 4px; color: #42536A; }
    </style>
    """,
    unsafe_allow_html=True,
)


def norm(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9%]+", " ", text.lower()).strip()


def cell_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def number(value: Any) -> float | None:
    if value is None or pd.isna(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("$", "").replace(" ", "")
    if not text:
        return None
    text = re.sub(r"[^0-9,.\-]", "", text)
    if not text or text in {".", ",", "-"}:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        tail = text.rsplit(",", 1)[-1]
        text = text.replace(",", "." if len(tail) <= 2 else "")
    try:
        return float(text)
    except ValueError:
        return None


def code_text(value: Any) -> str:
    text = cell_text(value)
    if not text:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\s+", "", text).rstrip(".")


def find_column(headers: list[Any], terms: tuple[str, ...]) -> int | None:
    clean = [norm(item) for item in headers]
    for i, label in enumerate(clean):
        if label and any(term in label for term in terms):
            return i
    return None


def find_unit_column(headers: list[Any]) -> int | None:
    exact = {"u", "un", "und", "unidad", "unidad medida", "unit"}
    for i, label in enumerate(headers):
        if norm(label) in exact:
            return i
    return find_column(headers, ("unidad",))


def find_price_column(headers: list[Any]) -> int | None:
    clean = [norm(item) for item in headers]
    for i, label in enumerate(clean):
        if "precio unitario" in label or "costo unitario" in label:
            return i
    for i, label in enumerate(clean):
        if label == "valor" or label.startswith("valor "):
            return i
    return None


def value_at(row: tuple[Any, ...] | list[Any], index: int | None) -> Any:
    return row[index] if index is not None and index < len(row) else None


def is_item_code(value: Any) -> bool:
    text = code_text(value)
    return bool(text and re.fullmatch(r"\d+(?:\.\d+)*", text))


def is_resource_code(value: Any) -> bool:
    text = code_text(value).upper()
    return bool(re.fullmatch(r"(?:MAT|MO|MAQ|SUB)-\d+", text))


def resource_category(code: str, context: str = "") -> str:
    prefix = code_text(code).upper().split("-", 1)[0]
    mapping = {
        "MAT": "Materiales",
        "MO": "Mano de obra",
        "MAQ": "Maquinaria",
        "SUB": "Subcontratos",
    }
    if prefix in mapping:
        return mapping[prefix]
    title = norm(context)
    if "material" in title:
        return "Materiales"
    if "mano de obra" in title or "personal" in title:
        return "Mano de obra"
    if "maquina" in title or "equipo" in title:
        return "Maquinaria"
    if "subcontrat" in title or "otros" in title:
        return "Subcontratos"
    return "Sin clasificar"


@st.cache_data(show_spinner="Analizando las hojas y los datos del libro...")
def parse_workbook(raw: bytes) -> dict[str, Any]:
    excel = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
    sheets: dict[str, pd.DataFrame] = {
        name: pd.read_excel(
            excel, sheet_name=name, header=None, dtype=object, engine="openpyxl"
        )
        for name in excel.sheet_names
    }
    budget_name, budget_header, budget_cols = detect_budget(sheets)
    if budget_name is None:
        raise ValueError(
            "No se encontró una hoja de presupuesto con columnas de ítem, "
            "designación/actividad, unidad, cantidad y total."
        )

    budget = parse_budget(
        sheets[budget_name], budget_header, budget_cols, budget_name
    )
    apu_sheets = detect_apu_sheets(sheets, budget_name)
    apus, resources, apu_issues = parse_apus(sheets, apu_sheets)
    catalogs = parse_catalogs(sheets, set(apu_sheets) | {budget_name})
    overhead, overhead_source = parse_overhead(sheets, budget_name)

    apu_codes = {item["code"] for item in apus if item["code"]}
    budget_codes = {item["code"] for item in budget if item["code"]}
    for item in budget:
        item["apu_status"] = (
            "Vinculado"
            if item["code"] in apu_codes
            else "Sin APU"
        )
    for resource in resources:
        match = catalog_match(resource, catalogs)
        resource["catalog_status"] = match

    issues: list[dict[str, str]] = []
    for issue in apu_issues:
        issues.append(issue)
    for item in budget:
        if item["code"] and item["code"] not in apu_codes and item["total"] not in (None, 0):
            issues.append(
                {
                    "Severidad": "Revisar",
                    "Tipo": "Partida sin APU",
                    "Código": item["code"],
                    "Detalle": f"{item['description']} no tiene una hoja APU con ese código.",
                }
            )
        if item["total"] not in (None, 0):
            missing = [
                label for label, value in (
                    ("unidad", item["unit"]),
                    ("cantidad", item["quantity"]),
                    ("valor unitario", item["unit_price"]),
                    ("total", item["total"]),
                )
                if value is None or value == ""
            ]
            if missing:
                issues.append(
                    {
                        "Severidad": "Revisar",
                        "Tipo": "Dato de presupuesto faltante",
                        "Código": item["code"],
                        "Detalle": f"Falta {', '.join(missing)} en la partida.",
                    }
                )
        qty, unit_price, total = item["quantity"], item["unit_price"], item["total"]
        if qty is not None and unit_price is not None and total is not None:
            expected = qty * unit_price
            if abs(expected - total) > max(1.0, abs(total) * 0.01):
                issues.append(
                    {
                        "Severidad": "Revisar",
                        "Tipo": "Fórmula de presupuesto",
                        "Código": item["code"],
                        "Detalle": (
                            f"Cantidad × valor unitario ({expected:,.2f}) no coincide "
                            f"con el total informado ({total:,.2f})."
                        ),
                    }
                )
    duplicate_codes = defaultdict(list)
    for row in catalogs:
        if row["code"]:
            duplicate_codes[row["code"]].append(row)
    for code, matches in duplicate_codes.items():
        identities = {
            (norm(row["description"]), norm(row["unit"]), round(row["unit_cost"] or 0, 4))
            for row in matches
        }
        if len(identities) > 1:
            issues.append(
                {
                    "Severidad": "Revisar",
                    "Tipo": "Código de catálogo duplicado",
                    "Código": code,
                    "Detalle": (
                        "El código aparece con descripciones, unidades o precios "
                        "distintos; no se seleccionó un precio automáticamente."
                    ),
                }
            )
    for resource in resources:
        if resource["catalog_status"] != "Coincide":
            issue_type = (
                "Recurso sin catálogo"
                if resource["catalog_status"] == "Sin coincidencia"
                else "Inconsistencia APU-catálogo"
            )
            issues.append(
                {
                    "Severidad": "Revisar",
                    "Tipo": issue_type,
                    "Código": resource["resource_code"],
                    "Detalle": (
                        f"En {resource['sheet']}: {resource['catalog_status']} "
                        f"para «{resource['resource']}»."
                    ),
                }
            )
    budget_occurrences = Counter(item["code"] for item in budget if item["code"])
    for code, count in budget_occurrences.items():
        if count > 1:
            issues.append(
                {
                    "Severidad": "Revisar",
                    "Tipo": "Código de partida duplicado",
                    "Código": code,
                    "Detalle": f"El código aparece {count} veces en el presupuesto.",
                }
            )
    apu_by_code = {item["code"]: item for item in apus if item["code"]}
    for item in budget:
        apu = apu_by_code.get(item["code"])
        if (
            apu
            and item["unit_price"] is not None
            and apu["unit_price"] is not None
            and abs(item["unit_price"] - apu["unit_price"])
            > max(1.0, abs(item["unit_price"]) * 0.01)
        ):
            issues.append(
                {
                    "Severidad": "Revisar",
                    "Tipo": "Inconsistencia presupuesto-APU",
                    "Código": item["code"],
                    "Detalle": (
                        f"Valor unitario de presupuesto ({item['unit_price']:,.2f}) "
                        f"distinto del APU ({apu['unit_price']:,.2f})."
                    ),
                }
            )

    return {
        "sheets": list(sheets),
        "budget_sheet": budget_name,
        "budget": budget,
        "apus": apus,
        "resources": resources,
        "catalogs": catalogs,
        "overhead": overhead,
        "overhead_source": overhead_source,
        "issues": issues,
        "budget_codes": budget_codes,
    }


def detect_budget(sheets: dict[str, pd.DataFrame]):
    candidates = []
    for name, df in sheets.items():
        for i in range(min(len(df), 80)):
            row = df.iloc[i].tolist()
            item = find_column(row, ("item",))
            desc = find_column(row, ("designacion", "actividad", "descripcion"))
            unit = find_unit_column(row)
            qty = find_column(row, ("cantidad", "cant"))
            total = find_column(row, ("total",))
            if None not in (item, desc, unit, qty, total):
                score = int("ppto" in norm(name) or "presupuesto" in norm(name))
                candidates.append((score, name, i, {
                    "code": item,
                    "description": desc,
                    "unit": unit,
                    "quantity": qty,
                    "unit_price": find_price_column(row),
                    "total": total,
                }))
    if not candidates:
        return None, None, None
    candidates.sort(key=lambda result: result[0], reverse=True)
    _, name, header, columns = candidates[0]
    return name, header, columns


def parse_budget(
    df: pd.DataFrame, header: int, columns: dict[str, int | None], sheet_name: str
) -> list[dict[str, Any]]:
    result = []
    for row_number in range(header + 1, len(df)):
        row = df.iloc[row_number].tolist()
        code = code_text(value_at(row, columns["code"]))
        description = cell_text(value_at(row, columns["description"]))
        if not is_item_code(code) or not description:
            continue
        result.append(
            {
                "code": code,
                "description": description,
                "unit": cell_text(value_at(row, columns["unit"])),
                "quantity": number(value_at(row, columns["quantity"])),
                "unit_price": number(value_at(row, columns["unit_price"])),
                "total": number(value_at(row, columns["total"])),
                "sheet": sheet_name,
                "excel_row": row_number + 1,
            }
        )
    return result


def detect_apu_sheets(sheets: dict[str, pd.DataFrame], budget_name: str) -> list[str]:
    result = []
    for name, df in sheets.items():
        if name == budget_name:
            continue
        all_text = " ".join(
            norm(value) for value in df.head(40).to_numpy().flatten()
            if value is not None and not pd.isna(value)
        )
        has_resource_headers = (
            "codigo" in all_text
            and "rendimiento" in all_text
            and "perdida" in all_text
        )
        if has_resource_headers and (
            is_item_code(name) or "analisis de costos unitarios" in all_text
        ):
            result.append(name)
    return result


def parse_apus(
    sheets: dict[str, pd.DataFrame], names: list[str]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, str]]]:
    apus, resources, issues = [], [], []
    for name in names:
        df = sheets[name]
        item_code = code_text(name)
        activity = ""
        activity_unit = ""
        activity_qty = None
        activity_rate = None
        activity_total = None
        resource_headers = []
        category_context = ""

        for row_idx in range(min(len(df), 40)):
            row = df.iloc[row_idx].tolist()
            labels = [norm(value) for value in row]
            if "actividad" in labels and "cantidad" in labels and "total" in labels:
                code_col = find_column(row, ("item",))
                desc_col = find_column(row, ("actividad",))
                unit_col = find_column(row, ("unidad",))
                qty_col = find_column(row, ("cantidad",))
                rate_col = find_column(row, ("costo unitario directo", "costo unitario"))
                total_col = find_column(row, ("total",))
                for data_idx in range(row_idx + 1, min(row_idx + 5, len(df))):
                    data = df.iloc[data_idx].tolist()
                    possible_code = code_text(value_at(data, code_col))
                    possible_desc = cell_text(value_at(data, desc_col))
                    if possible_desc:
                        if possible_code:
                            item_code = possible_code
                        activity = possible_desc
                        activity_unit = cell_text(value_at(data, unit_col))
                        activity_qty = number(value_at(data, qty_col))
                        activity_rate = number(value_at(data, rate_col))
                        activity_total = number(value_at(data, total_col))
                        break
            if {"codigo", "rendimiento", "unidad"}.issubset(set(labels)) and (
                "perdida" in " ".join(labels)
            ):
                resource_headers.append(
                    {
                        "row": row_idx,
                        "code": find_column(row, ("codigo",)),
                        "description": find_column(row, ("descripcion recurso", "descripcion")),
                        "yield": find_column(row, ("rendimiento",)),
                        "unit": find_unit_column(row),
                        "loss": find_column(row, ("perdida",)),
                        "unit_cost": find_column(row, ("costo unitario",)),
                        "total": find_column(row, ("total",)),
                    }
                )

        apus.append(
            {
                "code": item_code,
                "description": activity,
                "unit": activity_unit,
                "quantity": activity_qty,
                "unit_price": activity_rate,
                "total": activity_total,
                "sheet": name,
            }
        )
        for header_idx, headers in enumerate(resource_headers):
            end_row = (
                resource_headers[header_idx + 1]["row"]
                if header_idx + 1 < len(resource_headers)
                else len(df)
            )
            for row_idx in range(headers["row"] + 1, end_row):
                row = df.iloc[row_idx].tolist()
                code = code_text(value_at(row, headers["code"])).upper()
                description = cell_text(value_at(row, headers["description"]))
                if not is_resource_code(code):
                    continue
                category = resource_category(code, category_context)
                if category == "Sin clasificar":
                    issues.append(
                        {
                            "Severidad": "Revisar",
                            "Tipo": "Categoría no identificada",
                            "Código": code,
                            "Detalle": f"Recurso en APU {name}: {description}.",
                        }
                    )
                yield_value = number(value_at(row, headers["yield"]))
                loss_factor = number(value_at(row, headers["loss"]))
                unit_cost = number(value_at(row, headers["unit_cost"]))
                reported_total = number(value_at(row, headers["total"]))
                if (
                    not description
                    or yield_value is None
                    or loss_factor is None
                    or unit_cost is None
                ):
                    issues.append(
                        {
                            "Severidad": "Revisar",
                            "Tipo": "Dato APU faltante",
                            "Código": code,
                            "Detalle": (
                                f"Falta descripción, rendimiento, factor de pérdida "
                                f"o costo en {name}, fila {row_idx + 1}."
                            ),
                        }
                    )
                expected = (
                    yield_value * loss_factor * unit_cost
                    if yield_value is not None
                    and loss_factor is not None
                    and unit_cost is not None
                    else None
                )
                if (
                    reported_total is not None
                    and expected is not None
                    and abs(expected - reported_total) > max(1, abs(reported_total) * .01)
                ):
                    issues.append(
                        {
                            "Severidad": "Revisar",
                            "Tipo": "Fórmula de recurso APU",
                            "Código": code,
                            "Detalle": (
                                f"{name}, fila {row_idx + 1}: rendimiento × factor de pérdida "
                                f"× costo unitario "
                                f"({expected:,.2f}) difiere del total informado "
                                f"({reported_total:,.2f})."
                            ),
                        }
                    )
                resources.append(
                    {
                        "item_code": item_code,
                        "item_description": activity or name,
                        "resource_code": code,
                        "resource": description,
                        "category": category,
                        "yield": yield_value,
                        "unit": cell_text(value_at(row, headers["unit"])),
                        "loss_factor": loss_factor,
                        "unit_cost": unit_cost,
                        "reported_total": reported_total,
                        "sheet": name,
                        "excel_row": row_idx + 1,
                    }
                )
    return apus, resources, issues


def parse_catalogs(
    sheets: dict[str, pd.DataFrame], excluded: set[str]
) -> list[dict[str, Any]]:
    records = []
    for name, df in sheets.items():
        if name in excluded:
            continue
        for row_idx in range(min(len(df), 35)):
            row = df.iloc[row_idx].tolist()
            normalized = [norm(value) for value in row]
            code_col = find_column(row, ("codigo",))
            if code_col is None:
                code_col = next(
                    (
                        i for i, value in enumerate(row)
                        if norm(value).startswith("n ")
                        and any(prefix in norm(value) for prefix in ("mat", "mo", "maq", "sub"))
                    ),
                    None,
                )
            desc_col = find_column(row, ("item", "material", "mano de obra", "maquinas y equipos", "subcontratos"))
            unit_col = find_unit_column(row)
            price_col = find_price_column(row)
            if None in (code_col, desc_col, unit_col, price_col):
                continue
            category = resource_category("", name)
            if category == "Sin clasificar":
                category = resource_category("", " ".join(normalized))
            if category == "Sin clasificar":
                continue
            for data_idx in range(row_idx + 1, len(df)):
                data = df.iloc[data_idx].tolist()
                code = code_text(value_at(data, code_col)).upper()
                desc = cell_text(value_at(data, desc_col))
                price = number(value_at(data, price_col))
                if not code or not desc or price is None:
                    continue
                if not re.match(r"^(MAT|MO|MAQ|SUB)-\d+$", code):
                    continue
                records.append(
                    {
                        "code": code,
                        "description": desc,
                        "unit": cell_text(value_at(data, unit_col)),
                        "unit_cost": price,
                        "category": resource_category(code, category),
                        "sheet": name,
                    }
                )
            break
    return records


def catalog_match(resource: dict[str, Any], catalogs: list[dict[str, Any]]) -> str:
    matches = [row for row in catalogs if row["code"] == resource["resource_code"]]
    if not matches:
        return "Sin coincidencia"
    identities = {
        (norm(row["description"]), norm(row["unit"]), round(row["unit_cost"], 4))
        for row in matches
    }
    if len(identities) > 1:
        return "Código ambiguo"
    candidate = matches[0]
    issues = []
    if norm(candidate["unit"]) != norm(resource["unit"]):
        issues.append("unidad distinta")
    if abs(candidate["unit_cost"] - resource["unit_cost"]) > max(
        1, abs(resource["unit_cost"]) * .01
    ):
        issues.append("precio distinto")
    return ", ".join(issues) if issues else "Coincide"


def parse_overhead(
    sheets: dict[str, pd.DataFrame], budget_name: str
) -> tuple[float | None, str | None]:
    for name, df in sheets.items():
        if name == budget_name or not any(
            token in norm(name) for token in ("gastos generales", "gastos grales")
        ):
            continue
        for row_idx in range(len(df)):
            row = df.iloc[row_idx].tolist()
            labels = [norm(value) for value in row]
            for col_idx, label in enumerate(labels):
                if "total" in label and any(term in label for term in ("gasto", "general", "obra")):
                    for value in row[col_idx + 1 :]:
                        found = number(value)
                        if found is not None and found > 0:
                            return found, name
        for row_idx, row in enumerate(df.itertuples(index=False, name=None)):
            for col_idx, value in enumerate(row):
                label = norm(value)
                if label in {"total", "total gastos generales", "total general"}:
                    for later_row in range(row_idx, min(len(df), row_idx + 3)):
                        for next_value in df.iloc[later_row].tolist()[col_idx + 1 :]:
                            found = number(next_value)
                            if found is not None and found > 0:
                                return found, name
    return None, None


def fmt_money(value: float | int | None) -> str:
    if value is None:
        return "—"
    return "$" + f"{value:,.0f}".replace(",", ".")


def fmt_number(value: float | int | None, digits: int = 2) -> str:
    if value is None:
        return "—"
    return f"{value:,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def figure_style(fig, height: int = 360):
    fig.update_layout(
        height=height,
        margin=dict(l=12, r=12, t=42, b=14),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=PALETTE["navy"], family="Arial"),
        legend_title_text="",
    )
    return fig


def create_excel(data: dict[str, Any], budget_view: pd.DataFrame, resource_view: pd.DataFrame, issues: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        budget_view.to_excel(writer, sheet_name="Presupuesto", index=False)
        resource_view.to_excel(writer, sheet_name="Recursos", index=False)
        issues.to_excel(writer, sheet_name="Validaciones", index=False)
        pd.DataFrame(data["catalogs"]).to_excel(writer, sheet_name="Catalogos", index=False)
    return output.getvalue()


def create_html_report(
    project: str,
    budget_total: float,
    overhead: float,
    construction_area: float,
    cost_per_m2: float | None,
    budget: pd.DataFrame,
    resources: pd.DataFrame,
    category_fig,
    pareto_fig,
) -> bytes:
    chart_one = pio.to_html(category_fig, full_html=False, include_plotlyjs="cdn")
    chart_two = pio.to_html(pareto_fig, full_html=False, include_plotlyjs=False)
    cost_m2 = fmt_money(cost_per_m2) if cost_per_m2 is not None else "No disponible"
    table_budget = budget.to_html(index=False, classes="data", border=0)
    table_resources = resources.to_html(index=False, classes="data", border=0)
    html = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Reporte de presupuesto | {html.escape(project)}</title>
<style>
body{{font-family:Arial,sans-serif;background:#f3f6fa;color:#17283d;margin:0}}
header{{background:#12253f;color:white;padding:28px 5vw}}header h1{{margin:0}}
main{{max-width:1500px;margin:24px auto;padding:0 20px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px}}
.card,.panel{{background:white;border:1px solid #e1e7ef;border-radius:12px;padding:18px;margin:14px 0}}
.card small{{color:#64748b}}.card strong{{display:block;font-size:1.55rem;color:#12253f;margin-top:8px}}
.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(400px,1fr));gap:14px}}
.table-wrap{{overflow:auto;max-height:520px}}table.data{{border-collapse:collapse;width:100%;font-size:13px}}
table.data th{{background:#12253f;color:white;text-align:left;position:sticky;top:0}}
table.data th,table.data td{{padding:8px;border-bottom:1px solid #e1e7ef}}
footer{{color:#64748b;padding:24px 5vw;font-size:12px}}
</style></head>
<body><header><h1>Reporte de análisis presupuestario</h1><p>{html.escape(project)} · Generado desde el libro Excel cargado</p></header>
<main><section class="kpis">
<div class="card"><small>Costo directo presupuestado</small><strong>{fmt_money(budget_total)}</strong></div>
<div class="card"><small>Gastos generales</small><strong>{fmt_money(overhead)}</strong></div>
<div class="card"><small>Superficie construida validada</small><strong>{fmt_number(construction_area)} m²</strong></div>
<div class="card"><small>Costo total por m²</small><strong>{cost_m2}</strong></div>
</section><div class="grid"><section class="panel"><h2>Distribución por categoría</h2>{chart_one}</section>
<section class="panel"><h2>Pareto de partidas</h2>{chart_two}</section></div>
<section class="panel"><h2>Partidas del presupuesto</h2><div class="table-wrap">{table_budget}</div></section>
<section class="panel"><h2>Recursos calculados</h2><div class="table-wrap">{table_resources}</div></section>
</main><footer>Los recursos se calculan como cantidad de partida × rendimiento × factor de pérdida, según las fórmulas del libro. Las unidades se mantienen separadas. Verifique las alertas del dashboard antes de usar este reporte como base contractual.</footer>
</body></html>"""
    return html.encode("utf-8")


st.markdown(
    """
    <div class="hero">
      <h1>Control presupuestario de construcción</h1>
      <p>Presupuesto, análisis de precios unitarios y demanda de recursos en un solo dashboard.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Fuente de datos")
    upload = st.file_uploader("Cargar presupuesto Excel", type=["xlsx", "xlsm"])
    default_files = sorted(
        path for path in Path(__file__).parent.glob("*.xlsx")
        if not path.name.startswith("~$")
    )
    if upload is not None:
        workbook_bytes = upload.getvalue()
        source_name = upload.name
    elif default_files:
        workbook_bytes = default_files[0].read_bytes()
        source_name = default_files[0].name
        st.caption(f"Libro local detectado: {source_name}")
    else:
        workbook_bytes = None
        source_name = ""
        st.info("Sube un libro Excel para iniciar el análisis.")

if not workbook_bytes:
    st.markdown(
        '<div class="note">Se admiten libros Excel con una hoja de presupuesto, '
        'catálogos y hojas APU. La aplicación detecta hojas y encabezados automáticamente.</div>',
        unsafe_allow_html=True,
    )
    st.stop()

try:
    parsed = parse_workbook(workbook_bytes)
except (ValueError, OSError, BadZipFile, InvalidFileException, KeyError) as exc:
    st.error(f"No fue posible analizar el libro: {exc}")
    st.stop()

budget_records = parsed["budget"]
apu_by_code = {item["code"]: item for item in parsed["apus"] if item["code"]}
budget_by_code = {item["code"]: item for item in budget_records}
resource_records = []
for resource in parsed["resources"]:
    budget_item = budget_by_code.get(resource["item_code"])
    if not budget_item or budget_item["quantity"] is None:
        continue
    loss_factor = resource["loss_factor"]
    resource_quantity = (
        budget_item["quantity"] * resource["yield"] * loss_factor
        if loss_factor is not None
        else None
    )
    resource_records.append(
        {
            "Partida": resource["item_code"],
            "Designación": budget_item["description"],
            "Código recurso": resource["resource_code"],
            "Recurso": resource["resource"],
            "Categoría": resource["category"],
            "Cantidad partida": budget_item["quantity"],
            "Unidad partida": budget_item["unit"],
            "Rendimiento": resource["yield"],
            "Factor pérdida": loss_factor,
            "Cantidad recurso": resource_quantity,
            "Unidad recurso": resource["unit"],
            "Costo unitario": resource["unit_cost"],
            "Costo calculado": (
                resource_quantity * resource["unit_cost"]
                if resource_quantity is not None
                else None
            ),
            "Catálogo": resource["catalog_status"],
            "Hoja APU": resource["sheet"],
        }
    )

budget_df = pd.DataFrame(
    [
        {
            "Partida": item["code"],
            "Designación": item["description"],
            "Unidad": item["unit"],
            "Cantidad": item["quantity"],
            "Valor unitario": item["unit_price"],
            "Total": item["total"],
            "Vinculación APU": item["apu_status"],
        }
        for item in budget_records
    ]
)
resource_df = pd.DataFrame(resource_records)
if resource_df.empty:
    resource_df = pd.DataFrame(columns=[
        "Partida", "Designación", "Código recurso", "Recurso", "Categoría",
        "Cantidad partida", "Unidad partida", "Rendimiento", "Factor pérdida",
        "Cantidad recurso", "Unidad recurso", "Costo unitario", "Costo calculado",
        "Catálogo", "Hoja APU",
    ])

# Parent headings are excluded when a more detailed, non-zero descendant exists.
line_items = []
for item in budget_records:
    total = item["total"]
    if total is None:
        continue
    has_child = any(
        other["code"].startswith(item["code"] + ".")
        and other["total"] not in (None, 0)
        for other in budget_records
        if other["code"] != item["code"]
    )
    if not has_child and total != 0:
        line_items.append(item)
direct_cost = sum(item["total"] or 0 for item in line_items)

with st.sidebar:
    st.divider()
    st.subheader("Supuestos y filtros")
    construction_area = st.number_input(
        "Superficie construida validada (m²)",
        min_value=0.0,
        value=0.0,
        step=10.0,
        help="Ingrese la superficie total del proyecto según antecedente validado. "
        "No se infiere sumando partidas en m².",
    )
    detected_overhead = parsed["overhead"]
    overhead = st.number_input(
        "Gastos generales ($)",
        min_value=0.0,
        value=float(detected_overhead or 0),
        step=100000.0,
        help="Se precarga solo cuando se detecta un total legible en la hoja de gastos generales. "
        "Ajuste o ingrese el valor validado.",
    )
    st.caption("El factor de pérdida se aplica directamente, como en las fórmulas del APU.")

if construction_area <= 0:
    st.warning(
        "Ingrese una superficie construida mayor que cero y validada para calcular el costo por m². "
        "Las superficies de partidas no se usan como sustituto."
    )

budget_total = sum(item["total"] or 0 for item in line_items)
total_project = budget_total + overhead
cost_per_m2 = total_project / construction_area if construction_area > 0 else None

st.caption(
    f"Libro: **{source_name}** · Hoja de presupuesto: **{parsed['budget_sheet']}** · "
    f"Hojas detectadas: **{len(parsed['sheets'])}** · APU: **{len(parsed['apus'])}**"
)

metric_cols = st.columns(5)
metric_cols[0].metric("Partidas con monto", f"{len(line_items):,}")
metric_cols[1].metric("Costo directo", fmt_money(budget_total))
metric_cols[2].metric("Gastos generales", fmt_money(overhead))
metric_cols[3].metric("Presupuesto total", fmt_money(total_project))
metric_cols[4].metric("Costo por m²", fmt_money(cost_per_m2))

if parsed["overhead"] is None:
    st.info("No se detectó un total inequívoco de gastos generales; se muestra $0 hasta ingresar el valor validado.")
elif parsed["overhead_source"]:
    st.caption(f"Gastos generales detectados en: {parsed['overhead_source']}.")

tabs = st.tabs(["Resumen", "Partidas", "Recursos", "Validaciones", "Exportar"])

with tabs[0]:
    budget_totals = budget_df.dropna(subset=["Total"]).copy()
    budget_totals["Total"] = pd.to_numeric(budget_totals["Total"], errors="coerce")
    budget_totals = budget_totals.dropna(subset=["Total"])
    budget_totals = budget_totals[budget_totals["Total"] > 0]
    if not budget_totals.empty:
        budget_totals = budget_totals[
            ~budget_totals["Partida"].apply(
                lambda code: any(
                    other.startswith(str(code) + ".") and amount > 0
                    for other, amount in zip(budget_totals["Partida"], budget_totals["Total"])
                )
            )
        ]

    left, right = st.columns([1.2, 1])
    with left:
        if resource_df.empty:
            st.info("No se encontraron recursos APU vinculados para graficar.")
        else:
            category_costs = (
                resource_df.groupby("Categoría", as_index=False)["Costo calculado"]
                .sum()
                .sort_values("Costo calculado", ascending=False)
            )
            fig = px.pie(
                category_costs,
                names="Categoría",
                values="Costo calculado",
                hole=.58,
                color="Categoría",
                color_discrete_map=CATEGORY_COLORS,
                title="Costo de recursos por categoría",
            )
            st.plotly_chart(figure_style(fig), use_container_width=True)
    with right:
        if budget_totals.empty:
            st.info("No hay partidas con total para el gráfico.")
            pareto_fig = px.bar(title="Pareto de partidas")
        else:
            top = budget_totals.nlargest(15, "Total").sort_values("Total", ascending=False).copy()
            top["Acumulado (%)"] = top["Total"].cumsum() / max(top["Total"].sum(), 1) * 100
            pareto_fig = px.bar(
                top,
                x="Partida",
                y="Total",
                title="15 partidas de mayor costo",
                color_discrete_sequence=[PALETTE["navy_light"]],
                hover_data=["Designación"],
            )
            pareto_fig.add_scatter(
                x=top["Partida"],
                y=top["Acumulado (%)"],
                name="Acumulado (%)",
                yaxis="y2",
                mode="lines+markers",
                line=dict(color=PALETTE["orange"], width=3),
            )
            pareto_fig.update_layout(
                yaxis2=dict(overlaying="y", side="right", range=[0, 105], ticksuffix="%")
            )
            st.plotly_chart(figure_style(pareto_fig), use_container_width=True)

    st.subheader("Magnitudes por unidad de presupuesto")
    if budget_df.empty:
        st.info("No se identificaron partidas.")
    else:
        quantity_df = budget_df.copy()
        quantity_df["Cantidad"] = pd.to_numeric(quantity_df["Cantidad"], errors="coerce")
        quantity_df["Unidad normalizada"] = quantity_df["Unidad"].astype(str).str.strip().str.lower()
        quantity_df = quantity_df.dropna(subset=["Cantidad"])
        quantity_df = quantity_df[
            quantity_df["Unidad normalizada"].isin(
                {"m2", "m²", "m 2", "m3", "m³", "m 3"}
            )
        ]
        if quantity_df.empty:
            st.info("No hay partidas con unidades m² o m³ reconocidas.")
        else:
            quantity_df["Magnitud"] = quantity_df["Unidad normalizada"].map(
                {"m2": "Superficie (m²)", "m²": "Superficie (m²)", "m 2": "Superficie (m²)",
                 "m3": "Volumen (m³)", "m³": "Volumen (m³)", "m 3": "Volumen (m³)"}
            )
            magnitudes = quantity_df.groupby("Magnitud", as_index=False)["Cantidad"].sum()
            st.dataframe(magnitudes, hide_index=True, use_container_width=True)
            st.caption(
                "Totales agrupados solo dentro de la misma unidad. Son cubicaciones de partidas, "
                "no la superficie total construida del proyecto."
            )

with tabs[1]:
    st.subheader("Partidas presupuestarias")
    if not budget_df.empty:
        f1, f2, f3 = st.columns(3)
        codes = sorted(budget_df["Partida"].dropna().unique())
        chosen_codes = f1.multiselect("Partida", codes, default=[])
        units = sorted(value for value in budget_df["Unidad"].dropna().unique() if value)
        chosen_units = f2.multiselect("Unidad", units, default=[])
        statuses = sorted(budget_df["Vinculación APU"].dropna().unique())
        chosen_status = f3.multiselect("Vínculo APU", statuses, default=statuses)
        view = budget_df.copy()
        if chosen_codes:
            view = view[view["Partida"].isin(chosen_codes)]
        if chosen_units:
            view = view[view["Unidad"].isin(chosen_units)]
        if chosen_status:
            view = view[view["Vinculación APU"].isin(chosen_status)]
        st.dataframe(
            view,
            hide_index=True,
            use_container_width=True,
            column_config={
                "Cantidad": st.column_config.NumberColumn(format="%.3f"),
                "Valor unitario": st.column_config.NumberColumn(format="$ %.0f"),
                "Total": st.column_config.NumberColumn(format="$ %.0f"),
            },
        )

with tabs[2]:
    st.subheader("Consumo estimado de recursos")
    if resource_df.empty:
        st.info("No se hallaron líneas de recursos vinculadas con cantidades de presupuesto.")
    else:
        f1, f2, f3, f4 = st.columns(4)
        chosen_items = f1.multiselect(
            "Partida", sorted(resource_df["Partida"].unique()), key="resource_items"
        )
        chosen_categories = f2.multiselect(
            "Categoría", sorted(resource_df["Categoría"].unique()),
            default=sorted(resource_df["Categoría"].unique()), key="resource_categories"
        )
        chosen_resources = f3.multiselect(
            "Recurso", sorted(resource_df["Recurso"].unique()), key="resource_names"
        )
        chosen_resource_units = f4.multiselect(
            "Unidad", sorted(resource_df["Unidad recurso"].dropna().unique()),
            key="resource_units"
        )
        resource_view = resource_df.copy()
        if chosen_items:
            resource_view = resource_view[resource_view["Partida"].isin(chosen_items)]
        if chosen_categories:
            resource_view = resource_view[resource_view["Categoría"].isin(chosen_categories)]
        if chosen_resources:
            resource_view = resource_view[resource_view["Recurso"].isin(chosen_resources)]
        if chosen_resource_units:
            resource_view = resource_view[
                resource_view["Unidad recurso"].isin(chosen_resource_units)
            ]

        chart_col, pareto_col = st.columns(2)
        with chart_col:
            by_category = (
                resource_view.groupby("Categoría", as_index=False)["Costo calculado"]
                .sum().sort_values("Costo calculado", ascending=False)
            )
            category_bar = px.bar(
                by_category, x="Categoría", y="Costo calculado",
                color="Categoría", color_discrete_map=CATEGORY_COLORS,
                title="Costo estimado por categoría",
            )
            st.plotly_chart(figure_style(category_bar), use_container_width=True)
        with pareto_col:
            by_resource = (
                resource_view.groupby(["Recurso", "Unidad recurso"], as_index=False)
                .agg({"Costo calculado": "sum", "Cantidad recurso": "sum"})
                .sort_values("Costo calculado", ascending=False)
            )
            by_resource["Acumulado (%)"] = (
                by_resource["Costo calculado"].cumsum()
                / max(by_resource["Costo calculado"].sum(), 1) * 100
            )
            top_resources = by_resource.head(15)
            resource_pareto = px.bar(
                top_resources, x="Recurso", y="Costo calculado",
                title="Pareto de recursos por costo",
                color_discrete_sequence=[PALETTE["navy_light"]],
            )
            resource_pareto.add_scatter(
                x=top_resources["Recurso"],
                y=top_resources["Acumulado (%)"],
                name="Acumulado (%)",
                yaxis="y2",
                mode="lines+markers",
                line=dict(color=PALETTE["orange"], width=3),
            )
            resource_pareto.update_layout(
                yaxis2=dict(overlaying="y", side="right", range=[0, 105], ticksuffix="%")
            )
            st.plotly_chart(figure_style(resource_pareto), use_container_width=True)

        unit_totals = (
            resource_view.groupby(["Categoría", "Recurso", "Unidad recurso"], as_index=False)
            .agg({"Cantidad recurso": "sum", "Costo calculado": "sum"})
            .sort_values("Costo calculado", ascending=False)
        )
        st.dataframe(
            resource_view,
            hide_index=True,
            use_container_width=True,
            column_config={
                "Cantidad recurso": st.column_config.NumberColumn(format="%.4f"),
                "Rendimiento": st.column_config.NumberColumn(format="%.6f"),
                "Factor pérdida": st.column_config.NumberColumn(format="%.4f"),
                "Costo unitario": st.column_config.NumberColumn(format="$ %.0f"),
                "Costo calculado": st.column_config.NumberColumn(format="$ %.0f"),
            },
        )
        st.markdown("**Totales agrupados por recurso y unidad**")
        st.dataframe(unit_totals, hide_index=True, use_container_width=True)

with tabs[3]:
    st.subheader("Calidad y consistencia")
    apu_codes = {item["code"] for item in parsed["apus"]}
    if parsed["issues"]:
        issue_df = pd.DataFrame(parsed["issues"])
        st.warning(f"Se encontraron {len(issue_df)} observaciones que requieren revisión.")
        st.dataframe(issue_df, hide_index=True, use_container_width=True)
    else:
        st.success("No se detectaron inconsistencias según las validaciones disponibles.")

    linked_count = sum(1 for item in budget_records if item["code"] in apu_codes)
    st.caption(
        f"Vinculación exacta por código de partida: {linked_count}/{len(budget_records)} partidas. "
        "Los vínculos parciales o por semejanza de descripción no se inventan."
    )
    st.markdown(
        '<div class="note">Controles aplicados: fórmulas de cantidad × valor unitario, '
        'rendimiento × costo unitario, duplicidad de códigos en catálogos, código exacto '
        'de vínculo, datos requeridos y diferencias de unidad/precio.</div>',
        unsafe_allow_html=True,
    )

with tabs[4]:
    st.subheader("Descargas")
    st.write("Exporta los resultados actuales del libro, incluyendo las observaciones de validación.")
    issue_export = pd.DataFrame(parsed["issues"])
    if issue_export.empty:
        issue_export = pd.DataFrame(columns=["Severidad", "Tipo", "Código", "Detalle"])
    all_resource_export = resource_df.copy()
    excel_bytes = create_excel(parsed, budget_df, all_resource_export, issue_export)
    st.download_button(
        "Descargar análisis Excel",
        data=excel_bytes,
        file_name="analisis_presupuesto_construccion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    category_costs_for_report = (
        resource_df.groupby("Categoría", as_index=False)["Costo calculado"].sum()
        if not resource_df.empty
        else pd.DataFrame({"Categoría": [], "Costo calculado": []})
    )
    report_category_fig = px.pie(
        category_costs_for_report,
        names="Categoría",
        values="Costo calculado",
        hole=.58,
        color="Categoría",
        color_discrete_map=CATEGORY_COLORS,
        title="Costo de recursos por categoría",
    )
    report_pareto_df = budget_df[
        budget_df["Partida"].isin([item["code"] for item in line_items])
    ].copy()
    report_pareto_df["Total"] = pd.to_numeric(report_pareto_df["Total"], errors="coerce")
    report_pareto_df = report_pareto_df.dropna(subset=["Total"])
    report_pareto_df = report_pareto_df.nlargest(15, "Total")
    report_pareto_fig = px.bar(
        report_pareto_df,
        x="Partida",
        y="Total",
        title="Partidas de mayor costo",
        hover_data=["Designación"],
        color_discrete_sequence=[PALETTE["navy_light"]],
    )
    if not report_pareto_df.empty:
        report_pareto_df["Acumulado (%)"] = (
            report_pareto_df["Total"].cumsum()
            / max(report_pareto_df["Total"].sum(), 1) * 100
        )
        report_pareto_fig.add_scatter(
            x=report_pareto_df["Partida"],
            y=report_pareto_df["Acumulado (%)"],
            name="Acumulado (%)",
            yaxis="y2",
            mode="lines+markers",
            line=dict(color=PALETTE["orange"], width=3),
        )
        report_pareto_fig.update_layout(
            yaxis2=dict(overlaying="y", side="right", range=[0, 105], ticksuffix="%")
        )
    project_name = Path(source_name).stem
    html_bytes = create_html_report(
        project_name,
        budget_total,
        overhead,
        construction_area,
        cost_per_m2,
        budget_df,
        all_resource_export,
        report_category_fig,
        report_pareto_fig,
    )
    st.download_button(
        "Descargar reporte interactivo HTML",
        data=html_bytes,
        file_name="reporte_interactivo_presupuesto.html",
        mime="text/html",
        use_container_width=True,
    )
    st.caption(
        "El HTML incorpora gráficos Plotly interactivos. La librería Plotly se carga desde CDN "
        "al abrir el reporte; requiere conexión a internet para mostrar los gráficos."
    )

st.divider()
st.caption(
    "Cálculo de demanda según las fórmulas del libro: cantidad presupuestada × "
    "rendimiento APU × factor de pérdida. No se consolidan unidades distintas. "
    "Valide los supuestos y las alertas antes de emitir decisiones de compra o contrato."
)
