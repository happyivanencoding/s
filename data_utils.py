# -*- coding: utf-8 -*-
"""Utilitaires de données et de fichiers Excel.

1. Lecture des fichiers et dates mensuelles.
2. Historique : conserver la première observation de chaque mois.
3. Lecture des blocs, rendements, macro et taux.
4. Conservation des entrées, des graphiques et des métadonnées Excel.

Les chemins et les paramètres sont définis dans local_config.py.
"""

from copy import deepcopy
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

import numpy as np
import pandas as pd
from lxml import etree as ET
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from local_config import HISTORY_CONFIG, MACRO_EXCEL_FILE, EU_CONFIG, US_CONFIG

NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
S = "{" + NS["s"] + "}"
_history = None
_history_changed = False

# 1. Fichiers et dates


def open_workbooks(excel_file=None, macro_file=None):
    """Lit les valeurs enregistrées des deux fichiers source."""
    excel_file = Path(excel_file or EU_CONFIG["excel_file"])
    macro_file = Path(macro_file or MACRO_EXCEL_FILE)
    sector_workbook = load_workbook(excel_file, data_only=True)
    macro_workbook = load_workbook(macro_file, data_only=True)
    return sector_workbook, macro_workbook


def is_number(x):
    """Indique si une cellule contient une valeur numérique exploitable."""
    return isinstance(
        x,
        (int, float, np.integer, np.floating),
    ) and not pd.isna(x)


def convert_excel_date(x):
    """Convertit une date Excel ou Python en Timestamp pandas."""
    if isinstance(x, (pd.Timestamp, np.datetime64)):
        return pd.Timestamp(x)

    if hasattr(x, "year") and hasattr(x, "month") and hasattr(x, "day"):
        return pd.Timestamp(x)

    if is_number(x):
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(
            float(x),
            unit="D",
        )

    return pd.NaT


def is_excel_date(x):
    """Vérifie qu'une cellule contient une date exploitable."""
    date = convert_excel_date(x)

    if pd.isna(date):
        return False

    return 1980 <= date.year <= 2100


def normalize_month_end(x):
    """Utilise le dernier jour calendaire du mois comme clé de date."""
    date = convert_excel_date(x)

    if pd.isna(date):
        return pd.NaT

    return pd.Timestamp(date).to_period("M").to_timestamp("M")


def normalize_monthly_index(df):
    """Normalise l'index en fins de mois et conserve une ligne par mois."""
    result = df.copy()
    result.index = pd.DatetimeIndex([normalize_month_end(x) for x in result.index])
    result = result[~result.index.duplicated(keep="first")]
    return result.sort_index(ascending=False)


# 2. Historique mensuel


def history_path():
    """Retourne le fichier Excel qui contient tout l'historique."""
    return Path(__file__).resolve().parent / HISTORY_CONFIG["file"]


def load_history(reload=False):
    """Charge l'historique Excel une fois par exécution."""
    global _history, _history_changed
    if _history is None or reload:
        file_path = history_path()
        _history = (
            pd.read_excel(file_path, sheet_name="Historique")
            if file_path.exists()
            else pd.DataFrame()
        )
        for column in ("date", "date_source"):
            if column in _history:
                _history[column] = pd.to_datetime(_history[column])
        _history_changed = False
    return _history


def save_history():
    """Écrit les nouveaux mois sans remplacer les observations figées."""
    global _history_changed
    if not _history_changed:
        return
    file_path = history_path()
    temporary = file_path.with_name(file_path.stem + ".tmp.xlsx")
    _history.sort_values(["series_key", "date"], ascending=[True, False]).to_excel(
        temporary, sheet_name="Historique", index=False, freeze_panes=(1, 3)
    )
    normalize_metadata(temporary)
    temporary.replace(file_path)
    _history_changed = False


def freeze_history(df, source_dates, key):
    """
    Conserve la première observation enregistrée pour chaque mois.
    Une révision ultérieure du fichier source ne remplace pas le mois existant.
    """
    result = normalize_monthly_index(df)

    source_dates = pd.Series(
        [convert_excel_date(x) for x in source_dates],
        index=[normalize_month_end(x) for x in source_dates],
        name="date_source",
    )
    source_dates = source_dates[~source_dates.index.duplicated(keep="first")]

    current = result.copy()
    current.insert(
        0,
        "date_source",
        source_dates.reindex(current.index).values,
    )
    current.insert(0, "date", current.index)
    current.insert(0, "series_key", key)
    current = current.reset_index(drop=True)

    data_columns = list(df.columns)
    current = current[current[data_columns].notna().any(axis=1)]

    if not HISTORY_CONFIG.get("enabled", True):
        return result

    global _history, _history_changed
    history = load_history()
    existing_dates = (
        history.loc[history["series_key"] == key, "date"]
        if not history.empty
        else pd.Series(dtype="datetime64[ns]")
    )
    new_months = current.loc[~current["date"].isin(existing_dates)]
    if not new_months.empty:
        history = pd.concat([history, new_months], ignore_index=True, sort=False)
        _history = history
        _history_changed = True

    if history.empty:
        return result

    series = history[history["series_key"] == key].copy()

    series = series.sort_values(
        "date",
        ascending=False,
    ).set_index("date")

    columns = list(df.columns)

    for column in columns:
        if column not in series.columns:
            series[column] = np.nan

    return series[columns].copy()


# 3. Blocs de données


def read_sector_block(
    ws,
    start_column,
    sectors=None,
    key_prefix="",
    start_row=8,
    date_column="A",
):
    """Lit un bloc sectoriel puis fige son historique mensuel."""
    sectors = sectors or EU_CONFIG["sectors"]

    col_start = column_index_from_string(start_column)
    col_date = column_index_from_string(date_column)

    dates = []
    source_dates = []
    data = []
    row_number = start_row

    while True:
        raw_date = ws.cell(row_number, col_date).value

        if not is_excel_date(raw_date):
            break

        values = []

        for j in range(len(sectors)):
            value = ws.cell(
                row_number,
                col_start + j,
            ).value
            values.append(float(value) if is_number(value) else np.nan)

        source_date = convert_excel_date(raw_date)
        dates.append(source_date)
        source_dates.append(source_date)
        data.append(values)
        row_number += 1

    df = pd.DataFrame(
        data,
        index=dates,
        columns=sectors,
    )

    key = f"{key_prefix}{ws.title}_{start_column}"
    return freeze_history(
        df,
        source_dates,
        key,
    )


def read_return_indices(ws, sectors):
    """Lit les indices de rendement total avant tout calcul Excel."""
    rows = []
    dates = []
    for row_number in range(7, ws.max_row + 1):
        date = ws.cell(row_number, 1).value
        if not is_excel_date(date):
            break
        dates.append(convert_excel_date(date))
        rows.append(
            [
                (
                    float(ws.cell(row_number, column).value)
                    if is_number(ws.cell(row_number, column).value)
                    else np.nan
                )
                for column in range(2, len(sectors) + 3)
            ]
        )
    return pd.DataFrame(rows, index=dates, columns=["benchmark"] + sectors).sort_index(
        ascending=False
    )


def read_sector_returns(ws, volatility_config=None, sectors=None, key_prefix=""):
    """Calcule les rendements mensuels à partir des indices source."""
    sectors = sectors or EU_CONFIG["sectors"]
    indices = read_return_indices(ws, sectors)
    returns = indices / indices.shift(-1) - 1
    return freeze_history(
        returns[sectors],
        returns.index,
        f"{key_prefix}{ws.title}_returns",
    )


def read_benchmark_returns(ws, volatility_config=None, key_prefix=""):
    """Calcule les rendements du benchmark sans changer sa source."""
    sectors = US_CONFIG["sectors"] if key_prefix == "US_" else EU_CONFIG["sectors"]
    indices = read_return_indices(ws, sectors)
    returns = indices[["benchmark"]] / indices[["benchmark"]].shift(-1) - 1
    return freeze_history(
        returns,
        returns.index,
        f"{key_prefix}{ws.title}_benchmark",
    )


def extend_macro(ws, cfg, regimes):
    """Prolonge les résultats macro avec les règles du classeur."""
    score_col = column_index_from_string(cfg["score_column"])
    cycle_col = column_index_from_string(cfg["regime_column"])
    rows = [
        r
        for r in range(cfg["start_row"], ws.max_row + 1)
        if is_number(ws.cell(r, score_col).value)
        and ws.cell(r, cycle_col).value in regimes
    ]
    if not rows:
        return

    def read_data(col, r):
        return ws[f"{col}{r}"].value

    def cycle(score, previous):
        if score > previous:
            return "E" if score >= 0 else "R"
        return "SD" if score >= 0 else "C"

    for r in range(rows[-1] + 1, ws.max_row + 1):
        if not is_excel_date(read_data("A", r)):
            break
        if cfg["sheet"] == "Europe":
            values = [read_data("D", r - 1), read_data("I", r), read_data("N", r)]
            if not all(is_number(v) for v in values):
                break
            score = sum(values) / 3
            previous = read_data("R", r - 1)
            simple = cycle(score, previous)
            threshold = ws["T1"].value
            max_threshold = ws["W1"].value
            filtered = (
                simple
                if (abs(score - previous) > threshold or score * previous < 0)
                else read_data("T", r - 1)
            )
            pivot = (
                score if filtered != read_data("S", r - 1) else read_data("U", r - 1)
            )
            previous_pivot = read_data("U", r - 1)
            maximum = max(read_data("R", i) for i in range(r - 12, r))
            regime = (
                simple
                if (
                    abs(pivot - previous_pivot) > threshold
                    or pivot * previous_pivot < 0
                    or abs(score - maximum) > max_threshold
                )
                else read_data("V", r - 1)
            )
            for col, value in zip(
                ["R", "S", "T", "U", "V"], [score, simple, filtered, pivot, regime]
            ):
                ws[f"{col}{r}"] = value
        else:
            values = [
                read_data("M", r),
                read_data("J", r),
                read_data("G", r - 1),
                read_data("D", r),
            ]
            if not all(is_number(v) for v in values):
                break
            score = sum(v * weights for v, weights in zip(values, [0.3, 0.3, 0.3, 0.1]))
            ws[f"Y{r}"] = score
            ws[f"Z{r}"] = cycle(score, read_data("Y", r - 1))


def read_macro_data(
    macro_workbook,
    macro_config=None,
    regime_weights=None,
    key_prefix="",
):
    """Lit le macro score et le régime du marché."""
    cfg = macro_config or EU_CONFIG["macro"]
    regime_weights = regime_weights or EU_CONFIG["regime_weights"]

    ws = macro_workbook[cfg["sheet"]]
    extend_macro(ws, cfg, regime_weights)

    row_number = cfg["start_row"]
    col_date = column_index_from_string(cfg["date_column"])
    col_score = column_index_from_string(cfg["score_column"])
    col_regime = column_index_from_string(cfg["regime_column"])

    source_dates = []
    data = []

    while True:
        raw_date = ws.cell(
            row_number,
            col_date,
        ).value

        if not is_excel_date(raw_date):
            break

        score = ws.cell(
            row_number,
            col_score,
        ).value
        regime = ws.cell(
            row_number,
            col_regime,
        ).value
        source_date = convert_excel_date(raw_date)

        source_dates.append(source_date)
        data.append(
            {
                "macro_score": (float(score) if is_number(score) else np.nan),
                "cycle": (regime if regime in regime_weights else None),
            }
        )
        row_number += 1

    df = pd.DataFrame(
        data,
        index=source_dates,
    )

    df = freeze_history(
        df,
        source_dates,
        f"{key_prefix}macro_outputs",
    )

    return df.sort_index()


def read_us10y(
    sector_workbook,
    signal_config=None,
    key_prefix="",
):
    """Lit la série US 10Y utilisée par le rate overlay."""
    cfg = signal_config or EU_CONFIG["rate_signal"]
    ws = sector_workbook[cfg["sheet"]]

    row_number = cfg["start_row"]
    col_date = column_index_from_string(cfg["date_column"])
    col_us10y = column_index_from_string(cfg["us10y_column"])

    source_dates = []
    data = []

    while True:
        raw_date = ws.cell(
            row_number,
            col_date,
        ).value

        if not is_excel_date(raw_date):
            break

        rate = ws.cell(
            row_number,
            col_us10y,
        ).value
        source_date = convert_excel_date(raw_date)

        source_dates.append(source_date)
        data.append(
            {
                "us10y": (float(rate) if is_number(rate) else np.nan),
            }
        )
        row_number += 1

    df = pd.DataFrame(
        data,
        index=source_dates,
    )

    df = freeze_history(
        df,
        source_dates,
        f"{key_prefix}signal_taux_us10y",
    )

    return df.sort_index()


# 4. Format natif des classeurs


def normalize_metadata(file_path):
    """Conserve les noms usuels des styles et retire le chemin local."""
    from openpyxl.styles.builtins import styles

    names = {str(style.builtinId): name for name, style in styles.items()}
    font_names = {"Hans": "DengXian", "Hant": "PMingLiU", "Jpan": "Yu Gothic"}
    character_pattern = re.compile(r"[\u3400-\u9fff\uf900-\ufaff]")
    changes = {}
    with ZipFile(file_path) as archive:
        for path in archive.namelist():
            if path not in (
                "xl/workbook.xml",
                "xl/styles.xml",
                "docProps/app.xml",
            ) and not path.startswith("xl/theme/"):
                continue
            content = archive.read(path)
            document = ET.fromstring(content)
            xml_text = ET.tostring(document, encoding="unicode")
            if not character_pattern.search(xml_text) and "absPath" not in xml_text:
                continue
            for element in list(document.iter()):
                name = ET.QName(element).localname
                if name == "absPath":
                    element.getparent().remove(element)
                elif name == "cellStyle" and element.get("builtinId") in names:
                    element.set("name", names[element.get("builtinId")])
                elif name == "numFmt" and any(
                    x in element.get("formatCode", "") for x in ("mmm", "ddd")
                ):
                    date_format = element.get("formatCode")
                    if "[$-" not in date_format:
                        element.set("formatCode", "[$-040C]" + date_format)
                elif name == "theme" and character_pattern.search(
                    element.get("name", "")
                ):
                    element.set("name", "Office")
                elif name == "font" and element.get("script") in font_names:
                    font_name = element.get("typeface", "")
                    if character_pattern.search(font_name):
                        element.set(
                            "typeface",
                            font_names[element.get("script")]
                            + (" Light" if font_name.endswith(" Light") else ""),
                        )
            if path == "docProps/app.xml":
                labels = document.xpath(
                    "//*[local-name()='HeadingPairs']//*[local-name()='lpstr']"
                )
                for element, name in zip(labels, ["Worksheets", "Named Ranges"]):
                    element.text = name
            changes[path] = ET.tostring(
                document, encoding="UTF-8", xml_declaration=True
            )
        if not changes:
            return
        temporary = file_path.with_name(file_path.stem + ".langue" + file_path.suffix)
        with ZipFile(temporary, "w", ZIP_DEFLATED) as output:
            for element in archive.infolist():
                output.writestr(
                    element,
                    changes.get(element.filename, archive.read(element.filename)),
                )
    temporary.replace(file_path)


def sheet_paths(archive):
    workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
    relations = {
        e.get("Id"): e.get("Target")
        for e in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    }
    result = {}
    for sheet in workbook_xml.find("s:sheets", NS):
        path = relations[sheet.get("{" + NS["r"] + "}id")]
        result[sheet.get("name")] = (
            path.lstrip("/") if path.startswith("/") else "xl/" + path
        )
    return result


def preserve_inputs(source, destination):
    """Conserve les données de collecte et la présentation des graphiques."""
    # Un enregistrement Excel peut recalculer les entrées ou modifier les axes.
    # Reprendre ces éléments dans le fichier source conserve le modèle visuel.
    with ZipFile(source) as before, ZipFile(destination) as after:
        sheets_before, sheets_after = sheet_paths(before), sheet_paths(after)
        texts = []
        if "xl/sharedStrings.xml" in before.namelist():
            texts = [
                "".join(e.itertext())
                for e in ET.fromstring(before.read("xl/sharedStrings.xml"))
            ]
        changes = {}
        for name, path in sheets_before.items():
            if name in ("Dashboard", "Output", "Reco_histo"):
                continue
            data = deepcopy(ET.fromstring(before.read(path)).find("s:sheetData", NS))
            for cell in data.findall(".//s:c[@t='s']", NS):
                value = cell.find("s:v", NS)
                text = texts[int(value.text)]
                cell.remove(value)
                if cell.find("s:f", NS) is None:
                    cell.set("t", "inlineStr")
                    ET.SubElement(ET.SubElement(cell, S + "is"), S + "t").text = text
                else:
                    cell.set("t", "str")
                    ET.SubElement(cell, S + "v").text = text
            sheet = ET.fromstring(after.read(sheets_after[name]))
            sheet.replace(sheet.find("s:sheetData", NS), data)
            changes[sheets_after[name]] = ET.tostring(
                sheet, encoding="UTF-8", xml_declaration=True
            )
        # Excel peut modifier les axes pendant le chargement des données.
        for path in before.namelist():
            if not re.fullmatch(r"xl/charts/chart\d+\.xml", path):
                continue
            original = ET.fromstring(before.read(path))
            new = ET.fromstring(after.read(path))
            for cache in original.xpath(
                "//*[local-name()='numCache' or local-name()='strCache']"
            ):
                position = original.getroottree().getpath(cache)
                updated = new.xpath(
                    position, namespaces={k: v for k, v in cache.nsmap.items() if k}
                )
                if updated:
                    cache.getparent().replace(cache, deepcopy(updated[0]))
            changes[path] = ET.tostring(
                original, encoding="UTF-8", xml_declaration=True
            )
        for path in before.namelist():
            if re.fullmatch(r"xl/charts/(style|colors)\d+\.xml", path):
                changes[path] = before.read(path)
        workbook_xml = ET.fromstring(after.read("xl/workbook.xml"))
        calculation_settings = workbook_xml.find("s:calcPr", NS)
        if calculation_settings is not None:
            calculation_settings.set("calcMode", "auto")
            calculation_settings.set("fullCalcOnLoad", "0")
            calculation_settings.set("forceFullCalc", "0")
        changes["xl/workbook.xml"] = ET.tostring(
            workbook_xml, encoding="UTF-8", xml_declaration=True
        )
        for path in ("xl/_rels/workbook.xml.rels", "[Content_Types].xml"):
            document = ET.fromstring(after.read(path))
            for element in list(document):
                if "calcChain" in element.get("Type", "") or "calcChain" in element.get(
                    "PartName", ""
                ):
                    document.remove(element)
            changes[path] = ET.tostring(
                document, encoding="UTF-8", xml_declaration=True
            )
        temporary = destination.with_suffix(".tmp.xlsm")
        with ZipFile(temporary, "w", ZIP_DEFLATED) as output:
            for element in after.infolist():
                if element.filename != "xl/calcChain.xml":
                    output.writestr(
                        element,
                        changes.get(element.filename, after.read(element.filename)),
                    )
    temporary.replace(destination)
