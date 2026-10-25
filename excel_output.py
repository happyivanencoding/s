# -*- coding: utf-8 -*-
"""Tableaux et présentation des classeurs sectoriels.

1. Valeurs : portefeuille, grille Output, détails et historiques.
2. Présentation : formats, liens et placement du Dashboard.
3. Écriture : remplir les classeurs avec Excel pour Windows.

Les scores sont déjà calculés par model_secto. Ce fichier les présente.
"""

from pathlib import Path
from datetime import datetime
from tempfile import TemporaryDirectory
import shutil

import numpy as np
import pandas as pd
from openpyxl.utils import column_index_from_string, get_column_letter
from local_config import PORTFOLIO_DELTA, PORTFOLIO_TILT
from model_secto import compute_deviation_from_mean
from data_utils import normalize_metadata, preserve_inputs


PILLAR_COLUMNS = {
    "EU": dict(
        zip(
            ["Leverage", "Margin", "Value", "Momentum", "Growth", "Volatility"],
            ["I", "S", "AC", "AM", "AU", "BE"],
        )
    ),
    "US": dict(
        zip(
            ["Leverage", "Margin", "Value", "Momentum", "Growth", "Volatility"],
            ["I", "P", "Z", "AJ", "AR", "BE"],
        )
    ),
}

# Score, mesure brute, historique.
DETAIL_COLUMNS = {
    "EU": {
        "net_debt_ebitda": ("J", "L", "K"),
        "fcf_total_debt": ("M", "O", "N"),
        "debt_equity": ("P", "R", "Q"),
        "operating_margin": ("T", "V", "U"),
        "net_margin": ("W", "Y", "X"),
        "ebitda_margin": ("Z", "AB", "AA"),
        "price_fcf": ("AD", "AF", "AE"),
        "ev_ebitda": ("AG", "AI", "AH"),
        "price_sales": ("AJ", "AL", "AK"),
        "momentum_6m_1m": ("AN", "AO", None),
        "momentum_12m_1m": ("AP", "AQ", None),
        "earnings_revision_ratio": ("AR", "AT", "AS"),
        "eps_growth": ("AV", "AX", "AW"),
        "sales_growth": ("AY", "BA", "AZ"),
        "ebitda_growth": ("BB", "BD", "BC"),
        "volatility_6m": ("BF", "BG", None),
        "downside_volatility_18m": ("BH", "BI", None),
    },
    "US": {
        "net_debt_ebitda": ("J", "L", "K"),
        "fcf_total_debt": ("M", "O", "N"),
        "operating_margin": ("Q", "S", "R"),
        "net_margin": ("T", "V", "U"),
        "ebitda_margin": ("W", "Y", "X"),
        "price_fcf": ("AA", "AC", "AB"),
        "ev_ebitda": ("AD", "AF", "AE"),
        "price_sales": ("AG", "AI", "AH"),
        "momentum_6m_1m": ("AK", "AL", None),
        "momentum_12m_1m": ("AM", "AN", None),
        "earnings_revision_ratio": ("AO", "AQ", "AP"),
        "eps_growth": ("AS", "AU", "AT"),
        "sales_growth": ("AV", "AX", "AW"),
        "ebitda_growth": ("AY", "BA", "AZ"),
        "cash_flow_growth": ("BB", "BD", "BC"),
        "volatility_6m": ("BF", "BG", None),
        "downside_volatility_18m": ("BH", "BI", None),
    },
}


# 1. Valeurs à écrire dans Output


def compute_portfolio(results, end_date):
    """Applique les recommandations du mois précédent aux poids du benchmark."""
    history = results["history"]
    valid_dates = history.loc[history["score_global_macro"].notna(), "date"].unique()
    recommendation_history = (
        history[history["date"].isin(valid_dates)]
        .pivot(index="date", columns="secteur", values="recommendation")
        .sort_index()
    )
    weights = results["benchmark_weights"].sort_index()
    returns = results["returns"].sort_index()
    benchmark = results["benchmark_returns"]["benchmark"]
    rows = []
    for date, sector_returns in returns.loc[:end_date].iterrows():
        signal = date - pd.offsets.MonthEnd(1)
        if signal not in recommendation_history.index:
            continue
        available_weights = weights.loc[:date]
        if available_weights.empty:
            continue
        recommendations = (
            recommendation_history.loc[signal]
            .reindex(sector_returns.index)
            .fillna("Neutral")
        )
        w = available_weights.iloc[-1].reindex(sector_returns.index)
        adjusted_weights = w.copy()
        positive, negative = recommendations.eq("Positive"), recommendations.eq(
            "Negative"
        )
        adjusted_weights[positive] = np.maximum(
            w[positive] * (1 + PORTFOLIO_TILT), w[positive] + PORTFOLIO_DELTA
        )
        adjusted_weights[negative] = np.minimum(
            w[negative] * (1 - PORTFOLIO_TILT),
            (w[negative] - PORTFOLIO_DELTA).clip(lower=0),
        )
        if w.isna().any() or sector_returns[adjusted_weights > 0].isna().any():
            portfolio = np.nan
        else:
            portfolio = (
                adjusted_weights * sector_returns
            ).sum() / adjusted_weights.sum()
        rows.append(
            [
                date,
                benchmark.get(date, np.nan),
                portfolio,
                sector_returns[positive].mean(),
                sector_returns[negative].mean(),
            ]
        )
    monthly = pd.DataFrame(
        rows, columns=["date", "benchmark", "portfolio", "top", "bottom"]
    ).set_index("date")
    if monthly.empty:
        raise ValueError("Aucun rendement de portefeuille disponible.")
    monthly["long_short"] = monthly["top"] - monthly["bottom"]
    cumulative = (1 + monthly).cumprod(skipna=False) * 100
    start = monthly.index[0] - pd.offsets.MonthEnd(1)
    cumulative.loc[start] = 100.0
    return monthly, cumulative.sort_index()


def output_layout(config):
    """Place les blocs avec deux lignes libres entre eux."""
    n = len(config["sectors"])
    sparkline_count = (
        sum(
            spark is not None for _, _, spark in DETAIL_COLUMNS[config["name"]].values()
        )
        * n
    )
    measure_count = sum(len(v) for v in config["backtest_variables"].values()) * n
    recommendations = 22
    history = recommendations + n + 3
    details = history + sparkline_count + 3
    curves = details + measure_count + 3
    return dict(
        recommendations=recommendations, history=history, details=details, curves=curves
    )


def output_cell(address, market):
    """Relie une cellule du scorecard à sa nouvelle position dans Output."""
    headers = {
        "BO4": "B2",
        "BP4": "B9",
        "BQ4": "B10",
        "BP7": "B11",
        "BV7": "B12",
        "BP6": "B11",
        "BV6": "B12",
        "BJ8": "B13",
        "BK10": "B6",
    }
    if address in headers:
        return headers[address]
    column = "".join(c for c in address if c.isalpha())
    row_number = int("".join(c for c in address if c.isdigit()))
    if column == "CD":
        start = 9 if market == "EU" else 10
        return f"B{15 + row_number - start}"
    return f"{get_column_letter(column_index_from_string(column) - 4)}{row_number - 8}"


def build_output(results, config):
    """Prépare les scores, les historiques et les séries des graphiques."""
    sectors = config["sectors"]
    market = config["name"]
    history = results["history"]
    if history.empty:
        raise ValueError("Aucune recommandation complète disponible.")
    date = history["date"].max()
    previous = date - pd.offsets.MonthEnd(1)
    current = history[history["date"].eq(date)].set_index("secteur")
    previous_data = history[history["date"].eq(previous)].set_index("secteur")
    monthly, cumulative = compute_portfolio(results, date)
    layout = output_layout(config)
    grid = pd.DataFrame(
        None,
        index=range(layout["curves"] + len(cumulative)),
        columns=range(74),
        dtype=object,
    )

    def set_cell_value(address, value):
        column = "".join(c for c in address if c.isalpha())
        row_number = int("".join(c for c in address if c.isdigit()))
        grid.iat[row_number - 1, column_index_from_string(column) - 1] = (
            "n.a." if pd.isna(value) else value
        )

    def scorecard(address, value):
        set_cell_value(output_cell(address, market), value)

    def write_block(row_number, column, values):
        values = pd.DataFrame(values).astype(object).where(pd.notna(values), None)
        grid.iloc[
            row_number - 1 : row_number - 1 + len(values),
            column - 1 : column - 1 + values.shape[1],
        ] = values.to_numpy()

    pillar_date = max(df.index.max() for df in results["pillars"].values())
    cycle = current["cycle"].iloc[0]
    rate_signal = current["signal_taux"].iloc[0]
    regime = {"C": "Contraction", "R": "Recovery", "E": "Expansion", "SD": "Slowdown"}[
        cycle
    ]
    next_month = date + pd.offsets.MonthBegin(1)
    months = [
        "Jan",
        "Fev",
        "Mar",
        "Avr",
        "Mai",
        "Juin",
        "Juil",
        "Aout",
        "Sept",
        "Oct",
        "Nov",
        "Dec",
    ]

    def month_label(d):
        return f"{months[d.month - 1]}-{d.year}"

    write_block(
        1,
        1,
        [
            ["Marché", market],
            ["Date modèle", date],
            ["Derniers piliers", pillar_date],
            ["Allocation", next_month],
            ["Cycle", cycle],
            ["Signal taux", rate_signal],
            ["Derniers poids benchmark", results["benchmark_weights"].index.max()],
            ["Derniers rendements", results["returns"].index.max()],
            ["Mois allocation", month_label(next_month)],
            ["Mois signal", month_label(date)],
            ["Classement courant", "Sectors Ranking " + month_label(next_month)],
            ["Classement précédent", "Sectors Ranking " + month_label(date)],
            ["Classement macro", "Macro rank - Currently in " + regime],
            ["Référence", "Bench" if market == "EU" else "CumRet_S&P500"],
            ["Top", monthly.at[date, "top"]],
            ["Short", -monthly.at[date, "bottom"]],
            ["Long / Short", monthly.at[date, "long_short"]],
            ["Portefeuille", monthly.at[date, "portfolio"]],
            ["Benchmark", monthly.at[date, "benchmark"]],
        ],
    )

    returns = results["returns"]
    bench = results["benchmark_returns"]["benchmark"]
    recommendation_codes = {"Negative": -1, "Neutral": 0, "Positive": 1}
    details = DETAIL_COLUMNS[market]
    scorecard("H9", "Secteur")
    for pillar, column in PILLAR_COLUMNS[market].items():
        scorecard(column + "9", pillar)
    for variable, (score, raw, _) in details.items():
        scorecard(score + "9", "Score " + variable.replace("_", " "))
        scorecard(raw + "9", variable.replace("_", " "))
    for col, title in {
        "BJ": "Macro",
        "BM": "Secteur",
        "BN": "Votes Top",
        "BO": "Votes Worst",
        "BP": "Score global",
        "BQ": "Reco",
        "BR": "Variation reco",
        "BU": "Votes Top précédents",
        "BV": "Votes Worst précédents",
        "BW": "Score précédent",
        "BX": "Reco précédente",
        "BY": month_label(date) + " EQ Performance",
        "BZ": month_label(date)
        + (
            " XSRet vs FTSE World Europe"
            if market == "EU"
            else " XSRet vs FTSE World US"
        ),
    }.items():
        scorecard(col + "9", title)
    for i, sector in enumerate(sectors):
        row_number = i + 10
        scorecard(f"H{row_number}", sector)
        scorecard(f"BM{row_number}", sector)
        for pillar, column in PILLAR_COLUMNS[market].items():
            rank = results["ranks"][pillar].at[date, sector]
            scorecard(f"{column}{row_number}", len(sectors) + 1 - rank)
        for variable, (score_col, raw_column, _) in details.items():
            score = results["variable_scores"].get(variable)
            raw = results["variable_raw"].get(variable)
            scorecard(
                f"{score_col}{row_number}",
                score.at[date, sector] if score is not None else np.nan,
            )
            scorecard(
                f"{raw_column}{row_number}",
                raw.at[date, sector] if raw is not None else np.nan,
            )
        row = current.loc[sector]
        scorecard(f"BJ{row_number}", len(sectors) + 1 - row["rang_tilt_macro"])
        for column, field in [
            ("BN", "top_count"),
            ("BO", "bottom_count"),
            ("BP", "score_global_macro"),
            ("BQ", "recommendation"),
        ]:
            scorecard(f"{column}{row_number}", row[field])
        for column, field in [
            ("BU", "top_count"),
            ("BV", "bottom_count"),
            ("BW", "score_global_macro"),
            ("BX", "recommendation"),
        ]:
            scorecard(
                f"{column}{row_number}",
                (
                    previous_data.at[sector, field]
                    if sector in previous_data.index
                    else np.nan
                ),
            )
        previous_recommendation = (
            previous_data.at[sector, "recommendation"]
            if sector in previous_data.index
            else None
        )
        scorecard(
            f"BR{row_number}",
            (
                recommendation_codes[row["recommendation"]]
                - recommendation_codes[previous_recommendation]
                if previous_recommendation
                else np.nan
            ),
        )
        ret = returns.at[date, sector]
        scorecard(f"BY{row_number}", ret)
        scorecard(f"BZ{row_number}", ret - bench.at[date])

    recommendation_dates = pd.date_range(end=date, periods=15, freq="ME")
    recommendations = history.pivot(
        index="secteur", columns="date", values="recommendation"
    )
    write_block(
        layout["recommendations"],
        4,
        [
            [market + " Reco"]
            + [d + pd.offsets.MonthBegin(1) for d in recommendation_dates]
        ],
    )
    for i, sector in enumerate(sectors):
        write_block(
            layout["recommendations"] + 1 + i,
            4,
            [
                [sector]
                + [
                    (
                        recommendation_codes.get(recommendations.at[sector, d], None)
                        if d in recommendations.columns
                        else None
                    )
                    for d in recommendation_dates
                ]
            ],
        )

    title = ("Europe" if market == "EU" else "US") + " equity sectors Performance"
    set_cell_value(f"W{layout['recommendations']}", title)
    write_block(layout["recommendations"] + 1, 22, [[""] + sectors])
    series = returns.loc[returns.index <= date].sort_index()
    horizons = [
        ("1M", 1),
        ("3M", 3),
        ("YTD", date.month),
        ("1Y", 12),
        ("3Y", 36),
        ("5Y", 60),
    ]
    for i, (name, period_count) in enumerate(horizons):
        perf = (1 + series.tail(period_count)).prod(min_count=period_count) - 1
        write_block(
            layout["recommendations"] + 2 + i,
            22,
            [[name] + [v if pd.notna(v) else "n.a." for v in perf]],
        )

    # Chaque colonne représente un mois, du plus récent au plus ancien.
    history_dates = pd.date_range(end=date, periods=60, freq="ME")[::-1]
    write_block(
        layout["history"], 4, [["Variable", "Secteur"] + history_dates.tolist()]
    )
    history_row = layout["history"] + 1
    for variable, (_, _, spark_col) in details.items():
        if spark_col is None:
            continue
        raw = results["variable_raw"].get(variable)
        if raw is None:
            raw = pd.DataFrame(index=history_dates, columns=sectors, dtype=float)
        definition = next(
            (
                variables[variable]
                for variables in config["historical_variables"].values()
                if variable in variables
            ),
            None,
        )
        deviation = (
            compute_deviation_from_mean(raw, definition["exclude_financials_from_mean"])
            if definition
            else raw
        )
        deviation = deviation.reindex(history_dates)
        for sector in sectors:
            write_block(
                history_row, 4, [[variable, sector] + deviation[sector].tolist()]
            )
            history_row += 1

    write_block(
        layout["details"],
        4,
        [["Date", "Secteur", "Pilier", "Variable", "Mesure", "Score"]],
    )
    rows = []
    for pillar, variables in config["backtest_variables"].items():
        for variable in variables:
            raw = results["variable_raw"].get(variable)
            score = results["variable_scores"].get(variable)
            for sector in sectors:
                rows.append(
                    [
                        date,
                        sector,
                        pillar,
                        variable,
                        raw.at[date, sector] if raw is not None else None,
                        score.at[date, sector] if score is not None else None,
                    ]
                )
    write_block(layout["details"] + 1, 4, rows)

    titles = [
        "Date",
        "CumRet_FTSE World Europe" if market == "EU" else "CumRet_S&P500",
        "CumRet_Sector deviations portfolio",
        "Top 3 Sectors" if market == "EU" else "CumRet_Top",
        "Worst 3 sectors" if market == "EU" else "CumRet_Bottom",
        "CumRet_Long Short",
    ]
    write_block(layout["curves"], 4, [titles])
    write_block(layout["curves"] + 1, 4, cumulative.reset_index().values.tolist())
    one_year = cumulative.loc[
        cumulative.index >= date - pd.offsets.MonthEnd(12),
        ["benchmark", "portfolio", "top", "bottom"],
    ]
    one_year = one_year / one_year.iloc[0] * 100
    write_block(
        layout["curves"], 13, [[titles[0], titles[1], titles[2], titles[3], titles[4]]]
    )
    write_block(layout["curves"] + 1, 13, one_year.reset_index().values.tolist())
    return grid, monthly, cumulative


def save_excel(results, output_dir, config):
    """Exporte la feuille de valeurs et les rendements mensuels."""
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    grid, monthly, _ = build_output(results, config)
    file_path = directory / f"{config['output_prefix']}_output.xlsx"
    grid.to_excel(file_path, sheet_name="Output", index=False, header=False)
    normalize_metadata(file_path)
    monthly.to_csv(
        directory / f"{config['output_prefix']}_performance_portefeuille.csv"
    )
    return file_path


# 2. Présentation du classeur


def excel_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return (
            pd.Timestamp(value) - pd.Timestamp("1899-12-30")
        ).total_seconds() / 86400
    return value


def append_recommendations(sheet, history, sectors):
    """Ajoute les nouveaux mois sans modifier les recommandations archivées."""
    last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
    last_date = pd.Timestamp("1899-12-30") + pd.Timedelta(
        days=sheet.Cells(last_row, 1).Value2
    )
    new_recommendations = history[history["date"] > last_date]
    rows = []
    for date, group in new_recommendations.groupby("date", sort=True):
        recommendations = group.set_index("secteur")["recommendation"].reindex(sectors)
        if recommendations.isna().any():
            raise ValueError(f"Recommandations incomplètes : {date:%d/%m/%Y}.")
        rows.append(
            [excel_value(date)]
            + recommendations.eq("Positive").astype(int).tolist()
            + recommendations.eq("Negative").astype(int).tolist()
        )
    if rows:
        width = 1 + 2 * len(sectors)
        end = get_column_letter(width)
        destination = sheet.Range(f"A{last_row+1}:{end}{last_row+len(rows)}")
        sheet.Range(f"A{last_row}:{end}{last_row}").Copy()
        destination.PasteSpecial(Paste=-4122)
        destination.Value2 = tuple(tuple(row_number) for row_number in rows)
    return len(rows)


def link_dashboard(workbook, config):
    """Relie le scorecard, les mini-historiques et les graphiques à Output."""
    market = config["name"]
    n = len(config["sectors"])
    layout = output_layout(config)
    dash = workbook.Worksheets("Dashboard")
    title_row = 6 if market == "US" else 7
    cells = [
        "BO4",
        "BP4",
        "BQ4",
        "BJ8",
        "BK10",
        "BY9",
        "BZ9",
        f"BP{title_row}",
        f"BV{title_row}",
    ]
    for column in PILLAR_COLUMNS[market].values():
        dash.Columns(column).Hidden = False
    columns = list(PILLAR_COLUMNS[market].values()) + [
        "BJ",
        "BN",
        "BO",
        "BP",
        "BQ",
        "BR",
        "BU",
        "BV",
        "BW",
        "BX",
        "BY",
        "BZ",
    ]
    for score, measure, _ in DETAIL_COLUMNS[market].values():
        columns += [score, measure]
    cells += [f"{c}{r}" for r in range(10, 10 + n) for c in columns]
    start = 9 if market == "EU" else 10
    cells += [f"CD{r}" for r in range(start, start + 5)]
    dash.Range("H26").Formula = '=TEXT(Output!B2,"[$-040C]mmm-yyyy")&" Commentaire : "'
    # Excel fournit le nom du format général dans sa langue courante.
    general_format = workbook.Application.GetInternational(26)
    for address in cells:
        cell = dash.Range(address)
        # Le format texte empêcherait Excel de calculer la formule de liaison.
        if cell.NumberFormat == "@":
            cell.MergeArea.NumberFormat = general_format
        cell.Formula = "=Output!" + output_cell(address, market)
    for prefix, columns in [
        (
            "Courbe",
            {
                "Date": "D",
                "Bench": "E",
                "Ptf": "F",
                "Top": "G",
                "Bottom": "H",
                "LS": "I",
            },
        ),
        ("Annee", {"Date": "M", "Bench": "N", "Ptf": "O", "Top": "P", "Bottom": "Q"}),
    ]:
        start = layout["curves"] + 1
        date_col = columns["Date"]
        for name, col in columns.items():
            formula = (
                f"=Output!${col}${start}:INDEX(Output!${col}:${col},"
                f"{start - 1}+COUNT(Output!${date_col}${start}:${date_col}$10000))"
            )
            workbook.Names(prefix + "_" + name).RefersTo = formula
    row_number = layout["recommendations"]
    dash.Pictures("Historique_recommandations").Formula = (
        f"=Output!D{row_number}:S{row_number+n}"
    )
    end = get_column_letter(22 + n)
    dash.Pictures("Performances_secteurs").Formula = (
        f"=Output!V{row_number}:{end}{row_number+7}"
    )


def compact_dashboard(workbook, config):
    """Rapproche les objets sans changer leur taille."""
    dash = workbook.Worksheets("Dashboard")
    left = dash.Range("H1").Left
    row_number = len(config["sectors"]) + 12
    objects = {c.Chart.ChartTitle.Text: c for c in dash.ChartObjects()}
    order = [
        dash.Shapes("Historique_recommandations"),
        dash.Shapes("Performances_secteurs"),
        next(
            c for title, c in objects.items() if "equity sectors Performance" in title
        ),
    ]
    for item in order:
        item.Placement = 3
        item.Left, item.Top = left, dash.Cells(row_number, 8).Top
        row_number = item.BottomRightCell.Row + 3
    cumulative = objects["Cumulative Returns benchmarked portfolio"]
    one_year = objects["Cumulative returns 1Y"]
    cumulative.Placement = one_year.Placement = 3
    cumulative.Left, cumulative.Top = left, dash.Cells(row_number, 8).Top
    one_year.Left, one_year.Top = left + cumulative.Width + 32, cumulative.Top
    objects["Cumulative Returns"].Placement = 2


def format_output(sheet, config, row_count):
    """Applique les formats de dates, nombres et titres des tableaux."""
    layout = output_layout(config)
    n = len(config["sectors"])
    sheet.Rows.Hidden = False
    sheet.Columns("A").ColumnWidth = 26
    sheet.Columns("B").ColumnWidth = 34
    sheet.Columns("C").ColumnWidth = 3
    sheet.Columns("D").ColumnWidth = 27
    sheet.Columns("E:BV").ColumnWidth = 12
    sheet.Columns("E").ColumnWidth = 18
    sheet.Range(f"A1:BV{row_count}").VerticalAlignment = -4108
    sheet.Range("D1:BV1").Font.Bold = True
    sheet.Range("D1:BV1").WrapText = True
    sheet.Rows(1).RowHeight = 60
    sheet.Range("B2:B4").NumberFormat = "dd/mm/yyyy"
    sheet.Range("B7:B8").NumberFormat = "dd/mm/yyyy"
    sheet.Range("B9:B10").NumberFormat = "@"
    sheet.Range("B15:B19").NumberFormat = "0.00%"
    sheet.Range(f"BU2:BV{n+1}").NumberFormat = "0.00%"
    r = layout["recommendations"]
    sheet.Range(f"E{r}:S{r}").NumberFormat = "[$-040C]mmm-yyyy"
    sheet.Range(f"W{r}:{get_column_letter(22+n)}{r}").HorizontalAlignment = 7
    sheet.Range(f"W{r+2}:{get_column_letter(22+n)}{r+7}").NumberFormat = "0.00%"
    h, d, c = layout["history"], layout["details"], layout["curves"]
    for row_number, end in [(h, "BM"), (d, "I"), (c, "I")]:
        sheet.Range(f"D{row_number}:{end}{row_number}").Font.Bold = True
        sheet.Range(f"D{row_number}:{end}{row_number}").WrapText = True
        sheet.Rows(row_number).RowHeight = 32
    sheet.Range(f"F{h}:BM{h}").NumberFormat = "[$-040C]mmm-yy"
    sheet.Range(f"F{h+1}:BM{d-3}").NumberFormat = "0.00"
    sheet.Range(f"D{d+1}:D{c-3}").NumberFormat = "dd/mm/yyyy"
    sheet.Range(f"D{c+1}:D{row_count}").NumberFormat = "[$-040C]mmm-yy"
    sheet.Range(f"M{c+1}:M{c+13}").NumberFormat = "[$-040C]mmm-yy"
    sheet.Range(f"E{c+1}:I{row_count}").NumberFormat = "0.00"
    sheet.Range(f"N{c+1}:Q{c+13}").NumberFormat = "0.00"
    sheet.Range(f"M{c}:Q{c}").Font.Bold = True
    sheet.Range(f"M{c}:Q{c}").WrapText = True


def fill_workbook(workbook, results, config, export_file):
    """Écrit les valeurs, ajoute les archives et rafraîchit la présentation."""
    # 1. Remplacer les valeurs d'Output, sans toucher aux feuilles d'entrée.
    workbook.Activate()
    grid, _, _ = build_output(results, config)
    out = workbook.Worksheets("Output")
    out.UsedRange.ClearContents()
    out.Range("B9:B10").NumberFormat = "@"
    values = tuple(tuple(excel_value(v) for v in row) for row in grid.to_numpy())
    address = f"A1:{get_column_letter(grid.shape[1])}{len(grid)}"
    out.Range(address).Value2 = values
    expected_date = excel_value(results["history"]["date"].max())
    if out.Range("B2").Value2 != expected_date or out.Range("D1").Value2 != "Secteur":
        raise RuntimeError("Les valeurs n'ont pas été écrites dans Output.")
    # 2. Formater les tableaux et ajouter seulement les nouveaux mois archivés.
    format_output(out, config, len(grid))
    added_months = append_recommendations(
        workbook.Worksheets("Reco_histo"), results["history"], config["sectors"]
    )
    # 3. Mettre à jour les liens, les courbes et la position des objets.
    link_dashboard(workbook, config)
    compact_dashboard(workbook, config)
    dash = workbook.Worksheets("Dashboard")
    out.Calculate()
    dash.Calculate()
    for item in dash.ChartObjects():
        item.Chart.Refresh()
    out.Activate()
    window = workbook.Application.ActiveWindow
    window.FreezePanes = False
    out.Range("E2").Select()
    window.FreezePanes = True
    window.DisplayGridlines = False
    window.Zoom = 85
    window.ScrollColumn = 1
    window.ScrollRow = 1
    # 4. Exporter le même tableau, avec ses largeurs et ses hauteurs.
    export = workbook.Application.Workbooks.Add()
    export_sheet = export.Worksheets(1)
    export_sheet.Name = "Output"
    source_range = out.Range(address)
    source_range.Copy(Destination=export_sheet.Range("A1"))
    source_range.Copy()
    export_sheet.Range("A1").PasteSpecial(Paste=8)
    for row_number in range(1, len(grid) + 1):
        export_sheet.Rows(row_number).RowHeight = out.Rows(row_number).RowHeight
    export_sheet.Activate()
    export_sheet.Range("E2").Select()
    workbook.Application.ActiveWindow.FreezePanes = True
    workbook.Application.ActiveWindow.DisplayGridlines = False
    export.SaveAs(str(export_file), FileFormat=51)
    export.Close(False)
    # 5. Enregistrer le classeur sur le Dashboard, avec les détails repliés.
    workbook.Activate()
    dash.Activate()
    dash.Outline.ShowLevels(ColumnLevels=1)
    dash.Range("H4").Select()
    window = workbook.Application.ActiveWindow
    window.ScrollColumn = 8
    window.ScrollRow = 1
    window.Zoom = 70
    window.DisplayGridlines = False
    workbook.Application.CutCopyMode = False
    workbook.Save()
    return added_months


# 3. Mise à jour des deux classeurs avec Excel pour Windows


def update_workbooks(models, output_dir):
    """Prépare les deux fichiers, puis remplace les originaux après succès."""
    import win32com.client

    added_months = {}
    # Les copies de travail sont temporaires et supprimées à la fin.
    with TemporaryDirectory(prefix="sector_excel_") as temporary:
        work_dir = Path(temporary)
        # Utiliser une instance séparée avec une interface Python stable.
        app = win32com.client.gencache.EnsureDispatch(
            win32com.client.DispatchEx("Excel.Application")
        )
        app.Visible = True
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        app.AskToUpdateLinks = False
        blank_workbook = app.Workbooks.Add()
        # Recalculer seulement les sorties ; la collecte est déjà enregistrée.
        app.Calculation = -4135
        app.CalculateBeforeSave = False
        try:
            for model in models:
                print(f"Mise à jour Excel {model.universe}...", flush=True)
                original = model.excel_file
                copy_path = work_dir / original.name
                shutil.copy2(original, copy_path)
                wb = app.Workbooks.Open(str(copy_path), UpdateLinks=0, ReadOnly=False)
                export = output_dir / f"{model.config['output_prefix']}_output.xlsx"
                try:
                    if wb.ReadOnly:
                        raise RuntimeError(
                            f"Classeur en lecture seule : {copy_path.name}"
                        )
                    added_months[model.universe] = fill_workbook(
                        wb, model.results, model.config, export
                    )
                finally:
                    wb.Close(False)
                # Conserver les formules d'entrée, les axes et les couleurs d'origine.
                preserve_inputs(original, copy_path)
                normalize_metadata(copy_path)
                normalize_metadata(export)
        finally:
            blank_workbook.Close(False)
            app.Quit()
        # Aucun original n'est remplacé tant que les deux copies ne sont pas prêtes.
        for model in models:
            shutil.copy2(work_dir / model.excel_file.name, model.excel_file)
    return added_months
