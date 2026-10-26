# -*- coding: utf-8 -*-
"""Modèle sectoriel EU / US.

STRUCTURE DU FICHIER
1. SectorModel : choisir le marché, lancer le calcul et lire les résultats.
2. run_model : dérouler les étapes du calcul pour un marché.
3. Sous-scores et piliers : transformer les mesures et les agréger.
4. Macro et taux : déterminer le contexte du mois.
5. Rangs et recommandations : compter les votes et choisir les secteurs.
6. Exports : écrire les tableaux de scores en CSV.

FLUX DES DONNÉES
Actualiser.bat -> actualiser_excel.main()
    -> local_config.py : chemins et paramètres
    -> data_utils : Excel et historique mensuel figé
    -> SectorModel.run() -> run_model()
         -> mesures -> sous-scores -> six piliers
         -> macro et taux -> rangs -> votes -> recommandations
    -> excel_output : valeurs, formats et graphiques dans les XLSM
    -> actualiser_excel : bilan de production

Exemple : SectorModel("EU").run_production()
Pour actualiser aussi les XLSM : lancer Actualiser.bat.
"""

from copy import deepcopy
from pathlib import Path
import math

import numpy as np
import pandas as pd

from local_config import EU_CONFIG, US_CONFIG, MACRO_EXCEL_FILE
from data_utils import (
    open_workbooks,
    load_history,
    save_history,
    read_benchmark_returns,
    read_sector_returns,
    read_sector_block,
    read_macro_data,
    read_us10y,
    normalize_monthly_index,
)

MODEL_CONFIGS = {"EU": EU_CONFIG, "US": US_CONFIG}

# 1. Interface du modèle


class SectorModel:
    """Configure, exécute et inspecte un modèle sectoriel."""

    def __init__(
        self,
        universe="EU",
        variables=None,
        excel_file=None,
        macro_file=None,
    ):
        universe = universe.upper()

        if universe not in MODEL_CONFIGS:
            raise ValueError("Universe must be EU or US.")

        configuration = MODEL_CONFIGS[universe]

        self.universe = universe
        self._base_config = deepcopy(configuration)
        self.config = deepcopy(self._base_config)
        self.excel_file = Path(excel_file or self.config["excel_file"])
        self.macro_file = Path(macro_file or MACRO_EXCEL_FILE)
        self.results = None

        self._available_variables = deepcopy(self.config["backtest_variables"])
        self._variable_selection = {
            pillar: values.copy()
            for pillar, values in self._available_variables.items()
        }

        if variables:
            self._apply_variable_selection(variables)

    # Production : lire les entrées, calculer, puis exporter.
    def run(self, force=False):
        """Exécute le modèle et conserve les résultats en mémoire."""
        if self.results is not None and not force:
            return self.results

        wb_sector, macro_workbook = open_workbooks(
            self.excel_file,
            self.macro_file,
        )

        load_history(reload=True)
        try:
            self.results = run_model(
                wb_sector,
                macro_workbook,
                self.config,
            )
            save_history()
        finally:
            wb_sector.close()
            macro_workbook.close()

        return self.results

    def run_production(
        self,
        output_dir="output",
        force=True,
    ):
        """Exécute le modèle et exporte les CSV et la feuille Output."""
        results = self.run(force=force)

        save_production(
            results,
            Path(output_dir),
            self.config,
        )
        from excel_output import save_excel

        save_excel(results, output_dir, self.config)
        print_latest(results)

        return results

    # Consultation : listes, définitions et tableaux de résultats.
    @property
    def sectors(self):
        """Retourne les secteurs du modèle."""
        return self.config["sectors"].copy()

    @property
    def pillars(self):
        """Retourne les piliers du modèle."""
        return list(self.config["base_weights"])

    @property
    def benchmark_name(self):
        """Retourne le nom du benchmark du marché."""
        return self.config["volatility"]["benchmark_name"]

    def available_variables(self, pillar=None):
        """Retourne les variables disponibles avant sélection."""
        if pillar is None:
            return deepcopy(self._available_variables)

        self._validate_pillar(pillar)
        return self._available_variables[pillar].copy()

    def variables(self, pillar=None):
        """Retourne les variables actuellement utilisées."""
        if pillar is None:
            return deepcopy(self._variable_selection)

        self._validate_pillar(pillar)
        return self._variable_selection[pillar].copy()

    def variable_definitions(self, pillar=None):
        """Présente la source et les règles des variables disponibles."""
        pillars = self.pillars if pillar is None else [pillar]

        if pillar is not None:
            self._validate_pillar(pillar)

        rows = []

        for pillar_name in pillars:
            active = set(self._variable_selection[pillar_name])
            historical = self._base_config["historical_variables"].get(pillar_name, {})

            for variable in self._available_variables[pillar_name]:
                if variable in historical:
                    definition = historical[variable]
                    window = self._base_config["history_window"]
                    sector_windows = definition.get(
                        "sector_windows",
                        {},
                    )

                    if sector_windows:
                        overrides = ", ".join(
                            f"{sector}={months}"
                            for sector, months in sector_windows.items()
                        )
                        window = f"{window}; {overrides}"

                    rows.append(
                        {
                            "universe": self.universe,
                            "pillar": pillar_name,
                            "variable": variable,
                            "active": variable in active,
                            "source": "historical",
                            "sheet": definition["sheet"],
                            "column": definition["column"],
                            "direction": (
                                "higher" if definition["rank_order"] == 1 else "lower"
                            ),
                            "window": window,
                            "cross_sectional_mix": definition[
                                "mix_cross_sectional_rank"
                            ],
                        }
                    )
                    continue

                if pillar_name == "Momentum":
                    cfg = self._base_config["momentum"]

                    if variable == "momentum_6m_1m":
                        column = cfg["price_column"]
                        window = f"{cfg['short_horizon']}M-1M"
                        source = "price"
                    elif variable == "momentum_12m_1m":
                        column = cfg["price_column"]
                        window = f"{cfg['long_horizon']}M-1M"
                        source = "price"
                    else:
                        column = " / ".join(
                            [
                                cfg["revision_up_column"],
                                cfg["revision_down_column"],
                                cfg["revision_unchanged_column"],
                            ]
                        )
                        window = None
                        source = "earnings revisions"

                    rows.append(
                        {
                            "universe": self.universe,
                            "pillar": pillar_name,
                            "variable": variable,
                            "active": variable in active,
                            "source": source,
                            "sheet": cfg["sheet"],
                            "column": column,
                            "direction": "higher",
                            "window": window,
                            "cross_sectional_mix": True,
                        }
                    )
                    continue

                cfg = self._base_config["volatility"]
                window = (
                    cfg["volatility_offset"] + 1
                    if variable == "volatility_6m"
                    else cfg["offset_downside"] + 1
                )

                rows.append(
                    {
                        "universe": self.universe,
                        "pillar": pillar_name,
                        "variable": variable,
                        "active": variable in active,
                        "source": "sector returns",
                        "sheet": cfg["sheet_returns"],
                        "column": "derived",
                        "direction": "lower",
                        "window": f"{window} observations",
                        "cross_sectional_mix": False,
                    }
                )

        return pd.DataFrame(rows)

    def composition(self, pillar):
        """Présente les variables réellement utilisées par secteur."""
        self._validate_pillar(pillar)

        composition = self.config["pillar_composition"][pillar]
        default = composition["default"]
        by_sector = composition.get(
            "by_sector",
            {},
        )

        rows = []

        for sector in self.sectors:
            variables = by_sector.get(
                sector,
                default,
            )

            rows.append(
                {
                    "sector": sector,
                    "variables": " | ".join(variables),
                }
            )

        return pd.DataFrame(rows)

    def summary(self):
        """Présente la composition active du modèle."""
        rows = []

        for pillar in self.pillars:
            rows.append(
                {
                    "universe": self.universe,
                    "pillar": pillar,
                    "weight": self.config["base_weights"][pillar],
                    "variables": " | ".join(self._variable_selection[pillar]),
                }
            )

        return pd.DataFrame(rows)

    def variable_values(self, variable, date=None):
        """Retourne les scores sectoriels d'une variable pour une date."""
        results = self._ensure_results()

        if variable not in results["variable_scores"]:
            raise ValueError(f"Unknown variable: {variable}")

        score = results["variable_scores"][variable]
        selected_date = self._select_date(score.index, date)

        return score.loc[selected_date].rename(variable)

    def pillar_values(self, pillar, date=None):
        """Retourne les scores sectoriels d'un pilier pour une date."""
        self._validate_pillar(pillar)
        results = self._ensure_results()
        score = results["pillars"][pillar]
        selected_date = self._select_date(score.index, date)

        return score.loc[selected_date].rename(pillar)

    def variable_table(self, pillar, date=None):
        """Compare les variables actives et le score final d'un pilier."""
        self._validate_pillar(pillar)
        results = self._ensure_results()

        variables = self._variable_selection[pillar]
        pillar_score = results["pillars"][pillar]
        selected_date = self._select_date(
            pillar_score.index,
            date,
        )

        table = pd.DataFrame(index=self.sectors)

        for variable in variables:
            score = results["variable_scores"][variable]

            if selected_date in score.index:
                table[variable] = score.loc[selected_date]
            else:
                table[variable] = pd.NA

        table[pillar] = pillar_score.loc[selected_date]
        table.index.name = "sector"

        return table

    def recommendations(self, date=None):
        """Retourne les recommandations finales pour une date."""
        history = self._ensure_results()["history"]

        if history.empty:
            return history.copy()

        dates = pd.DatetimeIndex(history["date"].drop_duplicates()).sort_values()
        selected_date = self._select_date(dates, date)

        return (
            history[history["date"] == selected_date]
            .sort_values(
                "rang_global",
                ascending=False,
            )
            .reset_index(drop=True)
        )

    def save_outputs(self, output_dir="output"):
        """Sauvegarde les sorties du modèle."""
        results = self._ensure_results()

        save_all_outputs(
            results,
            Path(output_dir),
            self.config,
        )

    def show_latest(self):
        """Affiche la dernière recommandation disponible."""
        print_latest(self._ensure_results())

    # Recherche : changer les variables retenues dans un pilier.
    def _apply_variable_selection(self, selections):
        """Applique les variables choisies aux compositions des piliers."""
        for pillar, variables in selections.items():
            self._validate_pillar(pillar)

            if variables is None:
                continue

            variables = list(dict.fromkeys(variables))

            if not variables:
                raise ValueError(f"{pillar} must contain at least one variable.")

            available = self._available_variables[pillar]
            unknown = [variable for variable in variables if variable not in available]

            if unknown:
                raise ValueError(f"Unknown variables for {pillar}: {unknown}")

            self._update_pillar_composition(
                pillar,
                variables,
            )
            self._variable_selection[pillar] = variables
            self.config["backtest_variables"][pillar] = variables.copy()

            if pillar in self.config["historical_variables"]:
                definitions = self.config["historical_variables"][pillar]
                self.config["historical_variables"][pillar] = {
                    variable: definitions[variable] for variable in variables
                }

    def _update_pillar_composition(self, pillar, variables):
        """Filtre la composition existante avec la sélection explicite."""
        base_composition = self._base_config["pillar_composition"][pillar]
        base_default = base_composition["default"]
        base_by_sector = base_composition.get(
            "by_sector",
            {},
        )

        default = [variable for variable in base_default if variable in variables]

        if not default:
            raise ValueError(
                f"No selected variable is valid for the default "
                f"{pillar} composition."
            )

        by_sector = {}

        for sector, sector_variables in base_by_sector.items():
            selected = [
                variable for variable in sector_variables if variable in variables
            ]

            if not selected:
                raise ValueError(
                    f"No selected variable is valid for " f"{pillar} / {sector}."
                )

            by_sector[sector] = selected

        self.config["pillar_composition"][pillar] = {
            "default": default,
        }

        if by_sector:
            self.config["pillar_composition"][pillar]["by_sector"] = by_sector

    def _validate_pillar(self, pillar):
        if pillar not in self._available_variables:
            raise ValueError(f"Unknown pillar: {pillar}")

    def _ensure_results(self):
        if self.results is None:
            self.run()

        return self.results

    @staticmethod
    def _select_date(index, date):
        dates = pd.DatetimeIndex(index).sort_values()

        if len(dates) == 0:
            raise ValueError("No date available.")

        if date is None:
            return dates[-1]

        target = pd.Timestamp(date).to_period("M").to_timestamp("M")
        available = dates[dates <= target]

        if len(available) == 0:
            raise ValueError(f"No date available before {target.date()}.")

        return available[-1]


# 2. Pipeline de calcul


def run_model(sector_workbook, macro_workbook, config):
    """Exécute les calculs communs du marché."""
    raw_variables = {}
    pillars, variable_scores = compute_historical_pillars(
        sector_workbook,
        config,
        details=raw_variables,
    )

    momentum, momentum_scores = compute_momentum(
        sector_workbook,
        config,
        details=raw_variables,
    )
    volatility, volatility_scores, returns = compute_volatility(
        sector_workbook,
        config,
        details=raw_variables,
    )
    benchmark_returns = read_benchmark_returns(
        sector_workbook[config["volatility"]["sheet_returns"]],
        volatility_config=config["volatility"],
        key_prefix=config["history_prefix"],
    )

    pillars["Momentum"] = momentum
    pillars["Volatility"] = volatility

    variable_scores.update(momentum_scores)
    variable_scores.update(volatility_scores)

    macro_context = build_macro_context(
        sector_workbook,
        macro_workbook,
        config,
    )

    (
        history,
        pillars,
        ranks,
        macro_context,
    ) = compute_final_results(
        pillars,
        macro_context,
        config,
    )

    return {
        "history": history,
        "pillars": pillars,
        "ranks": ranks,
        "variable_scores": variable_scores,
        "variable_raw": raw_variables,
        "benchmark_weights": read_sector_block(
            sector_workbook["Bench"],
            "B",
            sectors=config["sectors"],
            key_prefix=config["history_prefix"],
            start_row=2,
        ),
        "returns": returns,
        "benchmark_returns": benchmark_returns,
        "macro_context": macro_context,
    }


# 3. Sous-scores et piliers


def excel_rank(value, values, order):
    """Reproduit RANK Excel. ordre=1 croissant, ordre=0 décroissant."""
    if pd.isna(value):
        return np.nan

    series = pd.Series(values, dtype="float64").dropna()
    if series.empty:
        return np.nan

    if order == 1:
        return 1 + int((series < value).sum())

    return 1 + int((series > value).sum())


def cross_sectional_rank(row_number, order, scale_0_10=True):
    """Classe les secteurs au même mois."""
    result = pd.Series(index=row_number.index, dtype=float)
    n_sectors = len(row_number.index)

    for sector in row_number.index:
        value = row_number[sector]

        if pd.isna(value):
            result[sector] = np.nan
            continue

        rank = excel_rank(value, row_number.values, order)

        if scale_0_10:
            result[sector] = rank / n_sectors * 10
        else:
            result[sector] = rank

    return result


def rank_sectors(row_number):
    """Classe les secteurs de 1 = Worst à N = Best."""
    result = pd.Series(index=row_number.index, dtype=float)

    for sector in row_number.index:
        value = row_number[sector]

        if pd.isna(value):
            result[sector] = np.nan
            continue

        result[sector] = excel_rank(
            value,
            row_number.values,
            order=1,
        )

    return result


def percentrank_inc(values, x):
    """Reproduit PERCENTRANK.INC d'Excel."""
    values = pd.Series(values, dtype=float).dropna().sort_values().to_numpy()

    if len(values) == 0 or pd.isna(x):
        return np.nan
    if len(values) == 1:
        return 1.0
    if x <= values[0]:
        return 0.0
    if x >= values[-1]:
        return 1.0

    positions = np.where(np.isclose(values, x, rtol=0, atol=1e-12))[0]

    if len(positions):
        return positions[0] / (len(values) - 1)

    right = np.searchsorted(values, x, side="right")
    left = right - 1
    x0 = values[left]
    x1 = values[right]
    fraction = (x - x0) / (x1 - x0)

    return (left + fraction) / (len(values) - 1)


def compute_deviation_from_mean(raw, exclude_financials_from_mean=False):
    """Calcule l'écart du secteur à la moyenne du même mois."""
    diff = pd.DataFrame(
        index=raw.index,
        columns=raw.columns,
        dtype=float,
    )

    for date, row_number in raw.iterrows():
        base = row_number.copy()

        if exclude_financials_from_mean:
            base = base.drop("Fin")

        mean = base.mean(skipna=True)
        diff.loc[date] = row_number - mean

    return diff


def compute_historical_score(
    diff,
    rank_order,
    history_window,
    sector_windows=None,
):
    """Classe l'écart courant dans son historique mensuel."""
    sector_windows = sector_windows or {}

    score = pd.DataFrame(
        index=diff.index,
        columns=diff.columns,
        dtype=float,
    )

    for i in range(len(diff)):
        for sector in diff.columns:
            value = diff.iloc[i][sector]

            if pd.isna(value):
                continue

            n_months = sector_windows.get(
                sector,
                history_window,
            )
            n_observations = n_months + 1
            end = i + n_observations

            if end > len(diff):
                continue

            values = diff.iloc[i:end][sector]
            rank = excel_rank(
                value,
                values.values,
                rank_order,
            )

            score.iloc[
                i,
                score.columns.get_loc(sector),
            ] = (
                10 * rank / n_observations
            )

    return score


def compute_variable_score(raw, variable_config, history_window):
    """Calcule le score d'une sous-variable."""
    diff = compute_deviation_from_mean(
        raw,
        variable_config["exclude_financials_from_mean"],
    )

    score_hist = compute_historical_score(
        diff,
        variable_config["rank_order"],
        history_window,
        variable_config.get("sector_windows"),
    )

    if not variable_config["mix_cross_sectional_rank"]:
        return score_hist

    score = score_hist.copy()

    for date in score.index:
        score_cross = cross_sectional_rank(
            raw.loc[date],
            variable_config["rank_order"],
            scale_0_10=True,
        )
        score.loc[date] = (score_hist.loc[date] + score_cross) / 2

    return score


def compose_pillar(scores, config, pillar):
    """Agrège les sous-scores selon la composition du pilier."""
    sectors = config["sectors"]
    composition = config["pillar_composition"][pillar]

    dates = None
    for score in scores.values():
        dates = score.index if dates is None else dates.intersection(score.index)

    if dates is None or len(dates) == 0:
        raise ValueError(f"Aucune date disponible pour le pilier {pillar}.")

    result = pd.DataFrame(
        index=dates,
        columns=sectors,
        dtype=float,
    )

    by_sector = composition.get("by_sector", {})

    for date in dates:
        for sector in sectors:
            variables = by_sector.get(
                sector,
                composition["default"],
            )
            values = [scores[variable].at[date, sector] for variable in variables]

            if any(pd.isna(value) for value in values):
                result.at[date, sector] = np.nan
            else:
                result.at[date, sector] = sum(values) / len(values)

    decimals = config["pillar_rounding"].get(pillar)

    if decimals is not None:
        result = result.round(decimals)

    return result


def compute_historical_pillars(wb, config, details=None):
    """Calcule Leverage, Margin, Value et Growth."""
    sectors = config["sectors"]
    variable_scores = {}
    pillars = {}

    for pillar, variables in config["historical_variables"].items():
        pillar_scores = {}

        for variable_name, cfg in variables.items():
            ws = wb[cfg["sheet"]]
            raw = read_sector_block(
                ws,
                cfg["column"],
                sectors=sectors,
                key_prefix=config["history_prefix"],
            )

            score = compute_variable_score(
                raw,
                cfg,
                config["history_window"],
            )

            if details is not None:
                details[variable_name] = raw

            variable_scores[variable_name] = score
            pillar_scores[variable_name] = score

        pillars[pillar] = compose_pillar(
            pillar_scores,
            config,
            pillar,
        )

    return pillars, variable_scores


def compute_momentum(wb, config, details=None):
    """Calcule les trois composantes Momentum."""
    cfg = config["momentum"]
    sectors = config["sectors"]
    ws = wb[cfg["sheet"]]

    def read_data(column):
        return read_sector_block(
            ws,
            column,
            sectors=sectors,
            key_prefix=config["history_prefix"],
        )

    prices = read_data(cfg["price_column"])
    up = read_data(cfg["revision_up_column"])
    down = read_data(cfg["revision_down_column"])
    unchanged = read_data(cfg["revision_unchanged_column"])

    short_horizon = cfg["short_horizon"]
    long_horizon = cfg["long_horizon"]
    current_month_sector = cfg.get("current_month_sector")

    short_term_returns = pd.DataFrame(
        index=prices.index,
        columns=sectors,
        dtype=float,
    )
    long_term_returns = pd.DataFrame(
        index=prices.index,
        columns=sectors,
        dtype=float,
    )

    for i in range(len(prices)):
        if i + short_horizon < len(prices):
            short_term_returns.iloc[i] = (
                prices.iloc[i + 1] / prices.iloc[i + short_horizon] - 1
            )

        if i + long_horizon < len(prices):
            long_term_returns.iloc[i] = (
                prices.iloc[i + 1] / prices.iloc[i + long_horizon] - 1
            )

        if current_month_sector:
            if i + short_horizon < len(prices):
                short_term_returns.at[
                    prices.index[i],
                    current_month_sector,
                ] = (
                    prices.iloc[i][current_month_sector]
                    / prices.iloc[i + short_horizon][current_month_sector]
                    - 1
                )

            if i + long_horizon < len(prices):
                long_term_returns.at[
                    prices.index[i],
                    current_month_sector,
                ] = (
                    prices.iloc[i][current_month_sector]
                    / prices.iloc[i + long_horizon][current_month_sector]
                    - 1
                )

    revision_ratio = (up - down) / (up + down + unchanged)

    if details is not None:
        details.update(
            {
                "momentum_6m_1m": short_term_returns,
                "momentum_12m_1m": long_term_returns,
                "earnings_revision_ratio": revision_ratio,
            }
        )

    scale_scores = cfg.get("scale_0_10", True)

    short_term_score = short_term_returns.apply(
        lambda row_number: cross_sectional_rank(
            row_number,
            order=1,
            scale_0_10=scale_scores,
        ),
        axis=1,
    )
    long_term_score = long_term_returns.apply(
        lambda row_number: cross_sectional_rank(
            row_number,
            order=1,
            scale_0_10=scale_scores,
        ),
        axis=1,
    )
    revision_score = revision_ratio.apply(
        lambda row_number: cross_sectional_rank(
            row_number,
            order=1,
            scale_0_10=scale_scores,
        ),
        axis=1,
    )

    variable_scores = {
        "momentum_6m_1m": short_term_score,
        "momentum_12m_1m": long_term_score,
        "earnings_revision_ratio": revision_score,
    }

    momentum = compose_pillar(
        variable_scores,
        config,
        "Momentum",
    )

    return momentum, variable_scores


def compute_volatility(wb, config, details=None):
    """Calcule total volatility et downside volatility."""
    cfg = config["volatility"]
    sectors = config["sectors"]

    returns = read_sector_returns(
        wb[cfg["sheet_returns"]],
        volatility_config=cfg,
        sectors=sectors,
        key_prefix=config["history_prefix"],
    )

    total_volatility = pd.DataFrame(
        index=returns.index,
        columns=sectors,
        dtype=float,
    )
    downside_volatility = pd.DataFrame(
        index=returns.index,
        columns=sectors,
        dtype=float,
    )

    n_obs_vol = cfg["volatility_offset"] + 1
    n_obs_downside = cfg["offset_downside"] + 1

    for i in range(len(returns)):
        if i + n_obs_vol <= len(returns):
            window = returns.iloc[i : i + n_obs_vol]
            total_volatility.iloc[i] = window.std(ddof=1) * math.sqrt(12)

        if i + n_obs_downside <= len(returns):
            window = returns.iloc[i : i + n_obs_downside]

            for sector in sectors:
                negative_returns = window[sector][window[sector] < 0].dropna()

                if len(negative_returns) >= 2:
                    downside_volatility.at[
                        returns.index[i],
                        sector,
                    ] = negative_returns.std(ddof=1) * math.sqrt(12)

    if details is not None:
        details.update(
            {
                "volatility_6m": total_volatility,
                "downside_volatility_18m": downside_volatility,
            }
        )

    score_config = {
        "rank_order": 0,
        "exclude_financials_from_mean": False,
        "mix_cross_sectional_rank": False,
    }

    score_total = compute_variable_score(
        total_volatility,
        score_config,
        config["history_window"],
    )
    score_downside = compute_variable_score(
        downside_volatility,
        score_config,
        config["history_window"],
    )

    variable_scores = {
        "volatility_6m": score_total,
        "downside_volatility_18m": score_downside,
    }

    volatility = compose_pillar(
        variable_scores,
        config,
        "Volatility",
    )

    return volatility, variable_scores, returns


# 4. Contexte macro et taux


def compute_rate_signal(us10y_rates, signal_config):
    """Calcule le rate overlay à partir du US 10Y."""
    alpha = signal_config["alpha_ewma"]
    threshold = signal_config["percentile_threshold"]

    df = us10y_rates.copy().sort_index()
    df["ewma_us10y"] = np.nan
    df["diff_ewma"] = np.nan
    df["percentile_taux"] = np.nan
    df["signal_taux"] = None

    for i in range(len(df)):
        rate = df.iloc[i]["us10y"]

        if pd.isna(rate):
            continue

        if i == 0 or pd.isna(df.iloc[i - 1]["ewma_us10y"]):
            ewma = rate
        else:
            previous_ewma = df.iloc[i - 1]["ewma_us10y"]
            ewma = alpha * previous_ewma + (1 - alpha) * rate

        df.iloc[i, df.columns.get_loc("ewma_us10y")] = ewma

        diff = rate - ewma
        df.iloc[i, df.columns.get_loc("diff_ewma")] = diff

        history = df["diff_ewma"].iloc[: i + 1]
        percentile = percentrank_inc(history, diff)

        df.iloc[i, df.columns.get_loc("percentile_taux")] = percentile
        df.iloc[i, df.columns.get_loc("signal_taux")] = (
            "On" if percentile > threshold else "Off"
        )

    return df


def build_macro_context(sector_workbook, macro_workbook, config):
    """Assemble macro score, régime et rate overlay."""
    macro = read_macro_data(
        macro_workbook,
        macro_config=config["macro"],
        regime_weights=config["regime_weights"],
        key_prefix=config["history_prefix"],
    )

    rate = read_us10y(
        sector_workbook,
        signal_config=config["rate_signal"],
        key_prefix=config["history_prefix"],
    )

    rate_signal = compute_rate_signal(
        rate,
        config["rate_signal"],
    )

    context = macro.join(
        rate_signal[["signal_taux"]],
        how="outer",
    )

    return context.sort_index()


# 5. Rangs, votes et recommandations


def align_pillars(pillars):
    """Aligne les piliers sur une clé mensuelle commune."""
    normalized_pillars = {
        name: normalize_monthly_index(df) for name, df in pillars.items()
    }

    dates = None

    for df in normalized_pillars.values():
        dates = df.index if dates is None else dates.intersection(df.index)

    if dates is None or len(dates) == 0:
        raise ValueError("Aucune date mensuelle commune entre les piliers.")

    dates = dates.sort_values(ascending=False)

    return {name: df.loc[dates].copy() for name, df in normalized_pillars.items()}


def compute_pillar_ranks(pillars):
    """Classe les secteurs dans chaque pilier."""
    ranks = {}

    for name, df in pillars.items():
        rank_frame = pd.DataFrame(
            index=df.index,
            columns=df.columns,
            dtype=float,
        )

        for date in df.index:
            rank_frame.loc[date] = rank_sectors(df.loc[date])

        ranks[name] = rank_frame

    return ranks


def compute_weighted_score(ranks, weights, date):
    """Agrège les rangs des piliers avec les poids fournis."""
    sectors = next(iter(ranks.values())).columns

    result = pd.Series(
        0.0,
        index=sectors,
    )
    valid = pd.Series(
        True,
        index=sectors,
    )

    for pillar, pillar_weight in weights.items():
        if pillar_weight == 0:
            continue
        series = ranks[pillar].loc[date]
        valid &= series.notna()
        result += series.fillna(0) * pillar_weight

    result[~valid] = np.nan

    return result


def select_top_worst(
    top_count,
    bottom_count,
    overall_rank,
    n_top,
    n_worst,
):
    """Sélectionne les secteurs Positive et Negative."""
    table = pd.DataFrame(
        {
            "top_count": top_count,
            "bottom_count": bottom_count,
            "rang_global": overall_rank,
        }
    )

    top_candidates = table[table["bottom_count"] < n_worst].copy()

    top_candidates = top_candidates.sort_values(
        [
            "top_count",
            "rang_global",
            "bottom_count",
        ],
        ascending=[
            False,
            False,
            True,
        ],
    )

    top = list(top_candidates.head(n_top).index)

    worst_candidates = table[table["top_count"] < n_top].copy()

    worst_candidates = worst_candidates.sort_values(
        [
            "bottom_count",
            "rang_global",
            "top_count",
        ],
        ascending=[
            False,
            True,
            True,
        ],
    )

    worst = list(worst_candidates.head(n_worst).index)

    return top, worst


def compute_macro_adjusted_score(ranks, tilt_rank, date, config):
    """Agrège les rangs avec la pondération macro du marché."""
    sectors = config["sectors"]
    weights = config["macro_adjusted_weights"]

    score = pd.Series(
        0.0,
        index=sectors,
    )
    valid = pd.Series(
        True,
        index=sectors,
    )

    for pillar, pillar_weight in weights.items():
        if pillar == "Macro" or pillar_weight == 0:
            continue

        series = ranks[pillar].loc[date]
        valid &= series.notna()
        score += series.fillna(0) * pillar_weight

    macro_weight = weights.get("Macro", 0)

    if macro_weight:
        valid &= tilt_rank.notna()
        score += tilt_rank.fillna(0) * macro_weight

    score[~valid] = np.nan

    return score


def compute_final_results(pillars, macro_context, config):
    """Construit scores finaux, votes et recommandations."""
    sectors = config["sectors"]
    n_sectors = len(sectors)
    n_top = config["n_top"]
    n_worst = config["n_worst"]

    pillars = align_pillars(pillars)
    ranks = compute_pillar_ranks(pillars)
    rows = []

    for date in pillars["Leverage"].index:
        if date not in macro_context.index:
            continue

        regime = macro_context.at[date, "cycle"]
        macro_score = macro_context.at[date, "macro_score"]
        rate_signal = macro_context.at[date, "signal_taux"]

        if regime not in config["regime_weights"]:
            continue
        if rate_signal not in {"On", "Off"}:
            continue

        base_score = compute_weighted_score(
            ranks,
            config["base_weights"],
            date,
        )
        base_rank = rank_sectors(base_score)

        tilt_score = compute_weighted_score(
            ranks,
            config["regime_weights"][regime],
            date,
        )
        tilt_rank = rank_sectors(tilt_score)

        macro_adjusted_score = compute_macro_adjusted_score(
            ranks,
            tilt_rank,
            date,
            config,
        )
        overall_rank = rank_sectors(macro_adjusted_score)

        top_count = pd.Series(
            0,
            index=sectors,
            dtype=int,
        )
        bottom_count = pd.Series(
            0,
            index=sectors,
            dtype=int,
        )

        for pillar in config["top_vote_pillars"]:
            r = ranks[pillar].loc[date]
            top_count += (r > n_sectors - n_top).fillna(False).astype(int)

        for pillar in config["bottom_vote_pillars"]:
            r = ranks[pillar].loc[date]
            bottom_count += (r <= n_worst).fillna(False).astype(int)

        macro_weight = config["macro_vote_weight"]

        top_count += macro_weight * (tilt_rank > n_sectors - n_top).fillna(
            False
        ).astype(int)

        bottom_count += macro_weight * (tilt_rank <= n_worst).fillna(False).astype(int)

        if rate_signal == "On":
            for pillar in config["rate_overlay_pillars"]:
                r = ranks[pillar].loc[date]

                top_count += (r > n_sectors - n_top).fillna(False).astype(int)

                bottom_count += (r <= n_worst).fillna(False).astype(int)

        for rule in config.get("no_vote_rules", []):
            start_date = (
                pd.Timestamp(rule["start_date"]).to_period("M").to_timestamp("M")
            )

            end_date = rule.get("end_date")
            enabled = date >= start_date

            if end_date:
                end_date = pd.Timestamp(end_date).to_period("M").to_timestamp("M")
                enabled = enabled and date <= end_date

            if enabled:
                sector = rule["sector"]
                top_count[sector] = 0
                bottom_count[sector] = 0

        top, worst = select_top_worst(
            top_count,
            bottom_count,
            overall_rank,
            n_top,
            n_worst,
        )

        for sector in sectors:
            if sector in top:
                recommendations = "Positive"
            elif sector in worst:
                recommendations = "Negative"
            else:
                recommendations = "Neutral"

            row_number = {
                "date": date,
                "secteur": sector,
                "cycle": regime,
                "macro_score": macro_score,
                "signal_taux": rate_signal,
                "score_base": base_score[sector],
                "rang_base": base_rank[sector],
                "score_tilt": tilt_score[sector],
                "rang_tilt_macro": tilt_rank[sector],
                "score_global_macro": macro_adjusted_score[sector],
                "rang_global": overall_rank[sector],
                "top_count": int(top_count[sector]),
                "bottom_count": int(bottom_count[sector]),
                "recommendation": recommendations,
            }

            for pillar in ranks:
                row_number[f"score_{pillar.lower()}"] = pillars[pillar].at[date, sector]
                row_number[f"rang_{pillar.lower()}"] = ranks[pillar].at[date, sector]

            rows.append(row_number)

    return (
        pd.DataFrame(rows),
        pillars,
        ranks,
        macro_context,
    )


# 6. Exports des scores


def save_production(results, output_dir, config):
    """Sauvegarde uniquement les sorties courantes du modèle."""
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)

    prefix = config["output_prefix"]
    history = results["history"]
    pillars = results["pillars"]
    ranks = results["ranks"]

    pillar_dates = None

    for df in pillars.values():
        pillar_dates = (
            df.index if pillar_dates is None else pillar_dates.intersection(df.index)
        )

    if pillar_dates is not None and len(pillar_dates):
        pillar_date = pillar_dates.max()
        rows = []

        for sector in config["sectors"]:
            row_number = {
                "date": pillar_date,
                "secteur": sector,
            }

            for pillar in config["base_weights"]:
                row_number[f"score_{pillar.lower()}"] = pillars[pillar].at[
                    pillar_date, sector
                ]
                row_number[f"rang_{pillar.lower()}"] = ranks[pillar].at[
                    pillar_date, sector
                ]

            rows.append(row_number)

        pd.DataFrame(rows).to_csv(
            directory / f"{prefix}_piliers_latest_available.csv",
            index=False,
        )

    if history.empty:
        return

    latest_date = history["date"].max()
    latest = history[history["date"] == latest_date].copy()
    latest = latest.sort_values(
        "rang_global",
        ascending=False,
    )

    latest.to_csv(
        directory / f"{prefix}_recommandations_latest.csv",
        index=False,
    )

    columns = [
        "date",
        "secteur",
        "score_leverage",
        "score_margin",
        "score_value",
        "score_momentum",
        "score_growth",
        "score_volatility",
        "macro_score",
        "cycle",
        "signal_taux",
        "rang_global",
        "recommendation",
    ]

    latest[columns].to_csv(
        directory / f"{prefix}_piliers_latest_with_reco.csv",
        index=False,
    )


def save_all_outputs(results, output_dir, config):
    """Sauvegarde l'historique complet et les sorties courantes."""
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)

    prefix = config["output_prefix"]

    results["history"].to_csv(
        directory / f"{prefix}_historique_modele.csv",
        index=False,
    )

    save_production(
        results,
        directory,
        config,
    )


def print_latest(results):
    """Affiche la dernière recommandation disponible."""
    history = results["history"]

    if history.empty:
        print("Aucun résultat calculé.")
        return

    latest_date = history["date"].max()
    latest = history[history["date"] == latest_date].copy()
    latest = latest.sort_values(
        "rang_global",
        ascending=False,
    )

    cols = [
        "secteur",
        "rang_global",
        "top_count",
        "bottom_count",
        "recommendation",
        "macro_score",
        "cycle",
        "signal_taux",
    ]

    print(f"\nDate de la recommandation finale : " f"{latest_date.date()}")
    print(latest[cols].to_string(index=False))

    pillar_dates = None
    for df in results["pillars"].values():
        pillar_dates = (
            df.index if pillar_dates is None else pillar_dates.intersection(df.index)
        )

    if pillar_dates is not None and len(pillar_dates):
        pillar_date = pillar_dates.max()

        if pillar_date > latest_date:
            print(
                f"\nAttention : les piliers sont disponibles jusqu'au "
                f"{pillar_date.date()}, mais le régime macro permet une "
                f"recommandation finale seulement jusqu'au {latest_date.date()}."
            )
