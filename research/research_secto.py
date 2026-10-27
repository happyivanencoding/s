"""Outils de recherche pour les modèles sectoriels EU et US."""

from copy import deepcopy
from pathlib import Path
import math
import sys

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Permettre le lancement direct depuis research/.
if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from model_secto import SectorModel, save_all_outputs, print_latest
from data_utils import save_table


def prepare_forward_returns(returns):
    """Aligne le rendement du mois suivant avec le signal courant."""
    return returns.sort_index().shift(-1)


def backtest_score(
    score,
    forward_returns,
    benchmark_returns=None,
    top_n=3,
    start=None,
    end=None,
):
    """Backteste un score sectoriel en pondération égale."""
    score = score.sort_index()
    dates = score.index.intersection(forward_returns.index).sort_values()

    if start:
        dates = dates[dates >= pd.Timestamp(start)]

    if end:
        dates = dates[dates <= pd.Timestamp(end)]

    rows = []

    for date in dates:
        signal = score.loc[date].dropna()
        returns = forward_returns.loc[date].dropna()

        sectors = signal.index.intersection(returns.index)
        signal = signal.loc[sectors]
        returns = returns.loc[sectors]

        if len(signal) < 2 * top_n:
            continue

        top = list(signal.nlargest(top_n).index)
        worst = list(signal.nsmallest(top_n).index)

        return_top = returns.loc[top].mean()
        return_worst = returns.loc[worst].mean()
        return_benchmark = np.nan

        if benchmark_returns is not None and date in benchmark_returns.index:
            return_benchmark = benchmark_returns.loc[date]

        rows.append(
            {
                "date_signal": date,
                "return_top": return_top,
                "return_worst": return_worst,
                "return_long_short": (return_top - return_worst),
                "return_benchmark": return_benchmark,
                "top_sectors": " | ".join(top),
                "worst_sectors": " | ".join(worst),
            }
        )

    return pd.DataFrame(rows)


def max_drawdown(returns):
    """Calcule le maximum drawdown."""
    if len(returns) == 0:
        return np.nan

    wealth = (1 + returns.fillna(0)).cumprod()
    drawdown = wealth / wealth.cummax() - 1

    return drawdown.min()


def statistics(backtest):
    """Calcule les statistiques principales du backtest."""
    if backtest.empty:
        return {
            "n_months": 0,
            "ann_return_top": np.nan,
            "ann_vol_top": np.nan,
            "sharpe_top": np.nan,
            "ann_return_ls": np.nan,
            "ann_vol_ls": np.nan,
            "sharpe_ls": np.nan,
            "win_rate_ls": np.nan,
            "max_drawdown_ls": np.nan,
        }

    top = backtest["return_top"].dropna()
    long_short = backtest["return_long_short"].dropna()

    def annual_return(series):
        if len(series) == 0:
            return np.nan

        return (1 + series).prod() ** (12 / len(series)) - 1

    def annual_volatility(series):
        if len(series) < 2:
            return np.nan

        return series.std(ddof=1) * math.sqrt(12)

    top_volatility = annual_volatility(top)
    long_short_volatility = annual_volatility(long_short)

    return {
        "n_months": len(backtest),
        "ann_return_top": annual_return(top),
        "ann_vol_top": top_volatility,
        "sharpe_top": (
            top.mean() * 12 / top_volatility
            if top_volatility and top_volatility > 0
            else np.nan
        ),
        "ann_return_ls": annual_return(long_short),
        "ann_vol_ls": long_short_volatility,
        "sharpe_ls": (
            long_short.mean() * 12 / long_short_volatility
            if (long_short_volatility and long_short_volatility > 0)
            else np.nan
        ),
        "win_rate_ls": ((long_short > 0).mean() if len(long_short) else np.nan),
        "max_drawdown_ls": max_drawdown(long_short),
    }


def _plot_heatmap(matrix, sectors, columns, title, figsize):
    """Dessine la même échelle et les annotations pour piliers et variables."""
    fig, ax = plt.subplots(figsize=figsize)
    image = ax.imshow(matrix, aspect="auto", vmin=0, vmax=max(10, np.nanmax(matrix)))
    ax.set_xticks(range(len(columns)))
    ax.set_xticklabels(columns)
    ax.set_yticks(range(len(sectors)))
    ax.set_yticklabels(sectors)
    ax.set_title(title)

    for i in range(len(sectors)):
        for j in range(len(columns)):
            value = matrix[i, j]
            label = "-" if np.isnan(value) else f"{value:.1f}"
            ax.text(j, i, label, ha="center", va="center", fontsize=8)
    fig.colorbar(image, ax=ax, label="Score")
    return fig, ax


class SectorResearch:
    """Regroupe la configuration, les consultations et les backtests de recherche."""

    def __init__(self, model, variables=None):
        if not isinstance(model, SectorModel):
            raise TypeError("model must be a SectorModel.")
        if variables and model.results is not None:
            raise ValueError("Select research variables before running the model.")

        self.model = model
        self.universe = model.universe
        self.config = model.config
        self._base_config = deepcopy(model.config)
        self._available_variables = deepcopy(self.config["backtest_variables"])
        self._variable_selection = {
            pillar: values.copy() for pillar, values in self._available_variables.items()
        }
        if variables:
            self.config = deepcopy(model.config)
            self._apply_variable_selection(variables)
            model.config = self.config

    def load(self):
        """Exécute le modèle et conserve les résultats en mémoire."""
        self.model.run()
        return self

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
        """Décrit les sources et transformations, une ligne par variable disponible."""
        if pillar is not None:
            self._validate_pillar(pillar)
        pillars = self.model.pillars if pillar is None else [pillar]
        rows = []

        for pillar_name in pillars:
            active = set(self._variable_selection[pillar_name])
            historical = self._base_config["historical_variables"].get(pillar_name, {})
            for variable in self._available_variables[pillar_name]:
                row = {
                    "universe": self.universe,
                    "pillar": pillar_name,
                    "variable": variable,
                    "active": variable in active,
                }
                if variable in historical:
                    definition = historical[variable]
                    window = self._base_config["history_window"]
                    sector_windows = definition.get("sector_windows", {})
                    if sector_windows:
                        overrides = ", ".join(
                            f"{sector}={months}" for sector, months in sector_windows.items()
                        )
                        window = f"{window}; {overrides}"
                    row.update({
                        "source": "historical",
                        "sheet": definition["sheet"],
                        "column": definition["column"],
                        "direction": "higher" if definition["rank_order"] == 1 else "lower",
                        "window": window,
                        "cross_sectional_mix": definition["mix_cross_sectional_rank"],
                    })
                elif pillar_name == "Momentum":
                    cfg = self._base_config["momentum"]
                    if variable in ("momentum_6m_1m", "momentum_12m_1m"):
                        horizon = "short_horizon" if variable == "momentum_6m_1m" else "long_horizon"
                        column = cfg["price_column"]
                        window = f"{cfg[horizon]}M-1M"
                        source = "price"
                    else:
                        column = " / ".join([
                            cfg["revision_up_column"], cfg["revision_down_column"],
                            cfg["revision_unchanged_column"],
                        ])
                        window = None
                        source = "earnings revisions"
                    row.update({
                        "source": source,
                        "sheet": cfg["sheet"],
                        "column": column,
                        "direction": "higher",
                        "window": window,
                        "cross_sectional_mix": True,
                    })
                else:
                    cfg = self._base_config["volatility"]
                    offset = "volatility_offset" if variable == "volatility_6m" else "offset_downside"
                    row.update({
                        "source": "sector returns",
                        "sheet": cfg["sheet_returns"],
                        "column": "derived",
                        "direction": "lower",
                        "window": f"{cfg[offset] + 1} observations",
                        "cross_sectional_mix": False,
                    })
                rows.append(row)

        return pd.DataFrame(rows)

    def composition(self, pillar):
        """Présente les variables réellement utilisées par secteur."""
        self._validate_pillar(pillar)
        composition = self.config["pillar_composition"][pillar]
        by_sector = composition.get("by_sector", {})
        rows = []
        for sector in self.model.sectors:
            variables = by_sector.get(sector, composition["default"])
            rows.append({"sector": sector, "variables": " | ".join(variables)})
        return pd.DataFrame(rows)

    def summary(self):
        """Présente la composition active du modèle."""
        rows = []

        for pillar in self.model.pillars:
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
        results = self._results()

        if variable not in results["variable_scores"]:
            raise ValueError(f"Unknown variable: {variable}")

        score = results["variable_scores"][variable]
        selected_date = self._select_date(score.index, date)

        return score.loc[selected_date].rename(variable)

    def pillar_values(self, pillar, date=None):
        """Retourne les scores sectoriels d'un pilier pour une date."""
        self._validate_pillar(pillar)
        results = self._results()
        score = results["pillars"][pillar]
        selected_date = self._select_date(score.index, date)

        return score.loc[selected_date].rename(pillar)

    def variable_table(self, pillar, date=None):
        """Compare les variables actives et le score final d'un pilier."""
        self._validate_pillar(pillar)
        results = self._results()

        variables = self._variable_selection[pillar]
        pillar_score = results["pillars"][pillar]
        selected_date = self._select_date(pillar_score.index, date)

        table = pd.DataFrame(index=self.model.sectors)

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
        history = self._results()["history"]

        if history.empty:
            return history.copy()

        dates = pd.DatetimeIndex(history["date"].drop_duplicates()).sort_values()
        selected_date = self._select_date(dates, date)

        return (
            history[history["date"] == selected_date].sort_values("rang_global", ascending=False)
            .reset_index(drop=True)
        )

    def save_outputs(self, output_dir="output"):
        """Sauvegarde les sorties du modèle."""
        results = self._results()

        save_all_outputs(results, Path(output_dir), self.config)

    def show_latest(self):
        """Affiche la dernière recommandation disponible."""
        print_latest(self._results())

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

            self._update_pillar_composition(pillar, variables)
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
        base_by_sector = base_composition.get("by_sector", {})

        default = [variable for variable in base_default if variable in variables]

        if not default:
            raise ValueError(f"No selected variable is valid for the default " f"{pillar} composition.")

        by_sector = {}

        for sector, sector_variables in base_by_sector.items():
            selected = [
                variable for variable in sector_variables if variable in variables
            ]

            if not selected:
                raise ValueError(f"No selected variable is valid for {pillar} / {sector}.")

            by_sector[sector] = selected

        self.config["pillar_composition"][pillar] = {
            "default": default,
        }

        if by_sector:
            self.config["pillar_composition"][pillar]["by_sector"] = by_sector

    def _validate_pillar(self, pillar):
        if pillar not in self._available_variables:
            raise ValueError(f"Unknown pillar: {pillar}")

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


    def _results(self):
        if self.model.results is None:
            self.model.run()
        return self.model.results

    def _pillar_date(self, date=None):
        """Choisit une date commune disponible pour les piliers."""
        pillars = self._results()["pillars"]
        dates = None

        for score in pillars.values():
            dates = score.index if dates is None else dates.intersection(score.index)

        dates = dates.sort_values()

        if len(dates) == 0:
            raise ValueError("Aucune date commune entre les piliers.")

        if date is None:
            return dates[-1]

        target = pd.Timestamp(date).to_period("M").to_timestamp("M")
        available = dates[dates <= target]

        if len(available) == 0:
            raise ValueError(f"Aucune date disponible avant {target.date()}.")

        return available[-1]

    def plot_pillars(self, date=None, pillars=None, figsize=(10, 7)):
        """Affiche les scores des piliers par secteur."""
        results = self._results()
        date = self._pillar_date(date)
        if pillars is None:
            pillars = list(self.config["base_weights"])
        unknown = [pillar for pillar in pillars if pillar not in results["pillars"]]
        if unknown:
            raise ValueError(f"Piliers inconnus : {unknown}")

        sectors = self.config["sectors"]
        matrix = np.array([
            [results["pillars"][pillar].at[date, sector] for pillar in pillars]
            for sector in sectors
        ], dtype=float)
        fig, ax = _plot_heatmap(
            matrix, sectors, pillars, f"{self.universe} - Scores des piliers au {date.date()}", figsize
        )
        fig.tight_layout()
        return fig, ax

    def plot_variables(self, pillar, variables=None, date=None, figsize=(9, 7)):
        """Affiche les variables sélectionnées d'un pilier."""
        results = self._results()
        date = self._pillar_date(date)
        active = self.variables(pillar)
        if variables is None:
            variables = active
        unknown = [variable for variable in variables if variable not in active]
        if unknown:
            raise ValueError(f"Variables inconnues pour {pillar} : {unknown}")

        sectors = self.config["sectors"]
        rows = []
        for sector in sectors:
            values = []
            for variable in variables:
                score = results["variable_scores"][variable]
                values.append(score.at[date, sector] if date in score.index else np.nan)
            rows.append(values)
        matrix = np.array(rows, dtype=float)
        fig, ax = _plot_heatmap(
            matrix, sectors, variables, f"{self.universe} - {pillar} au {date.date()}", figsize
        )
        ax.set_xticklabels(variables, rotation=20, ha="right")
        fig.tight_layout()
        return fig, ax

    def plot_global_rank(
        self,
        date=None,
        figsize=(9, 6),
    ):
        """Affiche le rang global pour une date de recommandation."""
        history = self._results()["history"]

        if history.empty:
            raise ValueError("Aucune recommandation finale disponible.")

        dates = history["date"].drop_duplicates().sort_values()

        if date is None:
            date = dates.iloc[-1]
        else:
            target = pd.Timestamp(date).to_period("M").to_timestamp("M")
            available = dates[dates <= target]

            if len(available) == 0:
                raise ValueError(f"Aucune recommandation avant {target.date()}.")

            date = available.iloc[-1]

        row = history[history["date"] == date].sort_values("rang_global")

        fig, ax = plt.subplots(figsize=figsize)
        ax.barh(row["secteur"], row["rang_global"])
        ax.set_xlabel(f"Rang final - 1 = Worst, " f"{len(self.config['sectors'])} = Best")
        ax.set_title(f"{self.universe} - Rang final au {date.date()}")
        fig.tight_layout()

        return fig, ax

    def backtest(
        self,
        pillar,
        variables=None,
        include_pillar=True,
        top_n=3,
        start=None,
        end=None,
        with_plot=True,
        plot_series=None,
        figsize=(10, 6),
    ):
        """Backteste les variables sélectionnées et le pilier final."""
        results = self._results()
        active = self.variables(pillar)

        if variables is None:
            variables = active

        unknown = [variable for variable in variables if variable not in active]

        if unknown:
            raise ValueError(f"Variables inconnues pour {pillar} : {unknown}")

        forward_returns = prepare_forward_returns(results["returns"])
        benchmark_returns = prepare_forward_returns(results["benchmark_returns"])[
            "benchmark"
        ]

        scores = {
            variable: results["variable_scores"][variable] for variable in variables
        }

        if include_pillar:
            scores[f"{pillar}_total"] = results["pillars"][pillar]

        rows = []
        backtests = {}

        for name, score in scores.items():
            backtest = backtest_score(
                score,
                forward_returns,
                benchmark_returns=benchmark_returns,
                top_n=top_n,
                start=start,
                end=end,
            )

            backtests[name] = backtest

            row = {"variable": name}
            row.update(statistics(backtest))
            rows.append(row)

        summary = (
            pd.DataFrame(rows).sort_values("sharpe_ls", ascending=False,).reset_index(drop=True)
        )

        fig = None
        ax = None

        if with_plot:
            fig, ax = self._plot_backtests(
                backtests, benchmark_returns, pillar, plot_series, figsize
            )

        return (summary, backtests, fig, ax)

    def _plot_backtests(self, backtests, benchmark_returns, pillar, plot_series, figsize):
        """Trace les séries déjà calculées, sur leurs dates de signal."""
        if plot_series is None:
            plot_series = ["top", "benchmark"]

        plot_series = list(plot_series)
        valid_series = {
            "top",
            "worst",
            "long_short",
            "benchmark",
        }
        unknown_series = [name for name in plot_series if name not in valid_series]

        if unknown_series:
            raise ValueError(f"Séries de plot inconnues : {unknown_series}")

        fig, ax = plt.subplots(figsize=figsize)

        columns = {
            "top": ("return_top", "Top"),
            "worst": ("return_worst", "Worst"),
            "long_short": ("return_long_short", "Long-Short"),
        }

        for name, backtest in backtests.items():
            if backtest.empty:
                continue

            indexed = backtest.set_index("date_signal")

            for series_name in plot_series:
                if series_name == "benchmark":
                    continue

                column, label = columns[series_name]
                series = indexed[column].dropna()
                cumulative = (1 + series).cumprod()

                ax.plot(cumulative.index, cumulative.values, label=(f"{name} - {label}"))

        if "benchmark" in plot_series and backtests:
            date_series = [
                backtest["date_signal"]
                for backtest in backtests.values()
                if not backtest.empty
            ]

            if date_series:
                date_min = min(series.min() for series in date_series)
                date_max = max(series.max() for series in date_series)

                benchmark = benchmark_returns[
                    (benchmark_returns.index >= date_min)
                    & (benchmark_returns.index <= date_max)
                ].dropna()

                cumulative = (1 + benchmark).cumprod()

                ax.plot(
                    cumulative.index,
                    cumulative.values,
                    label=self.config["volatility"]["benchmark_name"],
                )

        ax.set_title(f"{self.universe} - {pillar} : backtest")
        ax.set_ylabel("Valeur cumulée - base 1")
        ax.legend()
        ax.grid(alpha=0.2)
        fig.tight_layout()
        return fig, ax


    def save_backtest(
        self,
        summary,
        backtests,
        pillar,
        output_dir="output/backtests",
    ):
        """Sauvegarde un backtest déjà calculé."""
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        prefix = self.config["output_prefix"]

        save_table(summary, output_dir / f"{prefix}_{pillar.lower()}_summary.xlsx")

        for name, backtest in backtests.items():
            save_table(
                backtest,
                output_dir / f"{prefix}_{pillar.lower()}_{name}_monthly.xlsx",
            )
