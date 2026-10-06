# -*- coding: utf-8 -*-
"""Outils de recherche pour les modèles sectoriels EU et US."""

from pathlib import Path
import math

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from model_secto import SectorModel


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
    dates = score.index.intersection(
        forward_returns.index
    ).sort_values()

    if start:
        dates = dates[dates >= pd.Timestamp(start)]

    if end:
        dates = dates[dates <= pd.Timestamp(end)]

    rows = []

    for date in dates:
        signal = score.loc[date].dropna()
        returns = forward_returns.loc[date].dropna()

        sectors = signal.index.intersection(
            returns.index
        )
        signal = signal.loc[sectors]
        returns = returns.loc[sectors]

        if len(signal) < 2 * top_n:
            continue

        top = list(
            signal.nlargest(top_n).index
        )
        worst = list(
            signal.nsmallest(top_n).index
        )

        return_top = returns.loc[top].mean()
        return_worst = returns.loc[worst].mean()
        return_benchmark = np.nan

        if (
            benchmark_returns is not None
            and date in benchmark_returns.index
        ):
            return_benchmark = benchmark_returns.loc[date]

        rows.append(
            {
                "date_signal": date,
                "return_top": return_top,
                "return_worst": return_worst,
                "return_long_short": (
                    return_top - return_worst
                ),
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

    wealth = (
        1 + returns.fillna(0)
    ).cumprod()
    drawdown = (
        wealth
        / wealth.cummax()
        - 1
    )

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

    top = backtest[
        "return_top"
    ].dropna()
    long_short = backtest[
        "return_long_short"
    ].dropna()

    def annual_return(series):
        if len(series) == 0:
            return np.nan

        return (
            (1 + series).prod()
            ** (12 / len(series))
            - 1
        )

    def annual_volatility(series):
        if len(series) < 2:
            return np.nan

        return (
            series.std(ddof=1)
            * math.sqrt(12)
        )

    top_volatility = annual_volatility(top)
    long_short_volatility = annual_volatility(
        long_short
    )

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
            long_short.mean()
            * 12
            / long_short_volatility
            if (
                long_short_volatility
                and long_short_volatility > 0
            )
            else np.nan
        ),
        "win_rate_ls": (
            (long_short > 0).mean()
            if len(long_short)
            else np.nan
        ),
        "max_drawdown_ls": max_drawdown(
            long_short
        ),
    }


class SectorResearch:
    """Fournit les visualisations et backtests d'un SectorModel."""

    def __init__(self, model):
        if not isinstance(model, SectorModel):
            raise TypeError(
                "model must be a SectorModel."
            )

        self.model = model
        self.universe = model.universe
        self.config = model.config

    def load(self, force=False):
        """Exécute le modèle et conserve les résultats en mémoire."""
        self.model.run(force=force)
        return self

    def variables(self, pillar=None):
        """Retourne les variables actives du modèle."""
        return self.model.variables(pillar)

    def _results(self):
        return self.model.run()

    def _pillar_date(self, date=None):
        """Choisit une date commune disponible pour les piliers."""
        pillars = self._results()["pillars"]
        dates = None

        for score in pillars.values():
            dates = (
                score.index
                if dates is None
                else dates.intersection(
                    score.index
                )
            )

        dates = dates.sort_values()

        if len(dates) == 0:
            raise ValueError(
                "Aucune date commune entre les piliers."
            )

        if date is None:
            return dates[-1]

        target = (
            pd.Timestamp(date)
            .to_period("M")
            .to_timestamp("M")
        )
        available = dates[dates <= target]

        if len(available) == 0:
            raise ValueError(
                f"Aucune date disponible avant "
                f"{target.date()}."
            )

        return available[-1]

    def plot_pillars(
        self,
        date=None,
        pillars=None,
        figsize=(10, 7),
    ):
        """Affiche les scores des piliers par secteur."""
        results = self._results()
        date = self._pillar_date(date)

        if pillars is None:
            pillars = list(
                self.config["poids_base"]
            )

        unknown = [
            pillar
            for pillar in pillars
            if pillar not in results["pillars"]
        ]

        if unknown:
            raise ValueError(
                f"Piliers inconnus : {unknown}"
            )

        sectors = self.config["secteurs"]
        matrix = np.array(
            [
                [
                    results["pillars"][pillar].at[
                        date,
                        sector,
                    ]
                    for pillar in pillars
                ]
                for sector in sectors
            ],
            dtype=float,
        )

        vmax = max(
            10,
            np.nanmax(matrix),
        )

        fig, ax = plt.subplots(
            figsize=figsize
        )
        image = ax.imshow(
            matrix,
            aspect="auto",
            vmin=0,
            vmax=vmax,
        )

        ax.set_xticks(
            range(len(pillars))
        )
        ax.set_xticklabels(
            pillars
        )
        ax.set_yticks(
            range(len(sectors))
        )
        ax.set_yticklabels(
            sectors
        )
        ax.set_title(
            f"{self.universe} - Scores des piliers "
            f"au {date.date()}"
        )

        for i in range(len(sectors)):
            for j in range(len(pillars)):
                value = matrix[i, j]
                label = (
                    "-"
                    if np.isnan(value)
                    else f"{value:.1f}"
                )
                ax.text(
                    j,
                    i,
                    label,
                    ha="center",
                    va="center",
                    fontsize=8,
                )

        fig.colorbar(
            image,
            ax=ax,
            label="Score",
        )
        fig.tight_layout()

        return fig, ax

    def plot_variables(
        self,
        pillar,
        variables=None,
        date=None,
        figsize=(9, 7),
    ):
        """Affiche les variables sélectionnées d'un pilier."""
        results = self._results()
        date = self._pillar_date(date)

        active = self.variables(pillar)

        if variables is None:
            variables = active

        unknown = [
            variable
            for variable in variables
            if variable not in active
        ]

        if unknown:
            raise ValueError(
                f"Variables inconnues pour "
                f"{pillar} : {unknown}"
            )

        sectors = self.config["secteurs"]
        matrix = []

        for sector in sectors:
            row = []

            for variable in variables:
                score = results[
                    "variable_scores"
                ][variable]

                row.append(
                    score.at[date, sector]
                    if date in score.index
                    else np.nan
                )

            matrix.append(row)

        matrix = np.array(
            matrix,
            dtype=float,
        )
        vmax = max(
            10,
            np.nanmax(matrix),
        )

        fig, ax = plt.subplots(
            figsize=figsize
        )
        image = ax.imshow(
            matrix,
            aspect="auto",
            vmin=0,
            vmax=vmax,
        )

        ax.set_xticks(
            range(len(variables))
        )
        ax.set_xticklabels(
            variables,
            rotation=20,
            ha="right",
        )
        ax.set_yticks(
            range(len(sectors))
        )
        ax.set_yticklabels(
            sectors
        )
        ax.set_title(
            f"{self.universe} - {pillar} "
            f"au {date.date()}"
        )

        for i in range(len(sectors)):
            for j in range(len(variables)):
                value = matrix[i, j]
                label = (
                    "-"
                    if np.isnan(value)
                    else f"{value:.1f}"
                )
                ax.text(
                    j,
                    i,
                    label,
                    ha="center",
                    va="center",
                    fontsize=8,
                )

        fig.colorbar(
            image,
            ax=ax,
            label="Score",
        )
        fig.tight_layout()

        return fig, ax

    def plot_global_rank(
        self,
        date=None,
        figsize=(9, 6),
    ):
        """Affiche le rang global pour une date de recommandation."""
        history = self._results()[
            "history"
        ]

        if history.empty:
            raise ValueError(
                "Aucune recommandation finale disponible."
            )

        dates = (
            history["date"]
            .drop_duplicates()
            .sort_values()
        )

        if date is None:
            date = dates.iloc[-1]
        else:
            target = (
                pd.Timestamp(date)
                .to_period("M")
                .to_timestamp("M")
            )
            available = dates[
                dates <= target
            ]

            if len(available) == 0:
                raise ValueError(
                    f"Aucune recommandation avant "
                    f"{target.date()}."
                )

            date = available.iloc[-1]

        row = history[
            history["date"] == date
        ].sort_values(
            "rang_global"
        )

        fig, ax = plt.subplots(
            figsize=figsize
        )
        ax.barh(
            row["secteur"],
            row["rang_global"],
        )
        ax.set_xlabel(
            f"Rang final - 1 = Worst, "
            f"{len(self.config['secteurs'])} = Best"
        )
        ax.set_title(
            f"{self.universe} - Rang final "
            f"au {date.date()}"
        )
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

        unknown = [
            variable
            for variable in variables
            if variable not in active
        ]

        if unknown:
            raise ValueError(
                f"Variables inconnues pour "
                f"{pillar} : {unknown}"
            )

        forward_returns = prepare_forward_returns(
            results["returns"]
        )
        benchmark_returns = (
            prepare_forward_returns(
                results["benchmark_returns"]
            )["benchmark"]
        )

        scores = {
            variable: results[
                "variable_scores"
            ][variable]
            for variable in variables
        }

        if include_pillar:
            scores[f"{pillar}_total"] = (
                results["pillars"][pillar]
            )

        rows = []
        backtests = {}

        for name, score in scores.items():
            backtest = backtest_score(
                score,
                forward_returns,
                benchmark_returns=(
                    benchmark_returns
                ),
                top_n=top_n,
                start=start,
                end=end,
            )

            backtests[name] = backtest

            row = {
                "variable": name
            }
            row.update(
                statistics(backtest)
            )
            rows.append(row)

        summary = (
            pd.DataFrame(rows)
            .sort_values(
                "sharpe_ls",
                ascending=False,
            )
            .reset_index(drop=True)
        )

        fig = None
        ax = None

        if with_plot:
            if plot_series is None:
                plot_series = [
                    "top",
                    "benchmark",
                ]

            plot_series = list(
                plot_series
            )
            valid_series = {
                "top",
                "worst",
                "long_short",
                "benchmark",
            }
            unknown_series = [
                name
                for name in plot_series
                if name not in valid_series
            ]

            if unknown_series:
                raise ValueError(
                    f"Séries de plot inconnues : "
                    f"{unknown_series}"
                )

            fig, ax = plt.subplots(
                figsize=figsize
            )

            columns = {
                "top": (
                    "return_top",
                    "Top",
                ),
                "worst": (
                    "return_worst",
                    "Worst",
                ),
                "long_short": (
                    "return_long_short",
                    "Long-Short",
                ),
            }

            for name, backtest in backtests.items():
                if backtest.empty:
                    continue

                indexed = backtest.set_index(
                    "date_signal"
                )

                for series_name in plot_series:
                    if series_name == "benchmark":
                        continue

                    column, label = columns[
                        series_name
                    ]
                    series = indexed[
                        column
                    ].dropna()
                    cumulative = (
                        1 + series
                    ).cumprod()

                    ax.plot(
                        cumulative.index,
                        cumulative.values,
                        label=(
                            f"{name} - {label}"
                        ),
                    )

            if (
                "benchmark" in plot_series
                and backtests
            ):
                date_series = [
                    backtest["date_signal"]
                    for backtest in backtests.values()
                    if not backtest.empty
                ]

                if date_series:
                    date_min = min(
                        series.min()
                        for series in date_series
                    )
                    date_max = max(
                        series.max()
                        for series in date_series
                    )

                    benchmark = benchmark_returns[
                        (
                            benchmark_returns.index
                            >= date_min
                        )
                        & (
                            benchmark_returns.index
                            <= date_max
                        )
                    ].dropna()

                    cumulative = (
                        1 + benchmark
                    ).cumprod()

                    ax.plot(
                        cumulative.index,
                        cumulative.values,
                        label=self.config[
                            "volatilite"
                        ]["benchmark_nom"],
                    )

            ax.set_title(
                f"{self.universe} - "
                f"{pillar} : backtest"
            )
            ax.set_ylabel(
                "Valeur cumulée - base 1"
            )
            ax.legend()
            ax.grid(alpha=0.2)
            fig.tight_layout()

        return (
            summary,
            backtests,
            fig,
            ax,
        )

    def save_backtest(
        self,
        summary,
        backtests,
        pillar,
        output_dir="output/backtests",
    ):
        """Sauvegarde un backtest déjà calculé."""
        output_dir = Path(output_dir)
        output_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        prefix = self.config[
            "prefixe_sortie"
        ]

        summary.to_csv(
            output_dir
            / (
                f"{prefix}_"
                f"{pillar.lower()}_summary.csv"
            ),
            index=False,
        )

        for name, backtest in backtests.items():
            backtest.to_csv(
                output_dir
                / (
                    f"{prefix}_"
                    f"{pillar.lower()}_"
                    f"{name}_monthly.csv"
                ),
                index=False,
            )
