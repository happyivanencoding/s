# -*- coding: utf-8 -*-
"""Backtests des modèles sectoriels EU et US."""

from pathlib import Path
import argparse
import math

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from local_config import MODELE_EU, MODELE_US
import model_secto_eu
import model_secto_us


def preparer_retour_futur(retours):
    """Aligne le rendement du mois suivant avec le signal courant."""
    retours = retours.sort_index()
    return retours.shift(-1)


def backtester_score(
    score,
    retours_futurs,
    top_n=3,
    start=None,
    end=None,
):
    """Backtest égal-pondéré d'un score sectoriel."""
    score = score.sort_index()
    dates = score.index.intersection(
        retours_futurs.index
    ).sort_values()

    if start:
        dates = dates[
            dates >= pd.Timestamp(start)
        ]
    if end:
        dates = dates[
            dates <= pd.Timestamp(end)
        ]

    lignes = []

    for date in dates:
        s = score.loc[date].dropna()
        r = retours_futurs.loc[date].dropna()

        communs = s.index.intersection(
            r.index
        )
        s = s.loc[communs]
        r = r.loc[communs]

        if len(s) < 2 * top_n:
            continue

        top = list(
            s.nlargest(top_n).index
        )
        worst = list(
            s.nsmallest(top_n).index
        )

        ret_top = r.loc[top].mean()
        ret_worst = r.loc[worst].mean()

        lignes.append(
            {
                "date_signal": date,
                "return_top": ret_top,
                "return_worst": ret_worst,
                "return_long_short": (
                    ret_top - ret_worst
                ),
                "top_sectors": " | ".join(top),
                "worst_sectors": " | ".join(worst),
            }
        )

    return pd.DataFrame(lignes)


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


def statistiques(backtest):
    """Calcule les statistiques de comparaison."""
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
    ls = backtest[
        "return_long_short"
    ].dropna()

    def ann_return(x):
        if len(x) == 0:
            return np.nan
        return (
            (1 + x).prod()
            ** (12 / len(x))
            - 1
        )

    def ann_vol(x):
        if len(x) < 2:
            return np.nan
        return (
            x.std(ddof=1)
            * math.sqrt(12)
        )

    vol_top = ann_vol(top)
    vol_ls = ann_vol(ls)

    return {
        "n_months": len(backtest),
        "ann_return_top": ann_return(top),
        "ann_vol_top": vol_top,
        "sharpe_top": (
            top.mean() * 12 / vol_top
            if vol_top and vol_top > 0
            else np.nan
        ),
        "ann_return_ls": ann_return(ls),
        "ann_vol_ls": vol_ls,
        "sharpe_ls": (
            ls.mean() * 12 / vol_ls
            if vol_ls and vol_ls > 0
            else np.nan
        ),
        "win_rate_ls": (
            (ls > 0).mean()
            if len(ls)
            else np.nan
        ),
        "max_drawdown_ls": max_drawdown(ls),
    }


def plot_cumul(
    backtests,
    fichier_sortie,
    config,
):
    """Compare les performances Long-Short cumulées."""
    fig, ax = plt.subplots(
        figsize=(10, 6)
    )

    for nom, bt in backtests.items():
        if bt.empty:
            continue

        serie = bt.set_index(
            "date_signal"
        )["return_long_short"]

        cumul = (
            1 + serie
        ).cumprod()

        ax.plot(
            cumul.index,
            cumul.values,
            label=nom,
        )

    ax.set_title(
        f"{config['nom']} - Backtest Long-Short des sous-variables"
    )
    ax.set_ylabel(
        "Valeur cumulée - base 1"
    )
    ax.legend()
    ax.grid(alpha=0.2)

    fig.tight_layout()
    fig.savefig(
        fichier_sortie,
        dpi=160,
    )
    plt.close(fig)


def executer_backtests(
    resultats,
    config,
    pilier,
    dossier_sortie,
    top_n=3,
    start=None,
    end=None,
    avec_plot=False,
):
    """Backteste les sous-variables et le pilier agrégé."""
    dossier = Path(
        dossier_sortie
    )
    dossier.mkdir(
        parents=True,
        exist_ok=True,
    )

    retours_futurs = (
        preparer_retour_futur(
            resultats["retours"]
        )
    )

    variables = config[
        "variables_backtest"
    ][pilier].copy()

    scores_a_tester = {
        variable: resultats[
            "sous_scores"
        ][variable]
        for variable in variables
    }

    scores_a_tester[
        f"{pilier}_total"
    ] = resultats[
        "piliers"
    ][pilier]

    stats = []
    backtests = {}

    for nom, score in scores_a_tester.items():
        bt = backtester_score(
            score=score,
            retours_futurs=retours_futurs,
            top_n=top_n,
            start=start,
            end=end,
        )

        backtests[nom] = bt

        bt.to_csv(
            dossier
            / (
                f"{config['prefixe_sortie']}_"
                f"{pilier.lower()}_{nom}_monthly.csv"
            ),
            index=False,
        )

        ligne_stats = {
            "variable": nom
        }
        ligne_stats.update(
            statistiques(bt)
        )
        stats.append(
            ligne_stats
        )

    resume = pd.DataFrame(
        stats
    ).sort_values(
        "sharpe_ls",
        ascending=False,
    )

    resume.to_csv(
        dossier
        / (
            f"{config['prefixe_sortie']}_"
            f"{pilier.lower()}_summary.csv"
        ),
        index=False,
    )

    if avec_plot:
        plot_cumul(
            backtests,
            dossier
            / (
                f"{config['prefixe_sortie']}_"
                f"{pilier.lower()}_long_short.png"
            ),
            config,
        )

    return resume, backtests

def choisir_marche(nom_marche):
    """Retourne la configuration et le module du marché demandé."""
    if nom_marche == "EU":
        return MODELE_EU, model_secto_eu

    return MODELE_US, model_secto_us


def main():
    parser = argparse.ArgumentParser(
        description="Backtest des sous-variables sectorielles EU et US"
    )
    parser.add_argument(
        "--market",
        required=True,
        choices=["EU", "US"],
    )
    parser.add_argument(
        "--pillar",
        required=True,
    )
    parser.add_argument(
        "--excel",
        default=None,
    )
    parser.add_argument(
        "--macro-excel",
        default=None,
    )
    parser.add_argument(
        "--top",
        type=int,
        default=3,
    )
    parser.add_argument(
        "--start",
        default=None,
    )
    parser.add_argument(
        "--end",
        default=None,
    )
    parser.add_argument(
        "--plot",
        action="store_true",
    )
    parser.add_argument(
        "--output",
        default=str(
            Path(__file__).resolve().parent
            / "output"
            / "backtests"
        ),
    )
    args = parser.parse_args()

    config, modele = choisir_marche(args.market)

    if args.pillar not in config["variables_backtest"]:
        choix = ", ".join(config["variables_backtest"])
        raise ValueError(
            f"Pilier inconnu : {args.pillar}. Choix : {choix}"
        )

    resultats = modele.calculer_modele(
        args.excel,
        fichier_macro=args.macro_excel,
    )

    resume, _ = executer_backtests(
        resultats,
        config,
        args.pillar,
        args.output,
        top_n=args.top,
        start=args.start,
        end=args.end,
        avec_plot=args.plot,
    )

    colonnes = [
        "variable",
        "n_months",
        "ann_return_top",
        "ann_return_ls",
        "sharpe_ls",
        "win_rate_ls",
        "max_drawdown_ls",
    ]

    print("\nRésumé :")
    print(
        resume[colonnes].to_string(
            index=False
        )
    )
    print(
        f"\nRésultats sauvegardés dans : {args.output}"
    )


if __name__ == "__main__":
    main()
