# -*- coding: utf-8 -*-
"""
Backtest rapide des sous-variables du modèle sectoriel Europe.

Le principe est volontairement simple :
- signal calculé à la fin du mois t ;
- Top N secteurs selon le score 0-10 ;
- performance mesurée sur le mois t+1 ;
- comparaison Long Top / Short Worst / Long-Short.

Exemples :
    python backtest_piliers_eu.py --pillar Value
    python backtest_piliers_eu.py --pillar Momentum --top 3
    python backtest_piliers_eu.py --pillar Growth --start 2015-01-01 --plot
"""

from pathlib import Path
import argparse
import math

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from model_secto_eu import (
    FICHIER_EXCEL_PAR_DEFAUT,
    calculer_modele,
)


DOSSIER_SORTIE = Path(__file__).resolve().parent / "output" / "backtests"

VARIABLES_PAR_PILIER = {
    "Leverage": [
        "net_debt_ebitda",
        "fcf_total_debt",
        "debt_equity",
    ],
    "Margin": [
        "operating_margin",
        "net_margin",
        "ebitda_margin",
    ],
    "Value": [
        "price_fcf",
        "ev_ebitda",
        "price_sales",
    ],
    "Momentum": [
        "momentum_6m_1m",
        "momentum_12m_1m",
        "earnings_revision_ratio",
    ],
    "Growth": [
        "eps_growth",
        "sales_growth",
        "ebitda_growth",
    ],
    "Volatility": [
        "volatility_6m",
        "downside_volatility_18m",
    ],
}


def preparer_retour_futur(retours):
    """
    Le rendement daté t est le rendement du mois qui se termine à t.
    Pour tester un signal observé à t, on utilise donc le rendement
    de la date mensuelle suivante.
    """
    retours = retours.sort_index()
    return retours.shift(-1)


def backtester_score(score, retours_futurs, top_n=3, start=None, end=None):
    """Backtest égal-pondéré d'un score sectoriel déjà orienté 0-10."""
    score = score.sort_index()

    dates = score.index.intersection(retours_futurs.index).sort_values()

    if start:
        dates = dates[dates >= pd.Timestamp(start)]
    if end:
        dates = dates[dates <= pd.Timestamp(end)]

    lignes = []

    for date in dates:
        s = score.loc[date].dropna()
        r = retours_futurs.loc[date].dropna()

        communs = s.index.intersection(r.index)
        s = s.loc[communs]
        r = r.loc[communs]

        if len(s) < 2 * top_n:
            continue

        top = list(s.nlargest(top_n).index)
        worst = list(s.nsmallest(top_n).index)

        ret_top = r.loc[top].mean()
        ret_worst = r.loc[worst].mean()
        ret_ls = ret_top - ret_worst

        lignes.append(
            {
                "date_signal": date,
                "return_top": ret_top,
                "return_worst": ret_worst,
                "return_long_short": ret_ls,
                "top_sectors": " | ".join(top),
                "worst_sectors": " | ".join(worst),
            }
        )

    return pd.DataFrame(lignes)


def max_drawdown(returns):
    """Maximum drawdown d'une série de rendements mensuels."""
    if len(returns) == 0:
        return np.nan

    wealth = (1 + returns.fillna(0)).cumprod()
    drawdown = wealth / wealth.cummax() - 1
    return drawdown.min()


def statistiques(backtest):
    """Quelques statistiques simples pour comparer rapidement les variables."""
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
    ls = backtest["return_long_short"].dropna()

    def ann_return(x):
        if len(x) == 0:
            return np.nan
        return (1 + x).prod() ** (12 / len(x)) - 1

    def ann_vol(x):
        if len(x) < 2:
            return np.nan
        return x.std(ddof=1) * math.sqrt(12)

    vol_top = ann_vol(top)
    vol_ls = ann_vol(ls)

    return {
        "n_months": len(backtest),
        "ann_return_top": ann_return(top),
        "ann_vol_top": vol_top,
        "sharpe_top": (top.mean() * 12 / vol_top) if vol_top and vol_top > 0 else np.nan,
        "ann_return_ls": ann_return(ls),
        "ann_vol_ls": vol_ls,
        "sharpe_ls": (ls.mean() * 12 / vol_ls) if vol_ls and vol_ls > 0 else np.nan,
        "win_rate_ls": (ls > 0).mean() if len(ls) else np.nan,
        "max_drawdown_ls": max_drawdown(ls),
    }


def plot_cumul(backtests, fichier_sortie):
    """Plot comparatif des performances Long-Short cumulées."""
    fig, ax = plt.subplots(figsize=(10, 6))

    for nom, bt in backtests.items():
        if bt.empty:
            continue

        serie = bt.set_index("date_signal")["return_long_short"]
        cumul = (1 + serie).cumprod()
        ax.plot(cumul.index, cumul.values, label=nom)

    ax.set_title("EU - Backtest Long-Short des sous-variables")
    ax.set_ylabel("Valeur cumulée - base 1")
    ax.legend()
    ax.grid(alpha=0.2)
    fig.tight_layout()
    fig.savefig(fichier_sortie, dpi=160)
    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(
        description="Backtest rapide des sous-variables du modèle sectoriel EU"
    )
    parser.add_argument(
        "--pillar",
        required=True,
        choices=list(VARIABLES_PAR_PILIER),
    )
    parser.add_argument("--excel", default=FICHIER_EXCEL_PAR_DEFAUT)
    parser.add_argument("--top", type=int, default=3)
    parser.add_argument("--start", default=None)
    parser.add_argument("--end", default=None)
    parser.add_argument("--plot", action="store_true")
    parser.add_argument("--output", default=str(DOSSIER_SORTIE))
    args = parser.parse_args()

    dossier = Path(args.output)
    dossier.mkdir(parents=True, exist_ok=True)

    resultats = calculer_modele(args.excel)
    retours_futurs = preparer_retour_futur(resultats["retours"])

    variables = VARIABLES_PAR_PILIER[args.pillar].copy()

    # On ajoute aussi le pilier agrégé pour comparer chaque sous-variable
    # avec la version finale utilisée dans le modèle.
    scores_a_tester = {
        variable: resultats["sous_scores"][variable]
        for variable in variables
    }
    scores_a_tester[f"{args.pillar}_total"] = resultats["piliers"][args.pillar]

    stats = []
    backtests = {}

    for nom, score in scores_a_tester.items():
        bt = backtester_score(
            score=score,
            retours_futurs=retours_futurs,
            top_n=args.top,
            start=args.start,
            end=args.end,
        )

        backtests[nom] = bt

        fichier_bt = dossier / f"{args.pillar.lower()}_{nom}_monthly.csv"
        bt.to_csv(fichier_bt, index=False)

        ligne_stats = {"variable": nom}
        ligne_stats.update(statistiques(bt))
        stats.append(ligne_stats)

    resume = pd.DataFrame(stats)
    resume = resume.sort_values("sharpe_ls", ascending=False)
    resume.to_csv(
        dossier / f"{args.pillar.lower()}_summary.csv",
        index=False,
    )

    print("\nRésumé :")
    colonnes_console = [
        "variable",
        "n_months",
        "ann_return_top",
        "ann_return_ls",
        "sharpe_ls",
        "win_rate_ls",
        "max_drawdown_ls",
    ]
    print(resume[colonnes_console].to_string(index=False))

    if args.plot:
        plot_cumul(
            backtests,
            dossier / f"{args.pillar.lower()}_long_short.png",
        )

    print(f"\nRésultats sauvegardés dans : {dossier}")


if __name__ == "__main__":
    main()
