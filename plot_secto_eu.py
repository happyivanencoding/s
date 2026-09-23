# -*- coding: utf-8 -*-
"""
Visualisations simples du modèle sectoriel Europe.

Exemples :
    python plot_secto_eu.py
    python plot_secto_eu.py --date 2026-08-31
    python plot_secto_eu.py --pillar Value
"""

from pathlib import Path
import argparse

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from model_secto_eu import (
    FICHIER_EXCEL_PAR_DEFAUT,
    SECTEURS,
    calculer_modele,
)


DOSSIER_SORTIE = Path(__file__).resolve().parent / "output"


def trouver_date(piliers, date_demandee=None):
    """Choisit la date demandée, ou la date la plus récente commune aux piliers."""
    dates = None

    for df in piliers.values():
        dates = df.index if dates is None else dates.intersection(df.index)

    dates = dates.sort_values()

    if len(dates) == 0:
        raise ValueError("Aucune date commune entre les piliers.")

    if date_demandee is None:
        return dates[-1]

    cible = pd.Timestamp(date_demandee)
    dates_avant = dates[dates <= cible]

    if len(dates_avant) == 0:
        raise ValueError(f"Aucune date disponible avant {cible.date()}.")

    return dates_avant[-1]


def plot_heatmap_piliers(piliers, date, fichier_sortie):
    """Heatmap 0-10 des six piliers pour les 13 secteurs."""
    noms_piliers = ["Leverage", "Margin", "Value", "Momentum", "Growth", "Volatility"]

    matrice = np.array(
        [[piliers[p].at[date, s] for p in noms_piliers] for s in SECTEURS],
        dtype=float,
    )

    fig, ax = plt.subplots(figsize=(10, 7))
    image = ax.imshow(matrice, aspect="auto", vmin=0, vmax=10)

    ax.set_xticks(range(len(noms_piliers)))
    ax.set_xticklabels(noms_piliers)
    ax.set_yticks(range(len(SECTEURS)))
    ax.set_yticklabels(SECTEURS)

    ax.set_title(f"EU - Scores des piliers au {date.date()}")

    # Valeurs dans les cases pour une lecture rapide.
    for i in range(len(SECTEURS)):
        for j in range(len(noms_piliers)):
            valeur = matrice[i, j]
            texte = "-" if np.isnan(valeur) else f"{valeur:.1f}"
            ax.text(j, i, texte, ha="center", va="center", fontsize=8)

    fig.colorbar(image, ax=ax, label="Score 0-10")
    fig.tight_layout()
    fig.savefig(fichier_sortie, dpi=160)
    plt.close(fig)


def plot_rang_global(historique, date, fichier_sortie):
    """Bar chart du rang final si la couche macro est disponible pour la date."""
    ligne = historique[historique["date"] == date].copy()

    if ligne.empty:
        return False

    ligne = ligne.sort_values("rang_global")

    fig, ax = plt.subplots(figsize=(9, 6))
    ax.barh(ligne["secteur"], ligne["rang_global"])
    ax.set_xlabel("Rang final - 1 = Worst, 13 = Best")
    ax.set_title(f"EU - Rang final au {date.date()}")
    ax.set_xlim(0, 13.5)
    fig.tight_layout()
    fig.savefig(fichier_sortie, dpi=160)
    plt.close(fig)

    return True


def plot_sous_variables(resultats, pilier, date, fichier_sortie):
    """Compare les sous-variables d'un pilier pour la date choisie."""
    groupes = {
        "Leverage": ["net_debt_ebitda", "fcf_total_debt", "debt_equity"],
        "Margin": ["operating_margin", "net_margin", "ebitda_margin"],
        "Value": ["price_fcf", "ev_ebitda", "price_sales"],
        "Momentum": [
            "momentum_6m_1m",
            "momentum_12m_1m",
            "earnings_revision_ratio",
        ],
        "Growth": ["eps_growth", "sales_growth", "ebitda_growth"],
        "Volatility": ["volatility_6m", "downside_volatility_18m"],
    }

    if pilier not in groupes:
        raise ValueError(f"Pilier inconnu : {pilier}")

    variables = groupes[pilier]
    matrice = []

    for secteur in SECTEURS:
        ligne = []
        for variable in variables:
            df = resultats["sous_scores"][variable]
            ligne.append(df.at[date, secteur] if date in df.index else np.nan)
        matrice.append(ligne)

    matrice = np.array(matrice, dtype=float)

    fig, ax = plt.subplots(figsize=(9, 7))
    image = ax.imshow(matrice, aspect="auto", vmin=0, vmax=10)

    ax.set_xticks(range(len(variables)))
    ax.set_xticklabels(variables, rotation=20, ha="right")
    ax.set_yticks(range(len(SECTEURS)))
    ax.set_yticklabels(SECTEURS)
    ax.set_title(f"EU - {pilier} : sous-variables au {date.date()}")

    for i in range(len(SECTEURS)):
        for j in range(len(variables)):
            valeur = matrice[i, j]
            texte = "-" if np.isnan(valeur) else f"{valeur:.1f}"
            ax.text(j, i, texte, ha="center", va="center", fontsize=8)

    fig.colorbar(image, ax=ax, label="Score 0-10")
    fig.tight_layout()
    fig.savefig(fichier_sortie, dpi=160)
    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(description="Plots du modèle sectoriel EU")
    parser.add_argument("--excel", default=FICHIER_EXCEL_PAR_DEFAUT)
    parser.add_argument("--date", default=None, help="Exemple : 2026-08-31")
    parser.add_argument(
        "--pillar",
        default=None,
        choices=["Leverage", "Margin", "Value", "Momentum", "Growth", "Volatility"],
    )
    parser.add_argument("--output", default=str(DOSSIER_SORTIE))
    args = parser.parse_args()

    dossier = Path(args.output)
    dossier.mkdir(parents=True, exist_ok=True)

    resultats = calculer_modele(args.excel)
    date = trouver_date(resultats["piliers"], args.date)

    plot_heatmap_piliers(
        resultats["piliers"],
        date,
        dossier / "eu_heatmap_piliers.png",
    )

    # Le rang global peut être plus ancien si le signal macro n'est pas encore fourni.
    date_finale = None
    if not resultats["historique"].empty:
        dates_finales = resultats["historique"]["date"].drop_duplicates().sort_values()
        possibles = dates_finales[dates_finales <= date]
        if len(possibles):
            date_finale = possibles.iloc[-1]

    if date_finale is not None:
        plot_rang_global(
            resultats["historique"],
            date_finale,
            dossier / "eu_rang_global.png",
        )

    if args.pillar:
        plot_sous_variables(
            resultats,
            args.pillar,
            date,
            dossier / f"eu_{args.pillar.lower()}_sous_variables.png",
        )

    print(f"Plots sauvegardés dans : {dossier}")
    print(f"Date des piliers : {date.date()}")
    if date_finale is not None:
        print(f"Date du rang final : {date_finale.date()}")


if __name__ == "__main__":
    main()
