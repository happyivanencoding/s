# -*- coding: utf-8 -*-
"""Visualisations des modèles sectoriels EU et US."""

from pathlib import Path
import argparse

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from local_config import MODELE_EU, MODELE_US
import model_secto_eu
import model_secto_us


def trouver_date(piliers, date_demandee=None):
    """Choisit la date demandée ou la dernière date commune."""
    dates = None

    for df in piliers.values():
        dates = (
            df.index
            if dates is None
            else dates.intersection(df.index)
        )

    dates = dates.sort_values()

    if len(dates) == 0:
        raise ValueError(
            "Aucune date commune entre les piliers."
        )

    if date_demandee is None:
        return dates[-1]

    cible = pd.Timestamp(date_demandee)
    dates_avant = dates[
        dates <= cible
    ]

    if len(dates_avant) == 0:
        raise ValueError(
            f"Aucune date disponible avant {cible.date()}."
        )

    return dates_avant[-1]


def plot_heatmap_piliers(
    piliers,
    date,
    fichier_sortie,
    config,
):
    """Affiche les six piliers par secteur."""
    noms_piliers = list(
        config["poids_base"]
    )
    secteurs = config["secteurs"]

    matrice = np.array(
        [
            [
                piliers[p].at[date, s]
                for p in noms_piliers
            ]
            for s in secteurs
        ],
        dtype=float,
    )

    vmax = max(
        10,
        np.nanmax(matrice),
    )

    fig, ax = plt.subplots(
        figsize=(10, 7)
    )
    image = ax.imshow(
        matrice,
        aspect="auto",
        vmin=0,
        vmax=vmax,
    )

    ax.set_xticks(
        range(len(noms_piliers))
    )
    ax.set_xticklabels(
        noms_piliers
    )
    ax.set_yticks(
        range(len(secteurs))
    )
    ax.set_yticklabels(
        secteurs
    )

    ax.set_title(
        f"{config['nom']} - Scores des piliers au {date.date()}"
    )

    for i in range(len(secteurs)):
        for j in range(len(noms_piliers)):
            valeur = matrice[i, j]
            texte = (
                "-"
                if np.isnan(valeur)
                else f"{valeur:.1f}"
            )
            ax.text(
                j,
                i,
                texte,
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
    fig.savefig(
        fichier_sortie,
        dpi=160,
    )
    plt.close(fig)


def plot_rang_global(
    historique,
    date,
    fichier_sortie,
    config,
):
    """Affiche le rang final pour une date."""
    ligne = historique[
        historique["date"] == date
    ].copy()

    if ligne.empty:
        return False

    ligne = ligne.sort_values(
        "rang_global"
    )

    fig, ax = plt.subplots(
        figsize=(9, 6)
    )
    ax.barh(
        ligne["secteur"],
        ligne["rang_global"],
    )
    ax.set_xlabel(
        f"Rang final - 1 = Worst, {len(config['secteurs'])} = Best"
    )
    ax.set_title(
        f"{config['nom']} - Rang final au {date.date()}"
    )
    ax.set_xlim(
        0,
        len(config["secteurs"]) + 0.5,
    )

    fig.tight_layout()
    fig.savefig(
        fichier_sortie,
        dpi=160,
    )
    plt.close(fig)

    return True


def plot_sous_variables(
    resultats,
    pilier,
    date,
    fichier_sortie,
    config,
):
    """Compare les sous-variables d'un pilier."""
    groupes = config[
        "variables_backtest"
    ]

    if pilier not in groupes:
        raise ValueError(
            f"Pilier inconnu : {pilier}"
        )

    variables = groupes[pilier]
    secteurs = config["secteurs"]
    matrice = []

    for secteur in secteurs:
        ligne = []

        for variable in variables:
            df = resultats[
                "sous_scores"
            ][variable]

            ligne.append(
                df.at[date, secteur]
                if date in df.index
                else np.nan
            )

        matrice.append(ligne)

    matrice = np.array(
        matrice,
        dtype=float,
    )
    vmax = max(
        10,
        np.nanmax(matrice),
    )

    fig, ax = plt.subplots(
        figsize=(9, 7)
    )
    image = ax.imshow(
        matrice,
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
        range(len(secteurs))
    )
    ax.set_yticklabels(
        secteurs
    )
    ax.set_title(
        f"{config['nom']} - {pilier} : sous-variables au {date.date()}"
    )

    for i in range(len(secteurs)):
        for j in range(len(variables)):
            valeur = matrice[i, j]
            texte = (
                "-"
                if np.isnan(valeur)
                else f"{valeur:.1f}"
            )
            ax.text(
                j,
                i,
                texte,
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
    fig.savefig(
        fichier_sortie,
        dpi=160,
    )
    plt.close(fig)


def generer_plots(
    resultats,
    config,
    dossier_sortie,
    date_demandee=None,
    pilier=None,
):
    """Génère les graphiques de contrôle du marché."""
    dossier = Path(
        dossier_sortie
    )
    dossier.mkdir(
        parents=True,
        exist_ok=True,
    )

    prefixe = config[
        "prefixe_sortie"
    ]
    date = trouver_date(
        resultats["piliers"],
        date_demandee,
    )

    plot_heatmap_piliers(
        resultats["piliers"],
        date,
        dossier
        / f"{prefixe}_heatmap_piliers.png",
        config,
    )

    date_finale = None

    if not resultats[
        "historique"
    ].empty:
        dates_finales = (
            resultats["historique"][
                "date"
            ]
            .drop_duplicates()
            .sort_values()
        )

        possibles = dates_finales[
            dates_finales <= date
        ]

        if len(possibles):
            date_finale = possibles.iloc[-1]

    if date_finale is not None:
        plot_rang_global(
            resultats["historique"],
            date_finale,
            dossier
            / f"{prefixe}_rang_global.png",
            config,
        )

    if pilier:
        plot_sous_variables(
            resultats,
            pilier,
            date,
            dossier
            / (
                f"{prefixe}_"
                f"{pilier.lower()}_sous_variables.png"
            ),
            config,
        )

    return date, date_finale

def choisir_marche(nom_marche):
    """Retourne la configuration et le module du marché demandé."""
    if nom_marche == "EU":
        return MODELE_EU, model_secto_eu

    return MODELE_US, model_secto_us


def main():
    parser = argparse.ArgumentParser(
        description="Plots des modèles sectoriels EU et US"
    )
    parser.add_argument(
        "--market",
        required=True,
        choices=["EU", "US"],
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
        "--date",
        default=None,
    )
    parser.add_argument(
        "--pillar",
        default=None,
    )
    parser.add_argument(
        "--output",
        default=str(
            Path(__file__).resolve().parent / "output"
        ),
    )
    args = parser.parse_args()

    config, modele = choisir_marche(args.market)

    if (
        args.pillar
        and args.pillar not in config["variables_backtest"]
    ):
        choix = ", ".join(config["variables_backtest"])
        raise ValueError(
            f"Pilier inconnu : {args.pillar}. Choix : {choix}"
        )

    resultats = modele.calculer_modele(
        args.excel,
        fichier_macro=args.macro_excel,
    )

    date, date_finale = generer_plots(
        resultats,
        config,
        args.output,
        date_demandee=args.date,
        pilier=args.pillar,
    )

    print(f"Plots sauvegardés dans : {args.output}")
    print(f"Date des piliers : {date.date()}")

    if date_finale is not None:
        print(
            f"Date du rang final : {date_finale.date()}"
        )


if __name__ == "__main__":
    main()
