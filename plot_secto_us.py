# -*- coding: utf-8 -*-
"""Visualisations du modèle sectoriel US."""

from pathlib import Path
import argparse

from local_config import MODELE_US
from model_secto_us import (
    FICHIER_EXCEL_PAR_DEFAUT,
    FICHIER_MACRO_PAR_DEFAUT,
    calculer_modele,
)
from sector_plot import generer_plots


CONFIG = MODELE_US
DOSSIER_SORTIE = (
    Path(__file__).resolve().parent
    / "output"
)


def main():
    parser = argparse.ArgumentParser(
        description="Plots du modèle sectoriel US"
    )
    parser.add_argument(
        "--excel",
        default=FICHIER_EXCEL_PAR_DEFAUT,
    )
    parser.add_argument(
        "--macro-excel",
        default=FICHIER_MACRO_PAR_DEFAUT,
    )
    parser.add_argument(
        "--date",
        default=None,
    )
    parser.add_argument(
        "--pillar",
        default=None,
        choices=list(
            CONFIG["variables_backtest"]
        ),
    )
    parser.add_argument(
        "--output",
        default=str(DOSSIER_SORTIE),
    )
    args = parser.parse_args()

    resultats = calculer_modele(
        args.excel,
        fichier_macro=args.macro_excel,
    )

    date, date_finale = generer_plots(
        resultats,
        CONFIG,
        args.output,
        date_demandee=args.date,
        pilier=args.pillar,
    )

    print(
        f"Plots sauvegardés dans : {args.output}"
    )
    print(
        f"Date des piliers : {date.date()}"
    )

    if date_finale is not None:
        print(
            f"Date du rang final : {date_finale.date()}"
        )


if __name__ == "__main__":
    main()
