# -*- coding: utf-8 -*-
"""Backtest des sous-variables du modèle sectoriel US."""

from pathlib import Path
import argparse

from local_config import MODELE_US
from model_secto_us import (
    FICHIER_EXCEL_PAR_DEFAUT,
    FICHIER_MACRO_PAR_DEFAUT,
    calculer_modele,
)
from sector_backtest import executer_backtests


CONFIG = MODELE_US
DOSSIER_SORTIE = (
    Path(__file__).resolve().parent
    / "output"
    / "backtests"
)


def main():
    parser = argparse.ArgumentParser(
        description="Backtest des sous-variables du modèle sectoriel US"
    )
    parser.add_argument(
        "--pillar",
        required=True,
        choices=list(
            CONFIG["variables_backtest"]
        ),
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
        default=str(DOSSIER_SORTIE),
    )
    args = parser.parse_args()

    resultats = calculer_modele(
        args.excel,
        fichier_macro=args.macro_excel,
    )

    resume, _ = executer_backtests(
        resultats,
        CONFIG,
        args.pillar,
        args.output,
        top_n=args.top,
        start=args.start,
        end=args.end,
        avec_plot=args.plot,
    )

    print("\nRésumé :")
    colonnes = [
        "variable",
        "n_months",
        "ann_return_top",
        "ann_return_ls",
        "sharpe_ls",
        "win_rate_ls",
        "max_drawdown_ls",
    ]
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
