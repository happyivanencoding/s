# -*- coding: utf-8 -*-
"""Calcul du modèle sectoriel US."""

from pathlib import Path

from data_io import (
    FICHIER_EXCEL_US_PAR_DEFAUT,
    FICHIER_MACRO_PAR_DEFAUT,
    ouvrir_workbooks,
)
from local_config import MODELE_US
import sector_core as core


CONFIG = MODELE_US
FICHIER_EXCEL_PAR_DEFAUT = FICHIER_EXCEL_US_PAR_DEFAUT
DOSSIER_SORTIE_PAR_DEFAUT = (
    Path(__file__).resolve().parent / "output"
)


def calculer_diff_vs_moyenne(raw, moyenne_sans_finance=False):
    """Calcule l'écart du secteur à la moyenne du même mois."""
    return core.calculer_diff_vs_moyenne(
        raw,
        moyenne_sans_finance,
    )


def calculer_score_historique(
    diff,
    ordre_rank,
    fenetre_par_secteur=None,
):
    """Classe l'écart courant dans son historique mensuel."""
    return core.calculer_score_historique(
        diff,
        ordre_rank,
        CONFIG["fenetre_historique"],
        fenetre_par_secteur,
    )


def calculer_piliers_historiques(wb):
    """Calcule Leverage, Margin, Value et Growth."""
    return core.calculer_piliers_historiques(
        wb,
        CONFIG,
    )


def calculer_momentum(wb):
    """Calcule les trois composantes Momentum."""
    return core.calculer_momentum(
        wb,
        CONFIG,
    )


def calculer_volatilite(wb):
    """Calcule total volatility et downside volatility."""
    return core.calculer_volatilite(
        wb,
        CONFIG,
    )


def construire_contexte_macro(wb_secteur, wb_macro):
    """Assemble macro score, régime et rate overlay."""
    return core.construire_contexte_macro(
        wb_secteur,
        wb_macro,
        CONFIG,
    )


def aligner_piliers(piliers):
    """Aligne les piliers sur une clé mensuelle commune."""
    return core.aligner_piliers(piliers)


def calculer_rangs_piliers(piliers):
    """Classe les secteurs dans chaque pilier."""
    return core.calculer_rangs_piliers(piliers)


def calculer_score_pondere(rangs, poids, date):
    """Agrège les rangs avec les poids fournis."""
    return core.calculer_score_pondere(
        rangs,
        poids,
        date,
    )


def rang_secteurs(ligne):
    """Classe les secteurs de 1 = Worst à N = Best."""
    return core.rang_secteurs(ligne)


def choisir_top_worst(top_count, bottom_count, rang_global):
    """Sélectionne les secteurs Positive et Negative."""
    return core.choisir_top_worst(
        top_count,
        bottom_count,
        rang_global,
        CONFIG["n_top"],
        CONFIG["n_worst"],
    )


def calculer_resultats_finaux(piliers, contexte_macro):
    """Construit scores finaux, votes et recommandations."""
    return core.calculer_resultats_finaux(
        piliers,
        contexte_macro,
        CONFIG,
    )


def calculer_modele(
    fichier_excel=None,
    fichier_macro=None,
):
    """Exécute l'ensemble du modèle."""
    wb_secteur, wb_macro = ouvrir_workbooks(
        fichier_excel,
        fichier_macro,
        fichier_excel_defaut=FICHIER_EXCEL_PAR_DEFAUT,
    )

    try:
        return core.executer_modele(
            wb_secteur,
            wb_macro,
            CONFIG,
        )
    finally:
        wb_secteur.close()
        wb_macro.close()


def sauvegarder_sorties(resultats, dossier_sortie):
    """Sauvegarde les fichiers CSV du marché."""
    core.sauvegarder_sorties(
        resultats,
        dossier_sortie,
        CONFIG,
    )


def afficher_latest(resultats):
    """Affiche la dernière recommandation disponible."""
    core.afficher_latest(resultats)


def main():
    resultats = calculer_modele()
    sauvegarder_sorties(
        resultats,
        DOSSIER_SORTIE_PAR_DEFAUT,
    )
    afficher_latest(resultats)


if __name__ == "__main__":
    main()
