# -*- coding: utf-8 -*-
"""Calcul du modèle sectoriel Europe."""

from pathlib import Path
import math

import numpy as np
import pandas as pd


# Configuration

from local_config import (
    ARRONDI_PILIER,
    CONFIG_MOMENTUM,
    CONFIG_SIGNAL_TAUX,
    CONFIG_VOLATILITE,
    EXCLUSIONS_PILIER,
    FENETRE_HISTORIQUE,
    N_TOP,
    N_WORST,
    PILIERS_RATE_OVERLAY,
    POIDS_BASE,
    POIDS_REGIME,
    POIDS_VOTE_MACRO,
    SECTEURS,
    SECTEURS_SANS_VOTE_TOP_WORST,
    VARIABLES_HISTORIQUES,
)

from data_io import (
    FICHIER_EXCEL_PAR_DEFAUT,
    FICHIER_MACRO_PAR_DEFAUT,
    lire_bloc_retours,
    lire_dates_et_bloc,
    lire_macro_externe,
    lire_taux_us10y,
    normaliser_index_mensuel,
    ouvrir_workbooks,
)


DOSSIER_SORTIE_PAR_DEFAUT = Path(__file__).resolve().parent / "output"

N_SECTEURS = len(SECTEURS)
N_OBSERVATIONS_RANK = FENETRE_HISTORIQUE + 1


# Classements

def rang_excel(valeur, valeurs, ordre):
    """Reproduit RANK Excel. ordre=1 croissant, ordre=0 décroissant."""
    if pd.isna(valeur):
        return np.nan

    serie = pd.Series(valeurs, dtype="float64").dropna()
    if serie.empty:
        return np.nan

    # Excel compare les doubles directement. Une différence même minuscule
    # peut donc produire deux rangs différents.
    if ordre == 1:
        return 1 + int((serie < valeur).sum())
    return 1 + int((serie > valeur).sum())


def rang_transversal_0_10(ligne, ordre):
    """Convertit le classement transversal des secteurs sur une échelle 0-10."""
    result = pd.Series(index=ligne.index, dtype=float)

    for secteur in ligne.index:
        valeur = ligne[secteur]
        if pd.isna(valeur):
            result[secteur] = np.nan
            continue
        rang = rang_excel(valeur, ligne.values, ordre)
        result[secteur] = 10 * rang / N_SECTEURS

    return result


def rang_transversal_momentum(ligne):
    """
    Rang 0-10 utilisé dans MOM_FMA.

    Excel écrit explicitement :
        RANK(...) / 13 * 10

    L'ordre des opérations est conservé car il peut laisser une différence
    de flottant de l'ordre de 1e-15 entre deux scores presque identiques.
    """
    result = pd.Series(index=ligne.index, dtype=float)

    for secteur in ligne.index:
        valeur = ligne[secteur]
        if pd.isna(valeur):
            result[secteur] = np.nan
            continue

        rang = rang_excel(valeur, ligne.values, ordre=1)
        result[secteur] = rang / N_SECTEURS * 10

    return result


def rang_secteurs(ligne):
    """Classe les secteurs de 1 = Worst à 13 = Best."""
    result = pd.Series(index=ligne.index, dtype=float)

    for secteur in ligne.index:
        valeur = ligne[secteur]
        if pd.isna(valeur):
            result[secteur] = np.nan
            continue
        result[secteur] = rang_excel(valeur, ligne.values, ordre=1)

    return result


def percentrank_inc(valeurs, x):
    """Reproduit PERCENTRANK.INC d'Excel pour le signal de taux."""
    valeurs = pd.Series(valeurs, dtype=float).dropna().sort_values().to_numpy()

    if len(valeurs) == 0 or pd.isna(x):
        return np.nan
    if len(valeurs) == 1:
        return 1.0
    if x <= valeurs[0]:
        return 0.0
    if x >= valeurs[-1]:
        return 1.0

    positions = np.where(np.isclose(valeurs, x, rtol=0, atol=1e-12))[0]
    if len(positions):
        return positions[0] / (len(valeurs) - 1)

    droite = np.searchsorted(valeurs, x, side="right")
    gauche = droite - 1
    x0, x1 = valeurs[gauche], valeurs[droite]
    fraction = (x - x0) / (x1 - x0)

    return (gauche + fraction) / (len(valeurs) - 1)


def calculer_signal_taux(taux_us10y):
    """
    Reproduit le rate overlay du modèle sectoriel.

    Le signal est On lorsque l'écart US10Y - EWMA se situe
    au-dessus du percentile configuré.
    """
    cfg = CONFIG_SIGNAL_TAUX
    alpha = cfg["alpha_ewma"]
    seuil = cfg["seuil_percentile"]

    df = taux_us10y.copy().sort_index()
    df["ewma_us10y"] = np.nan
    df["diff_ewma"] = np.nan
    df["percentile_taux"] = np.nan
    df["signal_taux"] = None

    for i in range(len(df)):
        taux = df.iloc[i]["us10y"]
        if pd.isna(taux):
            continue

        if i == 0 or pd.isna(df.iloc[i - 1]["ewma_us10y"]):
            ewma = taux
        else:
            ewma_prec = df.iloc[i - 1]["ewma_us10y"]
            ewma = alpha * ewma_prec + (1 - alpha) * taux

        df.iloc[i, df.columns.get_loc("ewma_us10y")] = ewma

        diff = taux - ewma
        df.iloc[i, df.columns.get_loc("diff_ewma")] = diff

        historique = df["diff_ewma"].iloc[: i + 1]
        percentile = percentrank_inc(historique, diff)
        df.iloc[i, df.columns.get_loc("percentile_taux")] = percentile
        df.iloc[i, df.columns.get_loc("signal_taux")] = (
            "On" if percentile > seuil else "Off"
        )

    return df


def construire_contexte_macro(wb_eu, wb_macro):
    """
    Assemble le macro externe et le rate overlay.

    Le macro_score et le cycle viennent du fichier macro externe.
    Le signal_taux est recalculé séparément à partir du US 10Y.
    """
    macro = lire_macro_externe(wb_macro)
    taux = calculer_signal_taux(lire_taux_us10y(wb_eu))

    contexte = macro.join(taux[["signal_taux"]], how="outer")
    return contexte.sort_index()


# Sous-variables historiques

def calculer_diff_vs_moyenne(raw, moyenne_sans_finance=False):
    """
    Reproduit "diff vs. moy" :
    valeur du secteur - moyenne des secteurs du même mois.

    Pour certaines métriques Excel exclut Financials de la moyenne.
    """
    diff = pd.DataFrame(index=raw.index, columns=raw.columns, dtype=float)

    for date, ligne in raw.iterrows():
        base = ligne.copy()
        if moyenne_sans_finance:
            base = base.drop("Fin")

        moyenne = base.mean(skipna=True)
        diff.loc[date] = ligne - moyenne

    return diff


def calculer_score_historique(diff, ordre_rank, fenetre_par_secteur=None):
    """
    Reproduit la formule Excel :
    10 * RANK(valeur_courante, fenêtre_historique, ordre) / nombre_observations

    Par défaut Excel utilise 60 mois + le mois courant = 61 observations.
    Quelques cellules ont une fenêtre spécifique, par exemple Technology
    dans certaines métriques de Value.
    """
    fenetre_par_secteur = fenetre_par_secteur or {}
    score = pd.DataFrame(index=diff.index, columns=diff.columns, dtype=float)

    for i in range(len(diff)):
        for secteur in SECTEURS:
            valeur = diff.iloc[i][secteur]
            if pd.isna(valeur):
                continue

            n_mois = fenetre_par_secteur.get(secteur, FENETRE_HISTORIQUE)
            n_observations = n_mois + 1
            fin = i + n_observations

            if fin > len(diff):
                continue

            valeurs = diff.iloc[i:fin][secteur]
            rang = rang_excel(valeur, valeurs.values, ordre_rank)
            score.iloc[i, score.columns.get_loc(secteur)] = (
                10 * rang / n_observations
            )

    return score


def calculer_variable_historique(
    raw,
    ordre_rank,
    moyenne_sans_finance,
    mix_rank_transversal,
    fenetre_par_secteur=None,
):
    """Calcule le score historique d'une sous-variable FMA."""
    diff = calculer_diff_vs_moyenne(raw, moyenne_sans_finance)
    score_hist = calculer_score_historique(
        diff,
        ordre_rank,
        fenetre_par_secteur=fenetre_par_secteur,
    )

    if not mix_rank_transversal:
        return score_hist

    # Margin_FMA fait la moyenne entre :
    # - le score historique 5 ans ;
    # - le rang transversal du ratio brut au même mois.
    score = score_hist.copy()

    for date in score.index:
        score_cross = rang_transversal_0_10(raw.loc[date], ordre_rank)
        score.loc[date] = (score_hist.loc[date] + score_cross) / 2

    return score


# Piliers Leverage / Margin / Value / Growth

def calculer_piliers_historiques(wb):
    """Calcule Leverage, Margin, Value et Growth."""
    sous_scores = {}
    piliers = {}

    for pilier, variables in VARIABLES_HISTORIQUES.items():
        scores_du_pilier = {}

        for nom_variable, cfg in variables.items():
            ws = wb[cfg["sheet"]]
            raw = lire_dates_et_bloc(ws, cfg["colonne"])

            score = calculer_variable_historique(
                raw=raw,
                ordre_rank=cfg["ordre_rank"],
                moyenne_sans_finance=cfg["moyenne_sans_finance"],
                mix_rank_transversal=cfg["mix_rank_transversal"],
                fenetre_par_secteur=cfg.get("fenetre_par_secteur"),
            )

            sous_scores[nom_variable] = score
            scores_du_pilier[nom_variable] = score

        # Même logique que AVERAGE dans les formules principales.
        # Pour Finance, Growth utilise seulement EPS + Sales car EBITDA est N/A.
        index_commun = next(iter(scores_du_pilier.values())).index
        pilier_df = pd.DataFrame(index=index_commun, columns=SECTEURS, dtype=float)

        for date in index_commun:
            for secteur in SECTEURS:
                valeurs = []

                exclusions = EXCLUSIONS_PILIER.get(pilier, {}).get(
                    secteur,
                    set(),
                )

                for nom_variable, score in scores_du_pilier.items():
                    if nom_variable in exclusions:
                        continue

                    v = score.at[date, secteur]
                    valeurs.append(v)

                # Excel propage une erreur si une composante explicitement utilisée est N/A.
                if any(pd.isna(v) for v in valeurs):
                    pilier_df.at[date, secteur] = np.nan
                else:
                    # AVERAGE Excel : addition puis division.
                    # On évite np.mean pour garder les mêmes arrondis binaires.
                    pilier_df.at[date, secteur] = sum(valeurs) / len(valeurs)

        # Certains piliers ont un arrondi explicite pour reproduire
        # les ex-aequo observés dans Excel.
        if pilier in ARRONDI_PILIER:
            pilier_df = pilier_df.round(ARRONDI_PILIER[pilier])

        piliers[pilier] = pilier_df

    return piliers, sous_scores


# Momentum

def calculer_momentum(wb):
    """Calcule les trois composantes Momentum puis leur score moyen."""
    cfg = CONFIG_MOMENTUM
    ws = wb[cfg["sheet"]]

    prix = lire_dates_et_bloc(ws, cfg["colonne_prix"])
    up = lire_dates_et_bloc(ws, cfg["colonne_revision_up"])
    down = lire_dates_et_bloc(ws, cfg["colonne_revision_down"])
    unchanged = lire_dates_et_bloc(ws, cfg["colonne_revision_unchanged"])

    horizon_court = cfg["horizon_court"]
    horizon_long = cfg["horizon_long"]
    secteur_mois_courant = cfg["secteur_mois_courant"]

    ret_6m_1m = pd.DataFrame(index=prix.index, columns=SECTEURS, dtype=float)
    ret_12m_1m = pd.DataFrame(index=prix.index, columns=SECTEURS, dtype=float)

    for i in range(len(prix)):
        # Formule standard : le dernier mois est exclu.
        if i + horizon_court < len(prix):
            ret_6m_1m.iloc[i] = (
                prix.iloc[i + 1] / prix.iloc[i + horizon_court] - 1
            )

        if i + horizon_long < len(prix):
            ret_12m_1m.iloc[i] = (
                prix.iloc[i + 1] / prix.iloc[i + horizon_long] - 1
            )

        # Exception explicitement configurée pour Technology.
        if i + horizon_court < len(prix):
            ret_6m_1m.at[prix.index[i], secteur_mois_courant] = (
                prix.iloc[i][secteur_mois_courant]
                / prix.iloc[i + horizon_court][secteur_mois_courant]
                - 1
            )

        if i + horizon_long < len(prix):
            ret_12m_1m.at[prix.index[i], secteur_mois_courant] = (
                prix.iloc[i][secteur_mois_courant]
                / prix.iloc[i + horizon_long][secteur_mois_courant]
                - 1
            )

    revision_ratio = (up - down) / (up + down + unchanged)

    score_6m = ret_6m_1m.apply(rang_transversal_momentum, axis=1)
    score_12m = ret_12m_1m.apply(rang_transversal_momentum, axis=1)
    score_revision = revision_ratio.apply(rang_transversal_momentum, axis=1)

    momentum = pd.DataFrame(
        index=score_6m.index,
        columns=SECTEURS,
        dtype=float,
    )

    for date in momentum.index:
        for secteur in SECTEURS:
            valeurs = [
                score_6m.at[date, secteur],
                score_12m.at[date, secteur],
                score_revision.at[date, secteur],
            ]

            if any(pd.isna(v) for v in valeurs):
                momentum.at[date, secteur] = np.nan
            else:
                momentum.at[date, secteur] = sum(valeurs) / 3

    sous_scores = {
        "momentum_6m_1m": score_6m,
        "momentum_12m_1m": score_12m,
        "earnings_revision_ratio": score_revision,
    }

    return momentum, sous_scores


# Low Volatility

def calculer_volatilite(wb):
    """Calcule la volatilité totale, la downside volatility et leur score moyen."""
    cfg = CONFIG_VOLATILITE
    retours = lire_bloc_retours(wb[cfg["sheet_retours"]])

    vol_6m = pd.DataFrame(index=retours.index, columns=SECTEURS, dtype=float)
    vol_down_18m = pd.DataFrame(index=retours.index, columns=SECTEURS, dtype=float)

    n_obs_vol = cfg["offset_volatilite"] + 1
    n_obs_downside = cfg["offset_downside"] + 1

    for i in range(len(retours)):
        if i + n_obs_vol <= len(retours):
            fenetre = retours.iloc[i : i + n_obs_vol]
            vol_6m.iloc[i] = fenetre.std(ddof=1) * math.sqrt(12)

        if i + n_obs_downside <= len(retours):
            fenetre = retours.iloc[i : i + n_obs_downside]

            for secteur in SECTEURS:
                negatifs = fenetre[secteur][fenetre[secteur] < 0].dropna()
                if len(negatifs) >= 2:
                    vol_down_18m.at[retours.index[i], secteur] = (
                        negatifs.std(ddof=1) * math.sqrt(12)
                    )

    score_vol = calculer_variable_historique(
        vol_6m,
        ordre_rank=0,
        moyenne_sans_finance=False,
        mix_rank_transversal=False,
    )
    score_down = calculer_variable_historique(
        vol_down_18m,
        ordre_rank=0,
        moyenne_sans_finance=False,
        mix_rank_transversal=False,
    )

    volatility = (score_vol + score_down) / 2

    sous_scores = {
        "volatility_6m": score_vol,
        "downside_volatility_18m": score_down,
    }

    return volatility, sous_scores, retours


# Agrégation et recommandations

def aligner_piliers(piliers):
    """
    Aligne les piliers par mois calendaire.

    Toute date est d'abord ramenée au dernier jour de son mois.
    Ainsi, 28/08, 29/08 et 31/08 représentent tous 31/08.
    """
    piliers_normalises = {
        nom: normaliser_index_mensuel(df)
        for nom, df in piliers.items()
    }

    dates = None

    for df in piliers_normalises.values():
        dates = (
            df.index
            if dates is None
            else dates.intersection(df.index)
        )

    if dates is None or len(dates) == 0:
        raise ValueError("Aucune date mensuelle commune entre les piliers.")

    dates = dates.sort_values(ascending=False)

    return {
        nom: df.loc[dates].copy()
        for nom, df in piliers_normalises.items()
    }


def calculer_rangs_piliers(piliers):
    """Classe les secteurs dans chacun des six piliers."""
    rangs = {}

    for nom, df in piliers.items():
        rang_df = pd.DataFrame(index=df.index, columns=SECTEURS, dtype=float)
        for date in df.index:
            rang_df.loc[date] = rang_secteurs(df.loc[date])
        rangs[nom] = rang_df

    return rangs


def calculer_score_pondere(rangs, poids, date):
    """Agrège les rangs des piliers avec les poids fournis."""
    resultat = pd.Series(0.0, index=SECTEURS)
    valide = pd.Series(True, index=SECTEURS)

    for pilier, poids_pilier in poids.items():
        serie = rangs[pilier].loc[date]
        valide &= serie.notna()
        resultat += serie.fillna(0) * poids_pilier

    resultat[~valide] = np.nan
    return resultat


def choisir_top_worst(top_count, bottom_count, rang_global):
    """
    Remplacement transparent de la fonction XLL tab_invest_eu.

    Logique visible dans Excel :
    - un secteur Bottom >= 3 ne peut pas être Positive ;
    - un secteur Top >= 3 ne peut pas être Negative ;
    - priorité au nombre de confirmations ;
    - le rang global sert de départage.
    """
    tableau = pd.DataFrame(
        {
            "top_count": top_count,
            "bottom_count": bottom_count,
            "rang_global": rang_global,
        }
    )

    candidats_top = tableau[tableau["bottom_count"] < N_WORST].copy()
    candidats_top = candidats_top.sort_values(
        ["top_count", "rang_global", "bottom_count"],
        ascending=[False, False, True],
    )
    top = list(candidats_top.head(N_TOP).index)

    candidats_worst = tableau[tableau["top_count"] < N_TOP].copy()
    candidats_worst = candidats_worst.sort_values(
        ["bottom_count", "rang_global", "top_count"],
        ascending=[False, True, True],
    )
    worst = list(candidats_worst.head(N_WORST).index)

    return top, worst


def calculer_resultats_finaux(piliers, contexte_macro):
    """Construit les scores finaux, les votes Top/Worst et les recommandations."""
    piliers = aligner_piliers(piliers)
    rangs = calculer_rangs_piliers(piliers)

    lignes = []

    for date in piliers["Leverage"].index:
        # Le modèle final a besoin du macro externe et du rate overlay.
        if date not in contexte_macro.index:
            continue

        regime = contexte_macro.at[date, "cycle"]
        macro_score = contexte_macro.at[date, "macro_score"]
        signal_taux = contexte_macro.at[date, "signal_taux"]

        if regime not in POIDS_REGIME:
            continue
        if signal_taux not in {"On", "Off"}:
            continue

        # Score 5F de base : six piliers équipondérés.
        score_base = calculer_score_pondere(rangs, POIDS_BASE, date)
        rang_base = rang_secteurs(score_base)

        # Score Tilt : poids dépendants du régime macro.
        score_tilt = calculer_score_pondere(rangs, POIDS_REGIME[regime], date)
        rang_tilt = rang_secteurs(score_tilt)

        # "+ macro" dans Excel :
        # le rang Tilt devient une composante supplémentaire.
        n_composantes_macro = len(POIDS_BASE) + 1
        score_macro = pd.Series(0.0, index=SECTEURS)
        valide_macro = pd.Series(True, index=SECTEURS)

        for pilier in POIDS_BASE:
            serie = rangs[pilier].loc[date]
            valide_macro &= serie.notna()
            score_macro += serie.fillna(0) / n_composantes_macro

        valide_macro &= rang_tilt.notna()
        score_macro += rang_tilt.fillna(0) / n_composantes_macro
        score_macro[~valide_macro] = np.nan

        rang_global = rang_secteurs(score_macro)

        # Comptage Top 3 / Worst 3.
        top_count = pd.Series(0, index=SECTEURS, dtype=int)
        bottom_count = pd.Series(0, index=SECTEURS, dtype=int)

        for pilier in POIDS_BASE:
            r = rangs[pilier].loc[date]
            top_count += (r > N_SECTEURS - N_TOP).fillna(False).astype(int)
            bottom_count += (r <= N_WORST).fillna(False).astype(int)

        # Le vote macro utilise le multiplicateur défini dans la configuration.
        top_count += POIDS_VOTE_MACRO * (
            rang_tilt > N_SECTEURS - N_TOP
        ).fillna(False).astype(int)
        bottom_count += POIDS_VOTE_MACRO * (
            rang_tilt <= N_WORST
        ).fillna(False).astype(int)

        # Le rate overlay renforce les piliers définis dans la configuration.
        if signal_taux == "On":
            for pilier in PILIERS_RATE_OVERLAY:
                r = rangs[pilier].loc[date]
                top_count += (r > N_SECTEURS - N_TOP).fillna(False).astype(int)
                bottom_count += (r <= N_WORST).fillna(False).astype(int)

        # Certains secteurs sont explicitement exclus du vote qualitatif.
        for secteur_exclu in SECTEURS_SANS_VOTE_TOP_WORST:
            top_count[secteur_exclu] = 0
            bottom_count[secteur_exclu] = 0

        top, worst = choisir_top_worst(top_count, bottom_count, rang_global)

        for secteur in SECTEURS:
            if secteur in top:
                reco = "Positive"
            elif secteur in worst:
                reco = "Negative"
            else:
                reco = "Neutral"

            ligne = {
                "date": date,
                "secteur": secteur,
                "cycle": regime,
                "macro_score": macro_score,
                "signal_taux": signal_taux,
                "score_base": score_base[secteur],
                "rang_base": rang_base[secteur],
                "score_tilt": score_tilt[secteur],
                "rang_tilt_macro": rang_tilt[secteur],
                "score_global_macro": score_macro[secteur],
                "rang_global": rang_global[secteur],
                "top_count": int(top_count[secteur]),
                "bottom_count": int(bottom_count[secteur]),
                "recommendation": reco,
            }

            for pilier in rangs:
                ligne[f"score_{pilier.lower()}"] = piliers[pilier].at[date, secteur]
                ligne[f"rang_{pilier.lower()}"] = rangs[pilier].at[date, secteur]

            lignes.append(ligne)

    return pd.DataFrame(lignes), piliers, rangs, contexte_macro


# Exécution du modèle

def calculer_modele(
    fichier_excel=None,
    fichier_macro=None,
):
    """Exécute l'ensemble du modèle et retourne les résultats intermédiaires."""
    wb_eu, wb_macro = ouvrir_workbooks(
        fichier_excel,
        fichier_macro,
    )

    try:
        piliers, sous_scores = calculer_piliers_historiques(wb_eu)

        momentum, sous_momentum = calculer_momentum(wb_eu)
        volatility, sous_vol, retours = calculer_volatilite(wb_eu)

        piliers["Momentum"] = momentum
        piliers["Volatility"] = volatility

        sous_scores.update(sous_momentum)
        sous_scores.update(sous_vol)

        contexte_macro = construire_contexte_macro(wb_eu, wb_macro)

        historique, piliers, rangs, contexte_macro = calculer_resultats_finaux(
            piliers,
            contexte_macro,
        )

        return {
            "historique": historique,
            "piliers": piliers,
            "rangs": rangs,
            "sous_scores": sous_scores,
            "retours": retours,
            "contexte_macro": contexte_macro,
        }
    finally:
        wb_eu.close()
        wb_macro.close()


def sauvegarder_sorties(resultats, dossier_sortie):
    """Sauvegarde les fichiers CSV de contrôle et de résultat."""
    dossier = Path(dossier_sortie)
    dossier.mkdir(parents=True, exist_ok=True)

    historique = resultats["historique"].copy()
    historique.to_csv(dossier / "eu_historique_modele.csv", index=False)

    # Les six piliers peuvent être plus récents que le régime macro.
    # On sauvegarde donc toujours leur dernière date disponible.
    piliers = resultats["piliers"]
    rangs = resultats["rangs"]

    dates_piliers = None
    for df in piliers.values():
        dates_piliers = (
            df.index
            if dates_piliers is None
            else dates_piliers.intersection(df.index)
        )

    if dates_piliers is not None and len(dates_piliers):
        date_piliers = dates_piliers.max()
        lignes_piliers = []

        for secteur in SECTEURS:
            ligne = {
                "date": date_piliers,
                "secteur": secteur,
            }

            for pilier in [
                "Leverage",
                "Margin",
                "Value",
                "Momentum",
                "Growth",
                "Volatility",
            ]:
                ligne[f"score_{pilier.lower()}"] = piliers[pilier].at[
                    date_piliers, secteur
                ]
                ligne[f"rang_{pilier.lower()}"] = rangs[pilier].at[
                    date_piliers, secteur
                ]

            lignes_piliers.append(ligne)

        pd.DataFrame(lignes_piliers).to_csv(
            dossier / "eu_piliers_latest_available.csv",
            index=False,
        )

    if historique.empty:
        return

    date_latest = historique["date"].max()
    latest = historique[historique["date"] == date_latest].copy()
    latest = latest.sort_values("rang_global", ascending=False)
    latest.to_csv(dossier / "eu_recommandations_latest.csv", index=False)

    colonnes = [
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
    latest[colonnes].to_csv(dossier / "eu_piliers_latest_with_reco.csv", index=False)


def afficher_latest(resultats):
    """Affiche la dernière recommandation disponible."""
    historique = resultats["historique"]

    if historique.empty:
        print("Aucun résultat calculé.")
        return

    date_latest = historique["date"].max()
    latest = historique[historique["date"] == date_latest].copy()
    latest = latest.sort_values("rang_global", ascending=False)

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

    print(f"\nDate de la recommandation finale : {date_latest.date()}")
    print(latest[cols].to_string(index=False))

    # Information importante si le fichier macro est moins récent.
    dates_piliers = None
    for df in resultats["piliers"].values():
        dates_piliers = (
            df.index
            if dates_piliers is None
            else dates_piliers.intersection(df.index)
        )

    if dates_piliers is not None and len(dates_piliers):
        date_piliers = dates_piliers.max()
        if date_piliers > date_latest:
            print(
                f"\nAttention : les piliers sont disponibles jusqu'au "
                f"{date_piliers.date()}, mais le régime macro permet une "
                f"recommandation finale seulement jusqu'au {date_latest.date()}."
            )


def main():
    resultats = calculer_modele()
    sauvegarder_sorties(resultats, DOSSIER_SORTIE_PAR_DEFAUT)
    afficher_latest(resultats)


if __name__ == "__main__":
    main()
