# -*- coding: utf-8 -*-
"""Calculs communs aux modèles sectoriels."""

from pathlib import Path
import math

import numpy as np
import pandas as pd

from data_io import (
    lire_benchmark_returns,
    lire_bloc_returns,
    lire_dates_et_bloc,
    lire_macro_externe,
    lire_taux_us10y,
    normaliser_index_mensuel,
)


def rang_excel(valeur, valeurs, ordre):
    """Reproduit RANK Excel. ordre=1 croissant, ordre=0 décroissant."""
    if pd.isna(valeur):
        return np.nan

    serie = pd.Series(valeurs, dtype="float64").dropna()
    if serie.empty:
        return np.nan

    if ordre == 1:
        return 1 + int((serie < valeur).sum())

    return 1 + int((serie > valeur).sum())


def rang_transversal(ligne, ordre, echelle_0_10=True):
    """Classe les secteurs au même mois."""
    resultat = pd.Series(index=ligne.index, dtype=float)
    n_secteurs = len(ligne.index)

    for secteur in ligne.index:
        valeur = ligne[secteur]

        if pd.isna(valeur):
            resultat[secteur] = np.nan
            continue

        rang = rang_excel(valeur, ligne.values, ordre)

        if echelle_0_10:
            resultat[secteur] = rang / n_secteurs * 10
        else:
            resultat[secteur] = rang

    return resultat


def rang_secteurs(ligne):
    """Classe les secteurs de 1 = Worst à N = Best."""
    resultat = pd.Series(index=ligne.index, dtype=float)

    for secteur in ligne.index:
        valeur = ligne[secteur]

        if pd.isna(valeur):
            resultat[secteur] = np.nan
            continue

        resultat[secteur] = rang_excel(
            valeur,
            ligne.values,
            ordre=1,
        )

    return resultat


def percentrank_inc(valeurs, x):
    """Reproduit PERCENTRANK.INC d'Excel."""
    valeurs = (
        pd.Series(valeurs, dtype=float)
        .dropna()
        .sort_values()
        .to_numpy()
    )

    if len(valeurs) == 0 or pd.isna(x):
        return np.nan
    if len(valeurs) == 1:
        return 1.0
    if x <= valeurs[0]:
        return 0.0
    if x >= valeurs[-1]:
        return 1.0

    positions = np.where(
        np.isclose(valeurs, x, rtol=0, atol=1e-12)
    )[0]

    if len(positions):
        return positions[0] / (len(valeurs) - 1)

    droite = np.searchsorted(valeurs, x, side="right")
    gauche = droite - 1
    x0 = valeurs[gauche]
    x1 = valeurs[droite]
    fraction = (x - x0) / (x1 - x0)

    return (gauche + fraction) / (len(valeurs) - 1)


def calculer_signal_taux(taux_us10y, config_signal):
    """Calcule le rate overlay à partir du US 10Y."""
    alpha = config_signal["alpha_ewma"]
    seuil = config_signal["seuil_percentile"]

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


def construire_contexte_macro(wb_secteur, wb_macro, config):
    """Assemble macro score, régime et rate overlay."""
    macro = lire_macro_externe(
        wb_macro,
        config_macro=config["macro"],
        poids_regime=config["poids_regime"],
        cle_prefixe=config["historique_prefixe"],
    )

    taux = lire_taux_us10y(
        wb_secteur,
        config_signal=config["signal_taux"],
        cle_prefixe=config["historique_prefixe"],
    )

    signal_taux = calculer_signal_taux(
        taux,
        config["signal_taux"],
    )

    contexte = macro.join(
        signal_taux[["signal_taux"]],
        how="outer",
    )

    return contexte.sort_index()


def calculer_diff_vs_moyenne(raw, moyenne_sans_finance=False):
    """Calcule l'écart du secteur à la moyenne du même mois."""
    diff = pd.DataFrame(
        index=raw.index,
        columns=raw.columns,
        dtype=float,
    )

    for date, ligne in raw.iterrows():
        base = ligne.copy()

        if moyenne_sans_finance:
            base = base.drop("Fin")

        moyenne = base.mean(skipna=True)
        diff.loc[date] = ligne - moyenne

    return diff


def calculer_score_historique(
    diff,
    ordre_rank,
    fenetre_historique,
    fenetre_par_secteur=None,
):
    """Classe l'écart courant dans son historique mensuel."""
    fenetre_par_secteur = fenetre_par_secteur or {}

    score = pd.DataFrame(
        index=diff.index,
        columns=diff.columns,
        dtype=float,
    )

    for i in range(len(diff)):
        for secteur in diff.columns:
            valeur = diff.iloc[i][secteur]

            if pd.isna(valeur):
                continue

            n_mois = fenetre_par_secteur.get(
                secteur,
                fenetre_historique,
            )
            n_observations = n_mois + 1
            fin = i + n_observations

            if fin > len(diff):
                continue

            valeurs = diff.iloc[i:fin][secteur]
            rang = rang_excel(
                valeur,
                valeurs.values,
                ordre_rank,
            )

            score.iloc[
                i,
                score.columns.get_loc(secteur),
            ] = 10 * rang / n_observations

    return score


def calculer_variable_historique(raw, config_variable, fenetre_historique):
    """Calcule le score d'une sous-variable."""
    diff = calculer_diff_vs_moyenne(
        raw,
        config_variable["moyenne_sans_finance"],
    )

    score_hist = calculer_score_historique(
        diff,
        config_variable["ordre_rank"],
        fenetre_historique,
        config_variable.get("fenetre_par_secteur"),
    )

    if not config_variable["mix_rank_transversal"]:
        return score_hist

    score = score_hist.copy()

    for date in score.index:
        score_cross = rang_transversal(
            raw.loc[date],
            config_variable["ordre_rank"],
            echelle_0_10=True,
        )
        score.loc[date] = (
            score_hist.loc[date] + score_cross
        ) / 2

    return score


def compose_pillar(scores, config, pillar):
    """Agrège les sous-scores selon la composition du pilier."""
    sectors = config["secteurs"]
    composition = config["composition_piliers"][pillar]

    dates = None
    for score in scores.values():
        dates = (
            score.index
            if dates is None
            else dates.intersection(score.index)
        )

    if dates is None or len(dates) == 0:
        raise ValueError(
            f"Aucune date disponible pour le pilier {pillar}."
        )

    result = pd.DataFrame(
        index=dates,
        columns=sectors,
        dtype=float,
    )

    by_sector = composition.get("par_secteur", {})

    for date in dates:
        for sector in sectors:
            variables = by_sector.get(
                sector,
                composition["default"],
            )
            values = [
                scores[variable].at[date, sector]
                for variable in variables
            ]

            if any(pd.isna(value) for value in values):
                result.at[date, sector] = np.nan
            else:
                result.at[date, sector] = (
                    sum(values) / len(values)
                )

    decimals = config["arrondi_pilier"].get(pillar)

    if decimals is not None:
        result = result.round(decimals)

    return result


def calculer_piliers_historiques(wb, config):
    """Calcule Leverage, Margin, Value et Growth."""
    secteurs = config["secteurs"]
    sous_scores = {}
    piliers = {}

    for pilier, variables in config["variables_historiques"].items():
        scores_du_pilier = {}

        for nom_variable, cfg in variables.items():
            ws = wb[cfg["sheet"]]
            raw = lire_dates_et_bloc(
                ws,
                cfg["colonne"],
                secteurs=secteurs,
                cle_prefixe=config["historique_prefixe"],
            )

            score = calculer_variable_historique(
                raw,
                cfg,
                config["fenetre_historique"],
            )

            sous_scores[nom_variable] = score
            scores_du_pilier[nom_variable] = score

        piliers[pilier] = compose_pillar(
            scores_du_pilier,
            config,
            pilier,
        )

    return piliers, sous_scores


def calculer_momentum(wb, config):
    """Calcule les trois composantes Momentum."""
    cfg = config["momentum"]
    secteurs = config["secteurs"]
    ws = wb[cfg["sheet"]]

    def lire(colonne):
        return lire_dates_et_bloc(
            ws,
            colonne,
            secteurs=secteurs,
            cle_prefixe=config["historique_prefixe"],
        )

    prix = lire(cfg["colonne_prix"])
    up = lire(cfg["colonne_revision_up"])
    down = lire(cfg["colonne_revision_down"])
    unchanged = lire(cfg["colonne_revision_unchanged"])

    horizon_court = cfg["horizon_court"]
    horizon_long = cfg["horizon_long"]
    secteur_mois_courant = cfg.get("secteur_mois_courant")

    ret_court = pd.DataFrame(
        index=prix.index,
        columns=secteurs,
        dtype=float,
    )
    ret_long = pd.DataFrame(
        index=prix.index,
        columns=secteurs,
        dtype=float,
    )

    for i in range(len(prix)):
        if i + horizon_court < len(prix):
            ret_court.iloc[i] = (
                prix.iloc[i + 1]
                / prix.iloc[i + horizon_court]
                - 1
            )

        if i + horizon_long < len(prix):
            ret_long.iloc[i] = (
                prix.iloc[i + 1]
                / prix.iloc[i + horizon_long]
                - 1
            )

        if secteur_mois_courant:
            if i + horizon_court < len(prix):
                ret_court.at[
                    prix.index[i],
                    secteur_mois_courant,
                ] = (
                    prix.iloc[i][secteur_mois_courant]
                    / prix.iloc[i + horizon_court][secteur_mois_courant]
                    - 1
                )

            if i + horizon_long < len(prix):
                ret_long.at[
                    prix.index[i],
                    secteur_mois_courant,
                ] = (
                    prix.iloc[i][secteur_mois_courant]
                    / prix.iloc[i + horizon_long][secteur_mois_courant]
                    - 1
                )

    revision_ratio = (
        (up - down)
        / (up + down + unchanged)
    )

    echelle = cfg.get("echelle_0_10", True)

    score_court = ret_court.apply(
        lambda ligne: rang_transversal(
            ligne,
            ordre=1,
            echelle_0_10=echelle,
        ),
        axis=1,
    )
    score_long = ret_long.apply(
        lambda ligne: rang_transversal(
            ligne,
            ordre=1,
            echelle_0_10=echelle,
        ),
        axis=1,
    )
    score_revision = revision_ratio.apply(
        lambda ligne: rang_transversal(
            ligne,
            ordre=1,
            echelle_0_10=echelle,
        ),
        axis=1,
    )

    sous_scores = {
        "momentum_6m_1m": score_court,
        "momentum_12m_1m": score_long,
        "earnings_revision_ratio": score_revision,
    }

    momentum = compose_pillar(
        sous_scores,
        config,
        "Momentum",
    )

    return momentum, sous_scores


def calculer_volatilite(wb, config):
    """Calcule total volatility et downside volatility."""
    cfg = config["volatilite"]
    secteurs = config["secteurs"]

    returns = lire_bloc_returns(
        wb[cfg["sheet_returns"]],
        config_volatilite=cfg,
        secteurs=secteurs,
        cle_prefixe=config["historique_prefixe"],
    )

    vol_total = pd.DataFrame(
        index=returns.index,
        columns=secteurs,
        dtype=float,
    )
    vol_downside = pd.DataFrame(
        index=returns.index,
        columns=secteurs,
        dtype=float,
    )

    n_obs_vol = cfg["offset_volatilite"] + 1
    n_obs_downside = cfg["offset_downside"] + 1

    for i in range(len(returns)):
        if i + n_obs_vol <= len(returns):
            fenetre = returns.iloc[i : i + n_obs_vol]
            vol_total.iloc[i] = (
                fenetre.std(ddof=1) * math.sqrt(12)
            )

        if i + n_obs_downside <= len(returns):
            fenetre = returns.iloc[i : i + n_obs_downside]

            for secteur in secteurs:
                negatifs = fenetre[secteur][
                    fenetre[secteur] < 0
                ].dropna()

                if len(negatifs) >= 2:
                    vol_downside.at[
                        returns.index[i],
                        secteur,
                    ] = (
                        negatifs.std(ddof=1)
                        * math.sqrt(12)
                    )

    cfg_score = {
        "ordre_rank": 0,
        "moyenne_sans_finance": False,
        "mix_rank_transversal": False,
    }

    score_total = calculer_variable_historique(
        vol_total,
        cfg_score,
        config["fenetre_historique"],
    )
    score_downside = calculer_variable_historique(
        vol_downside,
        cfg_score,
        config["fenetre_historique"],
    )

    sous_scores = {
        "volatility_6m": score_total,
        "downside_volatility_18m": score_downside,
    }

    volatility = compose_pillar(
        sous_scores,
        config,
        "Volatility",
    )

    return volatility, sous_scores, returns


def aligner_piliers(piliers):
    """Aligne les piliers sur une clé mensuelle commune."""
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
        raise ValueError(
            "Aucune date mensuelle commune entre les piliers."
        )

    dates = dates.sort_values(ascending=False)

    return {
        nom: df.loc[dates].copy()
        for nom, df in piliers_normalises.items()
    }


def calculer_rangs_piliers(piliers):
    """Classe les secteurs dans chaque pilier."""
    rangs = {}

    for nom, df in piliers.items():
        rang_df = pd.DataFrame(
            index=df.index,
            columns=df.columns,
            dtype=float,
        )

        for date in df.index:
            rang_df.loc[date] = rang_secteurs(
                df.loc[date]
            )

        rangs[nom] = rang_df

    return rangs


def calculer_score_pondere(rangs, poids, date):
    """Agrège les rangs des piliers avec les poids fournis."""
    secteurs = next(iter(rangs.values())).columns

    resultat = pd.Series(
        0.0,
        index=secteurs,
    )
    valide = pd.Series(
        True,
        index=secteurs,
    )

    for pilier, poids_pilier in poids.items():
        serie = rangs[pilier].loc[date]
        valide &= serie.notna()
        resultat += serie.fillna(0) * poids_pilier

    resultat[~valide] = np.nan

    return resultat


def choisir_top_worst(
    top_count,
    bottom_count,
    rang_global,
    n_top,
    n_worst,
):
    """Sélectionne les secteurs Positive et Negative."""
    tableau = pd.DataFrame(
        {
            "top_count": top_count,
            "bottom_count": bottom_count,
            "rang_global": rang_global,
        }
    )

    candidats_top = tableau[
        tableau["bottom_count"] < n_worst
    ].copy()

    candidats_top = candidats_top.sort_values(
        [
            "top_count",
            "rang_global",
            "bottom_count",
        ],
        ascending=[
            False,
            False,
            True,
        ],
    )

    top = list(
        candidats_top.head(n_top).index
    )

    candidats_worst = tableau[
        tableau["top_count"] < n_top
    ].copy()

    candidats_worst = candidats_worst.sort_values(
        [
            "bottom_count",
            "rang_global",
            "top_count",
        ],
        ascending=[
            False,
            True,
            True,
        ],
    )

    worst = list(
        candidats_worst.head(n_worst).index
    )

    return top, worst


def calculer_score_global_macro(rangs, rang_tilt, date, config):
    """Agrège les rangs avec la pondération macro du marché."""
    secteurs = config["secteurs"]
    poids = config["poids_global_macro"]

    score = pd.Series(
        0.0,
        index=secteurs,
    )
    valide = pd.Series(
        True,
        index=secteurs,
    )

    for pilier, poids_pilier in poids.items():
        if pilier == "Macro":
            continue

        serie = rangs[pilier].loc[date]
        valide &= serie.notna()
        score += serie.fillna(0) * poids_pilier

    poids_macro = poids.get("Macro", 0)

    if poids_macro:
        valide &= rang_tilt.notna()
        score += rang_tilt.fillna(0) * poids_macro

    score[~valide] = np.nan

    return score


def calculer_resultats_finaux(piliers, contexte_macro, config):
    """Construit scores finaux, votes et recommandations."""
    secteurs = config["secteurs"]
    n_secteurs = len(secteurs)
    n_top = config["n_top"]
    n_worst = config["n_worst"]

    piliers = aligner_piliers(piliers)
    rangs = calculer_rangs_piliers(piliers)
    lignes = []

    for date in piliers["Leverage"].index:
        if date not in contexte_macro.index:
            continue

        regime = contexte_macro.at[date, "cycle"]
        macro_score = contexte_macro.at[date, "macro_score"]
        signal_taux = contexte_macro.at[date, "signal_taux"]

        if regime not in config["poids_regime"]:
            continue
        if signal_taux not in {"On", "Off"}:
            continue

        score_base = calculer_score_pondere(
            rangs,
            config["poids_base"],
            date,
        )
        rang_base = rang_secteurs(score_base)

        score_tilt = calculer_score_pondere(
            rangs,
            config["poids_regime"][regime],
            date,
        )
        rang_tilt = rang_secteurs(score_tilt)

        score_macro = calculer_score_global_macro(
            rangs,
            rang_tilt,
            date,
            config,
        )
        rang_global = rang_secteurs(score_macro)

        top_count = pd.Series(
            0,
            index=secteurs,
            dtype=int,
        )
        bottom_count = pd.Series(
            0,
            index=secteurs,
            dtype=int,
        )

        for pilier in config["piliers_top_vote"]:
            r = rangs[pilier].loc[date]
            top_count += (
                r > n_secteurs - n_top
            ).fillna(False).astype(int)

        for pilier in config["piliers_bottom_vote"]:
            r = rangs[pilier].loc[date]
            bottom_count += (
                r <= n_worst
            ).fillna(False).astype(int)

        poids_macro = config["poids_vote_macro"]

        top_count += poids_macro * (
            rang_tilt > n_secteurs - n_top
        ).fillna(False).astype(int)

        bottom_count += poids_macro * (
            rang_tilt <= n_worst
        ).fillna(False).astype(int)

        if signal_taux == "On":
            for pilier in config["piliers_rate_overlay"]:
                r = rangs[pilier].loc[date]

                top_count += (
                    r > n_secteurs - n_top
                ).fillna(False).astype(int)

                bottom_count += (
                    r <= n_worst
                ).fillna(False).astype(int)

        for regle in config.get("regles_sans_vote", []):
            date_debut = pd.Timestamp(
                regle["date_debut"]
            ).to_period("M").to_timestamp("M")

            date_fin = regle.get("date_fin")
            actif = date >= date_debut

            if date_fin:
                date_fin = pd.Timestamp(
                    date_fin
                ).to_period("M").to_timestamp("M")
                actif = actif and date <= date_fin

            if actif:
                secteur = regle["secteur"]
                top_count[secteur] = 0
                bottom_count[secteur] = 0

        top, worst = choisir_top_worst(
            top_count,
            bottom_count,
            rang_global,
            n_top,
            n_worst,
        )

        for secteur in secteurs:
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
                ligne[f"score_{pilier.lower()}"] = (
                    piliers[pilier].at[date, secteur]
                )
                ligne[f"rang_{pilier.lower()}"] = (
                    rangs[pilier].at[date, secteur]
                )

            lignes.append(ligne)

    return (
        pd.DataFrame(lignes),
        piliers,
        rangs,
        contexte_macro,
    )


def executer_modele(wb_secteur, wb_macro, config):
    """Exécute les calculs communs du marché."""
    piliers, sous_scores = calculer_piliers_historiques(
        wb_secteur,
        config,
    )

    momentum, sous_momentum = calculer_momentum(
        wb_secteur,
        config,
    )
    volatility, sous_vol, returns = calculer_volatilite(
        wb_secteur,
        config,
    )
    benchmark_returns = lire_benchmark_returns(
        wb_secteur[config["volatilite"]["sheet_returns"]],
        config_volatilite=config["volatilite"],
        cle_prefixe=config["historique_prefixe"],
    )

    piliers["Momentum"] = momentum
    piliers["Volatility"] = volatility

    sous_scores.update(sous_momentum)
    sous_scores.update(sous_vol)

    contexte_macro = construire_contexte_macro(
        wb_secteur,
        wb_macro,
        config,
    )

    (
        historique,
        piliers,
        rangs,
        contexte_macro,
    ) = calculer_resultats_finaux(
        piliers,
        contexte_macro,
        config,
    )

    return {
        "history": historique,
        "pillars": piliers,
        "ranks": rangs,
        "variable_scores": sous_scores,
        "returns": returns,
        "benchmark_returns": benchmark_returns,
        "macro_context": contexte_macro,
    }


def sauvegarder_production(resultats, dossier_sortie, config):
    """Sauvegarde uniquement les sorties courantes du modèle."""
    dossier = Path(dossier_sortie)
    dossier.mkdir(parents=True, exist_ok=True)

    prefixe = config["prefixe_sortie"]
    historique = resultats["history"]
    piliers = resultats["pillars"]
    rangs = resultats["ranks"]

    dates_piliers = None

    for df in piliers.values():
        dates_piliers = (
            df.index
            if dates_piliers is None
            else dates_piliers.intersection(df.index)
        )

    if dates_piliers is not None and len(dates_piliers):
        date_piliers = dates_piliers.max()
        lignes = []

        for secteur in config["secteurs"]:
            ligne = {
                "date": date_piliers,
                "secteur": secteur,
            }

            for pilier in config["poids_base"]:
                ligne[f"score_{pilier.lower()}"] = (
                    piliers[pilier].at[date_piliers, secteur]
                )
                ligne[f"rang_{pilier.lower()}"] = (
                    rangs[pilier].at[date_piliers, secteur]
                )

            lignes.append(ligne)

        pd.DataFrame(lignes).to_csv(
            dossier
            / f"{prefixe}_piliers_latest_available.csv",
            index=False,
        )

    if historique.empty:
        return

    date_latest = historique["date"].max()
    latest = historique[
        historique["date"] == date_latest
    ].copy()
    latest = latest.sort_values(
        "rang_global",
        ascending=False,
    )

    latest.to_csv(
        dossier
        / f"{prefixe}_recommandations_latest.csv",
        index=False,
    )

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

    latest[colonnes].to_csv(
        dossier
        / f"{prefixe}_piliers_latest_with_reco.csv",
        index=False,
    )


def sauvegarder_sorties(resultats, dossier_sortie, config):
    """Sauvegarde l'historique complet et les sorties courantes."""
    dossier = Path(dossier_sortie)
    dossier.mkdir(parents=True, exist_ok=True)

    prefixe = config["prefixe_sortie"]

    resultats["history"].to_csv(
        dossier / f"{prefixe}_historique_modele.csv",
        index=False,
    )

    sauvegarder_production(
        resultats,
        dossier,
        config,
    )


def afficher_latest(resultats):
    """Affiche la dernière recommandation disponible."""
    historique = resultats["history"]

    if historique.empty:
        print("Aucun résultat calculé.")
        return

    date_latest = historique["date"].max()
    latest = historique[
        historique["date"] == date_latest
    ].copy()
    latest = latest.sort_values(
        "rang_global",
        ascending=False,
    )

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

    print(
        f"\nDate de la recommandation finale : "
        f"{date_latest.date()}"
    )
    print(latest[cols].to_string(index=False))

    dates_piliers = None
    for df in resultats["pillars"].values():
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
