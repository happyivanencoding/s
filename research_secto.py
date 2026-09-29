# -*- coding: utf-8 -*-
"""Outils de recherche pour les modèles sectoriels EU et US."""

from pathlib import Path
import math

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from local_config import MODELE_EU, MODELE_US
import model_secto_eu
import model_secto_us


MARCHES = {
    "EU": (MODELE_EU, model_secto_eu),
    "US": (MODELE_US, model_secto_us),
}


def preparer_retour_futur(retours):
    """Aligne le rendement du mois suivant avec le signal courant."""
    return retours.sort_index().shift(-1)


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

        lignes.append(
            {
                "date_signal": date,
                "return_top": ret_top,
                "return_worst": ret_worst,
                "return_long_short": ret_top - ret_worst,
                "top_sectors": " | ".join(top),
                "worst_sectors": " | ".join(worst),
            }
        )

    return pd.DataFrame(lignes)


def max_drawdown(returns):
    """Calcule le maximum drawdown."""
    if len(returns) == 0:
        return np.nan

    wealth = (1 + returns.fillna(0)).cumprod()
    drawdown = wealth / wealth.cummax() - 1

    return drawdown.min()


def statistiques(backtest):
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


class RechercheSectorielle:
    """Charge un marché une fois puis permet de lancer plots et backtests."""

    def __init__(
        self,
        marche="EU",
        fichier_excel=None,
        fichier_macro=None,
    ):
        marche = marche.upper()

        if marche not in MARCHES:
            raise ValueError("Le marché doit être EU ou US.")

        self.marche = marche
        self.config, self.modele = MARCHES[marche]
        self.fichier_excel = fichier_excel
        self.fichier_macro = fichier_macro
        self.resultats = None

    def charger(self, forcer=False):
        """Calcule le modèle une seule fois et conserve les résultats en mémoire."""
        if self.resultats is None or forcer:
            self.resultats = self.modele.calculer_modele(
                self.fichier_excel,
                fichier_macro=self.fichier_macro,
            )

        return self

    def variables_disponibles(self, pilier=None):
        """Retourne les variables disponibles pour un pilier ou pour tous."""
        groupes = self.config["variables_backtest"]

        if pilier is None:
            return {
                nom: variables.copy()
                for nom, variables in groupes.items()
            }

        if pilier not in groupes:
            raise ValueError(
                f"Pilier inconnu : {pilier}"
            )

        return groupes[pilier].copy()

    def _resultats(self):
        if self.resultats is None:
            self.charger()

        return self.resultats

    def _date_piliers(self, date=None):
        """Choisit une date disponible commune aux piliers."""
        piliers = self._resultats()["piliers"]
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

        if date is None:
            return dates[-1]

        cible = (
            pd.Timestamp(date)
            .to_period("M")
            .to_timestamp("M")
        )
        dates_avant = dates[dates <= cible]

        if len(dates_avant) == 0:
            raise ValueError(
                f"Aucune date disponible avant {cible.date()}."
            )

        return dates_avant[-1]

    def plot_piliers(
        self,
        date=None,
        piliers=None,
        figsize=(10, 7),
    ):
        """Affiche les scores des piliers par secteur."""
        resultats = self._resultats()
        date = self._date_piliers(date)

        if piliers is None:
            piliers = list(self.config["poids_base"])

        inconnus = [
            p
            for p in piliers
            if p not in resultats["piliers"]
        ]

        if inconnus:
            raise ValueError(
                f"Piliers inconnus : {inconnus}"
            )

        secteurs = self.config["secteurs"]
        matrice = np.array(
            [
                [
                    resultats["piliers"][p].at[date, s]
                    for p in piliers
                ]
                for s in secteurs
            ],
            dtype=float,
        )

        vmax = max(10, np.nanmax(matrice))

        fig, ax = plt.subplots(figsize=figsize)
        image = ax.imshow(
            matrice,
            aspect="auto",
            vmin=0,
            vmax=vmax,
        )

        ax.set_xticks(range(len(piliers)))
        ax.set_xticklabels(piliers)
        ax.set_yticks(range(len(secteurs)))
        ax.set_yticklabels(secteurs)
        ax.set_title(
            f"{self.marche} - Scores des piliers au {date.date()}"
        )

        for i in range(len(secteurs)):
            for j in range(len(piliers)):
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

        fig.colorbar(image, ax=ax, label="Score")
        fig.tight_layout()

        return fig, ax

    def plot_variables(
        self,
        pilier,
        variables=None,
        date=None,
        figsize=(9, 7),
    ):
        """Affiche les sous-variables sélectionnées d'un pilier."""
        resultats = self._resultats()
        date = self._date_piliers(date)

        disponibles = self.variables_disponibles(pilier)

        if variables is None:
            variables = disponibles

        inconnues = [
            v
            for v in variables
            if v not in disponibles
        ]

        if inconnues:
            raise ValueError(
                f"Variables inconnues pour {pilier} : {inconnues}"
            )

        secteurs = self.config["secteurs"]
        matrice = []

        for secteur in secteurs:
            ligne = []

            for variable in variables:
                df = resultats["sous_scores"][variable]
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
        vmax = max(10, np.nanmax(matrice))

        fig, ax = plt.subplots(figsize=figsize)
        image = ax.imshow(
            matrice,
            aspect="auto",
            vmin=0,
            vmax=vmax,
        )

        ax.set_xticks(range(len(variables)))
        ax.set_xticklabels(
            variables,
            rotation=20,
            ha="right",
        )
        ax.set_yticks(range(len(secteurs)))
        ax.set_yticklabels(secteurs)
        ax.set_title(
            f"{self.marche} - {pilier} au {date.date()}"
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

        fig.colorbar(image, ax=ax, label="Score")
        fig.tight_layout()

        return fig, ax

    def plot_rang_global(
        self,
        date=None,
        figsize=(9, 6),
    ):
        """Affiche le rang global pour une date de recommandation."""
        historique = self._resultats()["historique"]

        if historique.empty:
            raise ValueError(
                "Aucune recommandation finale disponible."
            )

        dates = (
            historique["date"]
            .drop_duplicates()
            .sort_values()
        )

        if date is None:
            date = dates.iloc[-1]
        else:
            cible = (
                pd.Timestamp(date)
                .to_period("M")
                .to_timestamp("M")
            )
            possibles = dates[dates <= cible]

            if len(possibles) == 0:
                raise ValueError(
                    f"Aucune recommandation avant {cible.date()}."
                )

            date = possibles.iloc[-1]

        ligne = historique[
            historique["date"] == date
        ].sort_values("rang_global")

        fig, ax = plt.subplots(figsize=figsize)
        ax.barh(
            ligne["secteur"],
            ligne["rang_global"],
        )
        ax.set_xlabel(
            f"Rang final - 1 = Worst, "
            f"{len(self.config['secteurs'])} = Best"
        )
        ax.set_title(
            f"{self.marche} - Rang final au {date.date()}"
        )
        fig.tight_layout()

        return fig, ax

    def backtest(
        self,
        pilier,
        variables=None,
        inclure_pilier=True,
        top_n=3,
        start=None,
        end=None,
        avec_plot=True,
        figsize=(10, 6),
    ):
        """Backteste les variables sélectionnées et, si demandé, le pilier final."""
        resultats = self._resultats()
        disponibles = self.variables_disponibles(pilier)

        if variables is None:
            variables = disponibles

        inconnues = [
            v
            for v in variables
            if v not in disponibles
        ]

        if inconnues:
            raise ValueError(
                f"Variables inconnues pour {pilier} : {inconnues}"
            )

        retours_futurs = preparer_retour_futur(
            resultats["retours"]
        )

        scores = {
            variable: resultats["sous_scores"][variable]
            for variable in variables
        }

        if inclure_pilier:
            scores[f"{pilier}_total"] = (
                resultats["piliers"][pilier]
            )

        stats = []
        backtests = {}

        for nom, score in scores.items():
            bt = backtester_score(
                score,
                retours_futurs,
                top_n=top_n,
                start=start,
                end=end,
            )

            backtests[nom] = bt

            ligne = {"variable": nom}
            ligne.update(statistiques(bt))
            stats.append(ligne)

        resume = (
            pd.DataFrame(stats)
            .sort_values(
                "sharpe_ls",
                ascending=False,
            )
            .reset_index(drop=True)
        )

        fig = None
        ax = None

        if avec_plot:
            fig, ax = plt.subplots(figsize=figsize)

            for nom, bt in backtests.items():
                if bt.empty:
                    continue

                serie = bt.set_index(
                    "date_signal"
                )["return_long_short"]

                cumul = (1 + serie).cumprod()

                ax.plot(
                    cumul.index,
                    cumul.values,
                    label=nom,
                )

            ax.set_title(
                f"{self.marche} - {pilier} : backtest Long-Short"
            )
            ax.set_ylabel(
                "Valeur cumulée - base 1"
            )
            ax.legend()
            ax.grid(alpha=0.2)
            fig.tight_layout()

        return resume, backtests, fig, ax

    def sauvegarder_backtest(
        self,
        resume,
        backtests,
        pilier,
        dossier="output/backtests",
    ):
        """Sauvegarde un backtest déjà calculé."""
        dossier = Path(dossier)
        dossier.mkdir(
            parents=True,
            exist_ok=True,
        )

        prefixe = self.config["prefixe_sortie"]

        resume.to_csv(
            dossier
            / f"{prefixe}_{pilier.lower()}_summary.csv",
            index=False,
        )

        for nom, bt in backtests.items():
            bt.to_csv(
                dossier
                / (
                    f"{prefixe}_{pilier.lower()}_"
                    f"{nom}_monthly.csv"
                ),
                index=False,
            )
