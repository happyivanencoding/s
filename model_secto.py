# -*- coding: utf-8 -*-
"""Modèle sectoriel EU / US.

STRUCTURE DU FICHIER
1. SectorModel : choisir le marché et lancer le calcul.
2. run_model : dérouler les étapes du calcul pour un marché.
3. Sous-scores et piliers : transformer les mesures et les agréger.
4. Macro et taux : déterminer le contexte du mois.
5. Rangs et recommandations : compter les votes et choisir les secteurs.
6. Exports : écrire les tableaux de scores en Excel.

FLUX DES DONNÉES
python actualiser_excel.py -> actualiser_excel.main()
    -> local_config.py : chemins et paramètres
    -> data_utils : Excel et historique mensuel figé
    -> SectorModel.run() -> run_model()
         -> mesures -> sous-scores -> six piliers
         -> macro et taux -> rangs -> votes -> recommandations
    -> excel_output : valeurs, formats et graphiques dans les XLSM
    -> actualiser_excel : bilan de production

Exemple : SectorModel("EU").run_production()
Pour actualiser aussi les XLSM : lancer python actualiser_excel.py.
"""

from copy import deepcopy
from pathlib import Path
import math

import numpy as np
import pandas as pd

from local_config import EU_CONFIG, US_CONFIG, MACRO_EXCEL_FILE
from data_utils import (
    open_workbooks,
    load_history,
    history_path,
    save_history,
    save_table,
    read_benchmark_returns,
    read_sector_returns,
    read_sector_block,
    read_macro_data,
    read_us10y,
    normalize_monthly_index,
)

MODEL_CONFIGS = {"EU": EU_CONFIG, "US": US_CONFIG}

# 1. Interface du modèle
class SectorModel:
    """Configure et exécute un modèle sectoriel."""

    def __init__(
        self,
        universe="EU",
        excel_file=None,
        macro_file=None,
    ):
        universe = universe.upper()

        if universe not in MODEL_CONFIGS:
            raise ValueError("Universe must be EU or US.")

        configuration = MODEL_CONFIGS[universe]

        self.universe = universe
        self.config = deepcopy(configuration)
        self.excel_file = Path(excel_file or self.config["excel_file"])
        self.macro_file = Path(macro_file or MACRO_EXCEL_FILE)
        self.results = None

    # Les trois modes partagent le même calcul.
    def run(self, mode="production", history_file=None, fresh_history=False):
        """Calcule depuis Excel ou depuis les seules observations enregistrées."""
        if mode not in ("production", "rebuild", "replay"):
            raise ValueError(f"Mode inconnu : {mode}")
        if fresh_history and mode != "rebuild":
            raise ValueError("Seul rebuild peut commencer avec un historique vide.")
        if mode == "rebuild" and history_file is None:
            raise ValueError("rebuild nécessite un historique dans un dossier de sortie.")

        if mode == "replay":
            source = Path(history_file) if history_file else history_path()
            if not source.is_file():
                raise FileNotFoundError(f"Historique introuvable : {source}")
            load_history(reload=True, file_path=source)
            self.results = run_model(None, None, self.config)
            return self.results

        sector_workbook, macro_workbook = open_workbooks(self.excel_file, self.macro_file)
        load_history(reload=True, file_path=history_file, fresh=fresh_history)
        try:
            self.results = run_model(sector_workbook, macro_workbook, self.config)
            save_history()
        finally:
            sector_workbook.close()
            macro_workbook.close()
        return self.results

    def run_production(self, output_dir="output"):
        """Exécute le modèle et exporte les tableaux Excel."""
        results = self.run()

        save_production(results, Path(output_dir), self.config)
        from excel_output import save_excel

        save_excel(results, output_dir, self.config)
        print_latest(results)

        return results

    # Paramètres de base du modèle.
    @property
    def sectors(self):
        """Retourne les secteurs du modèle."""
        return self.config["sectors"].copy()

    @property
    def pillars(self):
        """Retourne les piliers du modèle."""
        return list(self.config["base_weights"])

    @property
    def benchmark_name(self):
        """Retourne le nom du benchmark du marché."""
        return self.config["volatility"]["benchmark_name"]

# 2. Pipeline de calcul
def run_model(sector_workbook, macro_workbook, config):
    """Exécute les calculs communs du marché."""
    raw_variables = {}
    pillars, variable_scores = compute_historical_pillars(
        sector_workbook,
        config,
        details=raw_variables,
    )
    # Exemple EU au 30/09/2026 : extraits de DataFrame, valeurs arrondies.
    # pillars = {"Leverage": DataFrame, "Margin": DataFrame,
    #            "Value": DataFrame, "Growth": DataFrame}
    # pillars["Leverage"] :
    #              Tech   Fin   Oil
    # 2026-09-30   4.32  6.28  4.15
    #
    # variable_scores = {"net_debt_ebitda": DataFrame,
    #                    "fcf_total_debt": DataFrame, ...}
    # variable_scores["net_debt_ebitda"] :
    #              Tech   Fin   Oil
    # 2026-09-30   5.08  7.05  5.25
    #
    # raw_variables["net_debt_ebitda"] : valeurs avant notation.
    #              Tech    Fin   Oil
    # 2026-09-30   0.44  10.12  0.75
    # Les dates sont en index et les secteurs en colonnes.
    # Momentum et Volatility sont ajoutés ensuite, plus bas.

    momentum, momentum_scores = compute_momentum(sector_workbook, config, details=raw_variables)

    # Exemple EU au 30/09/2026 :
    # momentum est un DataFrame : dates en index, secteurs en colonnes.
    #              Tech     Fin     Oil
    # 2026-09-30  7.6923  8.2051  7.1795
    #
    # momentum_scores contient trois DataFrame (dates x secteurs) :
    # {"momentum_6m_1m": DataFrame, "momentum_12m_1m": DataFrame,
    #  "earnings_revision_ratio": DataFrame}
    # momentum_scores["momentum_6m_1m"] :
    #              Tech     Fin     Oil
    # 2026-09-30 10.0000  9.2308  1.5385
    # Pour Tech, les trois scores sont 10.0000, 7.6923 et 5.3846 ;
    # leur moyenne donne le score momentum 7.6923.
    # details ajoute aussi les trois signaux bruts dans raw_variables :
    # pour Tech, le signal brut 6 mois vaut ici 0.3202 avant classement.
    # Les scores EU sont sur 10 ; les scores US restent des rangs bruts.

    volatility, volatility_scores, returns = compute_volatility(
        sector_workbook,
        config,
        details=raw_variables,
    )
    benchmark_returns = read_benchmark_returns(
        sector_workbook[config["volatility"]["sheet_returns"]] if sector_workbook is not None else None,
        volatility_config=config["volatility"],
        key_prefix=config["history_prefix"],
        sheet_name=config["volatility"]["sheet_returns"],
    )

    pillars["Momentum"] = momentum
    pillars["Volatility"] = volatility

    variable_scores.update(momentum_scores)
    variable_scores.update(volatility_scores)

    macro_context = build_macro_context(sector_workbook, macro_workbook, config)
    # macro_context est un DataFrame : dates en index, une ligne par mois.
    # Exemple EU, extrait historique (macro_score arrondi) :
    #             macro_score  cycle  signal_taux
    # 2026-07-31      -0.1255     R       On
    # 2026-08-31       0.2236     R       Off
    # 2026-09-30       0.5733     E       On
    # macro_score et cycle viennent des données macro ; signal_taux du taux US 10 ans.
    # R = Recovery, E = Expansion ; certaines valeurs peuvent manquer selon les dates.

    (history, pillars, ranks, macro_context) = compute_final_results(pillars, macro_context, config)

    return {
        "history": history,
        "pillars": pillars,
        "ranks": ranks,
        "variable_scores": variable_scores,
        "variable_raw": raw_variables,
        "benchmark_weights": read_sector_block(
            sector_workbook["Bench"] if sector_workbook is not None else None,
            "B",
            sectors=config["sectors"],
            key_prefix=config["history_prefix"],
            start_row=2,
            sheet_name="Bench",
        ),
        "returns": returns,
        "benchmark_returns": benchmark_returns,
        "macro_context": macro_context,
    }
    # Structure du résultat : exemples EU du 30/09/2026, valeurs arrondies.
    # Les lignes et colonnes montrées ci-dessous ne sont qu’un extrait.
    #
    # results["history"] : DataFrame avec une ligne par mois et secteur.
    # date        secteur      score_global_macro  rang_global  recommendation
    # 2026-09-30  Materials    6.714286            7.0          Neutral
    # 2026-09-30  ConsStaples  3.714286            2.0          Negative
    #
    # results["pillars"] : dictionnaire de DataFrame, un tableau par pilier.
    # Exemple : results["pillars"]["Leverage"]
    #              Materials  ConsStaples  Fin  ...
    # 2026-09-30      ...          ...      ...
    # 2026-08-31      ...          ...      ...
    # Les dates sont en index et les secteurs en colonnes.
    # results["ranks"]["Leverage"] a la même forme, avec les rangs du pilier.
    #
    # results["variable_scores"]["net_debt_ebitda"] : scores de la variable.
    # results["variable_raw"]["net_debt_ebitda"] : valeurs avant notation.
    # Ces deux dictionnaires contiennent aussi des DataFrame dates x secteurs.
    #
    # results["benchmark_weights"] : DataFrame des poids Bench, dates x secteurs.
    # results["returns"] : DataFrame des rendements sectoriels, dates x secteurs.
    # results["benchmark_returns"] : DataFrame avec une colonne "benchmark".
    #                benchmark
    # 2026-09-30     -0.025180
    #
    # results["macro_context"] : DataFrame du score, du cycle et du taux.
    #              macro_score  cycle  signal_taux
    # 2026-09-30      0.573336   E      On


# 3. Sous-scores et piliers
def excel_rank(value, values, order):
    """Reproduit RANK Excel. ordre=1 croissant, ordre=0 décroissant."""
    if pd.isna(value):
        return np.nan

    series = pd.Series(values, dtype="float64").dropna()
    if series.empty:
        return np.nan

    if order == 1:
        return 1 + int((series < value).sum())

    return 1 + int((series > value).sum())


def cross_sectional_rank(row_number, order, scale_0_10=True):
    """Classe les secteurs au même mois."""
    result = pd.Series(index=row_number.index, dtype=float)
    n_sectors = len(row_number.index)

    for sector in row_number.index:
        value = row_number[sector]

        if pd.isna(value):
            result[sector] = np.nan
            continue

        rank = excel_rank(value, row_number.values, order)

        if scale_0_10:
            result[sector] = rank / n_sectors * 10
        else:
            result[sector] = rank

    return result


def rank_sectors(row_number):
    """Classe les secteurs de 1 = Worst à N = Best, avec le rang minimal des ex æquo."""
    return cross_sectional_rank(row_number, order=1, scale_0_10=False)


def percentrank_inc(values, x):
    """Reproduit PERCENTRANK.INC d'Excel."""
    values = pd.Series(values, dtype=float).dropna().sort_values().to_numpy()

    if len(values) == 0 or pd.isna(x):
        return np.nan
    if len(values) == 1:
        return 1.0
    if x <= values[0]:
        return 0.0
    if x >= values[-1]:
        return 1.0

    positions = np.where(np.isclose(values, x, rtol=0, atol=1e-12))[0]

    if len(positions):
        return positions[0] / (len(values) - 1)

    right = np.searchsorted(values, x, side="right")
    left = right - 1
    x0 = values[left]
    x1 = values[right]
    fraction = (x - x0) / (x1 - x0)

    return (left + fraction) / (len(values) - 1)


def compute_deviation_from_mean(raw, exclude_financials_from_mean=False):
    """Calcule l'écart du secteur à la moyenne du même mois."""
    diff = pd.DataFrame(index=raw.index, columns=raw.columns, dtype=float)

    for date, row_number in raw.iterrows():
        base = row_number.copy()

        if exclude_financials_from_mean:
            base = base.drop("Fin")

        mean = base.mean(skipna=True)
        diff.loc[date] = row_number - mean

    return diff


def compute_historical_score(diff, rank_order, history_window, sector_windows=None):
    """Note chaque écart dans sa fenêtre ; les dates vont du plus récent au plus ancien."""
    sector_windows = sector_windows or {}
    score = pd.DataFrame(index=diff.index, columns=diff.columns, dtype=float)

    for i in range(len(diff)):
        for j, sector in enumerate(diff.columns):
            value = diff.iloc[i][sector]
            if pd.isna(value):
                continue

            # La fenêtre inclut le mois courant : 60 mois donne 61 observations.
            n_observations = sector_windows.get(sector, history_window) + 1
            end = i + n_observations
            if end > len(diff):
                continue

            window = diff.iloc[i:end][sector]
            rank = excel_rank(value, window.values, rank_order)
            score.iloc[i, j] = 10 * rank / n_observations

    return score


def compute_variable_score(raw, variable_config, history_window):
    """Calcule le score d'une sous-variable."""
    diff = compute_deviation_from_mean(raw, variable_config["exclude_financials_from_mean"])
    # Exemple avec trois secteurs (Fin incluse dans la moyenne) :
    # raw              Tech   Fin   Oil   Moyenne
    # 2026-09-30        18     12     9      13
    # 2026-08-31        12     10     8      10
    # diff             Tech   Fin   Oil
    # 2026-09-30         5     -1    -4
    # 2026-08-31         2      0    -2
    # Dates en index, secteurs en colonnes ; chaque valeur est un écart.

    score_hist = compute_historical_score(
        diff,
        variable_config["rank_order"],
        history_window,
        variable_config.get("sector_windows"),
    )
    # score_hist : note de chaque secteur comparé à son propre historique.
    # Exemple fictif : history_window=2, rank_order=1 (valeur élevée favorable).
    # diff             Tech   Fin   Oil
    # 2026-09-30         5     -1    -4
    # 2026-08-31         2      0    -2
    # 2026-07-31         1     -2     1
    # En septembre, Tech est 3e sur [5, 2, 1] : 10 x 3 / 3 = 10.
    # score_hist (arrondi) :
    #                 Tech   Fin   Oil
    # 2026-09-30    10.00  6.67  3.33
    # 2026-08-31      NaN   NaN   NaN
    # 2026-07-31      NaN   NaN   NaN
    # Les deux dates anciennes n’ont pas assez d’historique dans cet exemple.
    # Par défaut : 60 observations antérieures + mois courant = 61.

    if not variable_config["mix_cross_sectional_rank"]:
        return score_hist

    score = score_hist.copy()

    for date in score.index:
        score_cross = cross_sectional_rank(
            raw.loc[date],
            variable_config["rank_order"],
            scale_0_10=True,
        )
        # Exemple pour un même mois, rank_order=1 :
        # raw.loc[date] : Tech=18, Fin=12, Oil=9.
        # Rangs parmi les trois secteurs : Tech=3, Fin=2, Oil=1.
        # score_cross = 10 x rang / 3 (Series indexée par secteur) :
        # Tech    10.00
        # Fin      6.67
        # Oil      3.33
        # La note de la variable est ensuite la moyenne de score_hist et score_cross.
        score.loc[date] = (score_hist.loc[date] + score_cross) / 2

    return score


def compose_pillar(scores, config, pillar):
    """Agrège les sous-scores selon la composition du pilier."""
    # scores contient un DataFrame par variable : dates en index, secteurs en colonnes.
    sectors = config["sectors"]
    composition = config["pillar_composition"][pillar]

    # Garder les dates communes à tous les tableaux de scores.
    dates = None
    for score in scores.values():
        dates = score.index if dates is None else dates.intersection(score.index)

    if dates is None or len(dates) == 0:
        raise ValueError(f"Aucune date disponible pour le pilier {pillar}.")

    # Le tableau final contiendra un score par date et par secteur.
    result = pd.DataFrame(index=dates, columns=sectors, dtype=float)

    # Certaines compositions sont propres à un secteur ; sinon utiliser default.
    by_sector = composition.get("by_sector", {})

    for date in dates:
        for sector in sectors:
            # Exemple EU / Growth : Fin utilise eps_growth et sales_growth.
            # Si leurs scores sont 6 et 8, le score du pilier vaut (6 + 8) / 2 = 7.
            variables = by_sector.get(sector, composition["default"])
            values = [scores[variable].at[date, sector] for variable in variables]

            # Une valeur manquante rend le pilier indisponible ; sinon calculer la moyenne.
            if any(pd.isna(value) for value in values):
                result.at[date, sector] = np.nan
            else:
                result.at[date, sector] = sum(values) / len(values)

    # Arrondir seulement si une précision est configurée pour ce pilier.
    decimals = config["pillar_rounding"].get(pillar)

    if decimals is not None:
        result = result.round(decimals)

    return result


def compute_historical_pillars(wb, config, details=None):
    """Calcule Leverage, Margin, Value et Growth."""
    sectors = config["sectors"]
    variable_scores = {}
    pillars = {}

    for pillar, variables in config["historical_variables"].items():
        pillar_scores = {}

        # Exemple EU, pilier Leverage :
        # variable_name = "net_debt_ebitda"
        # cfg = {"sheet": "Leverage_FMA", "column": "AF", "rank_order": 0,
        #        "exclude_financials_from_mean": False,
        #        "mix_cross_sectional_rank": False}
        for variable_name, cfg in variables.items():
            ws = wb[cfg["sheet"]] if wb is not None else None
            raw = read_sector_block(
                ws,
                cfg["column"],
                sectors=sectors,
                key_prefix=config["history_prefix"],
                sheet_name=cfg["sheet"],
            )

            score = compute_variable_score(raw, cfg, config["history_window"])

            if details is not None:
                details[variable_name] = raw

            variable_scores[variable_name] = score
            pillar_scores[variable_name] = score

        pillars[pillar] = compose_pillar(pillar_scores, config, pillar)

    return pillars, variable_scores


def compute_momentum(wb, config, details=None):
    """Calcule les deux signaux de prix et la balance des révisions bénéficiaires."""
    cfg = config["momentum"]
    sectors = config["sectors"]
    sheet = wb[cfg["sheet"]] if wb is not None else None

    def read_data(column):
        return read_sector_block(
            sheet, column, sectors=sectors, key_prefix=config["history_prefix"],
            sheet_name=cfg["sheet"],
        )

    prices = read_data(cfg["price_column"])
    revisions_up = read_data(cfg["revision_up_column"])
    revisions_down = read_data(cfg["revision_down_column"])
    revisions_unchanged = read_data(cfg["revision_unchanged_column"])
    short_returns = pd.DataFrame(index=prices.index, columns=sectors, dtype=float)
    long_returns = pd.DataFrame(index=prices.index, columns=sectors, dtype=float)
    current_month_sector = cfg.get("current_month_sector")

    # Les séries sont décroissantes ; i + horizon désigne un mois antérieur.
    for i, date in enumerate(prices.index):
        for horizon, returns in (
            (cfg["short_horizon"], short_returns),
            (cfg["long_horizon"], long_returns),
        ):
            if i + horizon >= len(prices):
                continue
            # Exemple en septembre, horizon=6 et mois consécutifs :
            # cours août 120 / cours mars 100 - 1 = 0.20, soit 20 %.
            returns.iloc[i] = prices.iloc[i + 1] / prices.iloc[i + horizon] - 1
            # Tech (configuration actuelle) utilise septembre : 132 / mars 100 - 1 = 0.32.
            if current_month_sector:
                returns.at[date, current_month_sector] = (
                    prices.iloc[i][current_month_sector]
                    / prices.iloc[i + horizon][current_month_sector] - 1
                )

    # Exemple : 7 hausses, 2 baisses, 1 inchangée -> (7 - 2) / 10 = 0.50.
    revision_ratio = (revisions_up - revisions_down) / (
        revisions_up + revisions_down + revisions_unchanged
    )
    raw_signals = {
        "momentum_6m_1m": short_returns,
        "momentum_12m_1m": long_returns,
        "earnings_revision_ratio": revision_ratio,
    }
    # details conserve les trois signaux avant leur classement.
    if details is not None:
        details.update(raw_signals)

    # Exemple fictif à une date : Tech=0.50, Fin=0.20, Oil=-0.10.
    # Rangs : 3, 2, 1 ; EU sur 10 : 10, 6.67, 3.33 ; US garde 3, 2, 1.
    scale_scores = cfg.get("scale_0_10", True)
    variable_scores = {}
    for name, values in raw_signals.items():
        variable_scores[name] = values.apply(
            lambda row: cross_sectional_rank(row, order=1, scale_0_10=scale_scores),
            axis=1,
        )
    # Exemple EU : scores 10, 6.67 et 3.33 -> moyenne Momentum = 6.67.
    # Une composante manquante rend le pilier indisponible.
    momentum = compose_pillar(variable_scores, config, "Momentum")
    return momentum, variable_scores


def compute_volatility(wb, config, details=None):
    """Calcule total volatility et downside volatility."""
    cfg = config["volatility"]
    sectors = config["sectors"]

    returns = read_sector_returns(
        wb[cfg["sheet_returns"]] if wb is not None else None,
        volatility_config=cfg,
        sectors=sectors,
        key_prefix=config["history_prefix"],
        sheet_name=cfg["sheet_returns"],
    )

    total_volatility = pd.DataFrame(index=returns.index, columns=sectors, dtype=float)
    downside_volatility = pd.DataFrame(index=returns.index, columns=sectors, dtype=float)

    n_obs_vol = cfg["volatility_offset"] + 1
    n_obs_downside = cfg["offset_downside"] + 1

    for i in range(len(returns)):
        if i + n_obs_vol <= len(returns):
            window = returns.iloc[i : i + n_obs_vol]
            total_volatility.iloc[i] = window.std(ddof=1) * math.sqrt(12)

        if i + n_obs_downside <= len(returns):
            window = returns.iloc[i : i + n_obs_downside]

            for sector in sectors:
                negative_returns = window[sector][window[sector] < 0].dropna()

                if len(negative_returns) >= 2:
                    downside_volatility.at[
                        returns.index[i], sector,
                    ] = negative_returns.std(ddof=1) * math.sqrt(12)

    if details is not None:
        details.update(
            {
                "volatility_6m": total_volatility,
                "downside_volatility_18m": downside_volatility,
            }
        )

    score_config = {
        "rank_order": 0,
        "exclude_financials_from_mean": False,
        "mix_cross_sectional_rank": False,
    }

    score_total = compute_variable_score(total_volatility, score_config, config["history_window"])
    score_downside = compute_variable_score(downside_volatility, score_config, config["history_window"])

    variable_scores = {
        "volatility_6m": score_total,
        "downside_volatility_18m": score_downside,
    }

    volatility = compose_pillar(variable_scores, config, "Volatility")

    return volatility, variable_scores, returns


# 4. Contexte macro et taux


def compute_rate_signal(us10y_rates, signal_config):
    """Calcule le rate overlay à partir du US 10Y."""
    alpha = signal_config["alpha_ewma"]
    threshold = signal_config["percentile_threshold"]

    # Exemple fictif : taux 2, 3, 3 ; alpha=0.715 et seuil=0.85.
    # Les dates sont traitées de la plus ancienne à la plus récente.
    df = us10y_rates.copy().sort_index()
    df["ewma_us10y"] = np.nan
    df["diff_ewma"] = np.nan
    df["percentile_taux"] = np.nan
    df["signal_taux"] = None

    for i in range(len(df)):
        rate = df.iloc[i]["us10y"]

        if pd.isna(rate):
            continue

        if i == 0 or pd.isna(df.iloc[i - 1]["ewma_us10y"]):
            # Premier mois : l’EWMA prend directement le taux, ici 2.
            ewma = rate
        else:
            previous_ewma = df.iloc[i - 1]["ewma_us10y"]
            # Deuxième mois : 0.715 x 2 + 0.285 x 3 = 2.285.
            ewma = alpha * previous_ewma + (1 - alpha) * rate

        df.iloc[i, df.columns.get_loc("ewma_us10y")] = ewma

        # Deuxième mois : taux 3 - EWMA 2.285 = écart 0.715.
        diff = rate - ewma
        df.iloc[i, df.columns.get_loc("diff_ewma")] = diff

        # Troisième mois : écarts [0, 0.511225, 0.715] classés.
        # L’écart courant 0.511225 est au milieu : percentile 0.50.
        history = df["diff_ewma"].iloc[: i + 1]
        percentile = percentrank_inc(history, diff)

        df.iloc[i, df.columns.get_loc("percentile_taux")] = percentile
        # Avec un seuil de 0.85 : 0.50 donne Off, 1.00 donne On.
        # Le seuil doit être strictement dépassé : 0.85 donne Off.
        df.iloc[i, df.columns.get_loc("signal_taux")] = (
            "On" if percentile > threshold else "Off"
        )

    # Résultat de l’exemple :
    # date        us10y  ewma_us10y  diff_ewma  percentile_taux  signal_taux
    # 2026-01-31    2.0    2.000000   0.000000             1.00       On
    # 2026-02-28    3.0    2.285000   0.715000             1.00       On
    # 2026-03-31    3.0    2.488775   0.511225             0.50       Off
    # Avec une seule observation, percentrank_inc renvoie 1.00.
    return df


def build_macro_context(sector_workbook, macro_workbook, config):
    """Assemble macro score, régime et rate overlay."""
    macro = read_macro_data(
        macro_workbook,
        macro_config=config["macro"],
        regime_weights=config["regime_weights"],
        key_prefix=config["history_prefix"],
    )

    rate = read_us10y(
        sector_workbook,
        signal_config=config["rate_signal"],
        key_prefix=config["history_prefix"],
    )

    rate_signal = compute_rate_signal(rate, config["rate_signal"])

    context = macro.join(rate_signal[["signal_taux"]], how="outer")

    return context.sort_index()


# 5. Rangs, votes et recommandations


def align_pillars(pillars):
    """Aligne les piliers sur une clé mensuelle commune."""
    normalized_pillars = {
        name: normalize_monthly_index(df) for name, df in pillars.items()
    }

    dates = None

    for df in normalized_pillars.values():
        dates = df.index if dates is None else dates.intersection(df.index)

    if dates is None or len(dates) == 0:
        raise ValueError("Aucune date mensuelle commune entre les piliers.")

    dates = dates.sort_values(ascending=False)

    return {name: df.loc[dates].copy() for name, df in normalized_pillars.items()}


def compute_pillar_ranks(pillars):
    """Classe les secteurs dans chaque pilier."""
    ranks = {}

    for name, df in pillars.items():
        rank_frame = pd.DataFrame(index=df.index, columns=df.columns, dtype=float)

        for date in df.index:
            rank_frame.loc[date] = rank_sectors(df.loc[date])

        ranks[name] = rank_frame

    return ranks


def compute_weighted_score(ranks, weights, date):
    """Agrège les rangs des piliers avec les poids fournis."""
    sectors = next(iter(ranks.values())).columns

    result = pd.Series(0.0, index=sectors)
    valid = pd.Series(True, index=sectors)

    for pillar, pillar_weight in weights.items():
        if pillar_weight == 0:
            continue
        series = ranks[pillar].loc[date]
        valid &= series.notna()
        result += series.fillna(0) * pillar_weight

    result[~valid] = np.nan

    return result


def select_top_worst(
    top_count,
    bottom_count,
    overall_rank,
    n_top,
    n_worst,
):
    """Sélectionne les secteurs Positive et Negative."""
    table = pd.DataFrame(
        {
            "top_count": top_count,
            "bottom_count": bottom_count,
            "rang_global": overall_rank,
        }
    )

    top_candidates = table[table["bottom_count"] < n_worst].copy()

    top_candidates = top_candidates.sort_values(
        ["top_count", "rang_global", "bottom_count"],
        ascending=[False, False, True],
    )

    top = list(top_candidates.head(n_top).index)

    worst_candidates = table[table["top_count"] < n_top].copy()

    worst_candidates = worst_candidates.sort_values(
        ["bottom_count", "rang_global", "top_count"],
        ascending=[False, True, True],
    )

    worst = list(worst_candidates.head(n_worst).index)

    return top, worst


def compute_macro_adjusted_score(ranks, tilt_rank, date, config):
    """Agrège les rangs avec la pondération macro du marché."""
    sectors = config["sectors"]
    weights = config["macro_adjusted_weights"]

    score = pd.Series(0.0, index=sectors)
    valid = pd.Series(True, index=sectors)

    for pillar, pillar_weight in weights.items():
        if pillar == "Macro" or pillar_weight == 0:
            continue

        series = ranks[pillar].loc[date]
        valid &= series.notna()
        score += series.fillna(0) * pillar_weight

    macro_weight = weights.get("Macro", 0)

    if macro_weight:
        valid &= tilt_rank.notna()
        score += tilt_rank.fillna(0) * macro_weight

    score[~valid] = np.nan

    return score


def _count_sector_votes(ranks, tilt_rank, date, rate_signal, config):
    """Applique les votes des piliers, du régime et du signal de taux."""
    sectors = config["sectors"]
    n_sectors = len(sectors)
    n_top, n_worst = config["n_top"], config["n_worst"]
    # Deux Series indexées par secteur ; au départ chaque compteur vaut 0.
    # Exemple EU : 13 secteurs, Top = rangs 11 à 13 ; Worst = rangs 1 à 3.
    top_count = pd.Series(0, index=sectors, dtype=int)
    bottom_count = pd.Series(0, index=sectors, dtype=int)

    # Exemple EU Tech : Margin=12, Growth=13, Momentum=12 -> 3 votes Top.
    for pillar in config["top_vote_pillars"]:
        r = ranks[pillar].loc[date]
        top_count += (r > n_sectors - n_top).fillna(False).astype(int)

    # Exemple fictif EU : Fin avec un rang Value=2 reçoit 1 vote Worst.
    for pillar in config["bottom_vote_pillars"]:
        r = ranks[pillar].loc[date]
        bottom_count += (r <= n_worst).fillna(False).astype(int)

    # Le rang du tilt macro vote aussi : 2 votes en EU, 1 en US.
    # Exemple EU : rang Tech=12 -> 2 votes Top de plus.
    macro_weight = config["macro_vote_weight"]

    top_count += macro_weight * (tilt_rank > n_sectors - n_top).fillna(False).astype(int)

    bottom_count += macro_weight * (tilt_rank <= n_worst).fillna(False).astype(int)

    # Avec On, Leverage et Value apportent un vote supplémentaire si leur rang le permet.
    # Exemple fictif EU : Fin, Leverage=12 et Value=2 -> +1 Top et +1 Worst.
    # Avec Off, aucun vote de taux n’est ajouté.
    if rate_signal == "On":
        for pillar in config["rate_overlay_pillars"]:
            r = ranks[pillar].loc[date]

            top_count += (r > n_sectors - n_top).fillna(False).astype(int)

            bottom_count += (r <= n_worst).fillna(False).astype(int)

    # EU et US : Travel & leisure ne reçoit aucun vote à partir du 31/10/2011.
    for rule in config.get("no_vote_rules", []):
        start_date = (
            pd.Timestamp(rule["start_date"]).to_period("M").to_timestamp("M")
        )

        end_date = rule.get("end_date")
        enabled = date >= start_date

        if end_date:
            end_date = pd.Timestamp(end_date).to_period("M").to_timestamp("M")
            enabled = enabled and date <= end_date

        if enabled:
            sector = rule["sector"]
            top_count[sector] = 0
            bottom_count[sector] = 0
    # Exemple EU du 30/09/2026 : Tech = 5 Top et 0 Worst.
    # Les recommandations Positive/Negative sont choisies après ce retour.
    return top_count, bottom_count


def compute_final_results(pillars, macro_context, config):
    """Construit scores finaux, votes et recommandations."""
    # Exemple EU du 30/09/2026 : 13 secteurs, jusqu’à 3 Positive et 3 Negative.
    sectors = config["sectors"]
    n_sectors = len(sectors)
    n_top = config["n_top"]
    n_worst = config["n_worst"]

    # Aligner les dates puis classer les secteurs dans chaque pilier.
    # Rangs de Tech : Leverage=6, Margin=12, Value=9, Growth=13,
    # Momentum=12, Volatility=4. Un rang plus élevé est plus favorable.
    pillars = align_pillars(pillars)
    ranks = compute_pillar_ranks(pillars)
    rows = []

    for date in pillars["Leverage"].index:
        if date not in macro_context.index:
            continue

        # Dans cet exemple : cycle E, macro_score=0.5733 et signal_taux=On.
        # Une date sans cycle ou signal de taux valide est ignorée.
        regime = macro_context.at[date, "cycle"]
        macro_score = macro_context.at[date, "macro_score"]
        rate_signal = macro_context.at[date, "signal_taux"]

        if regime not in config["regime_weights"]:
            continue
        if rate_signal not in {"On", "Off"}:
            continue

        # Score de base EU : chaque rang compte pour 1/6.
        # Tech : (6 + 12 + 9 + 13 + 12 + 4) / 6 = 9.3333 ; rang_base=12.
        base_score = compute_weighted_score(ranks, config["base_weights"], date)
        base_rank = rank_sectors(base_score)

        # Le régime choisit les poids du tilt. En E pour EU, ils valent aussi 1/6 :
        # Tech obtient score_tilt=9.3333 et rang_tilt_macro=12.
        tilt_score = compute_weighted_score(ranks, config["regime_weights"][regime], date)
        tilt_rank = rank_sectors(tilt_score)

        # Score global EU : six rangs et rang_tilt_macro comptent chacun pour 1/7.
        # Tech : (6 + 12 + 9 + 13 + 12 + 4 + 12) / 7 = 9.7143 ; rang_global=12.
        # La composante Macro est le rang du tilt, pas le macro_score brut.
        macro_adjusted_score = compute_macro_adjusted_score(ranks, tilt_rank, date, config)
        overall_rank = rank_sectors(macro_adjusted_score)

        # EU : rangs 11 à 13 -> vote Top ; rangs 1 à 3 -> vote Worst.
        # Tech : 3 votes Top des piliers + 2 du tilt = 5 ; aucun vote Worst.
        # Le taux On examine aussi Leverage et Value : aucun vote supplémentaire ici.
        top_count, bottom_count = _count_sector_votes(
            ranks, tilt_rank, date, rate_signal, config
        )

        # La sélection utilise les votes, puis le rang global pour départager.
        # Dans cet exemple, Tech est retenu dans Positive.
        top, worst = select_top_worst(top_count, bottom_count, overall_rank, n_top, n_worst)

        for sector in sectors:
            if sector in top:
                recommendations = "Positive"
            elif sector in worst:
                recommendations = "Negative"
            else:
                recommendations = "Neutral"

            # Exemple de ligne : 2026-09-30, Tech, E, score global 9.7143,
            # rang 12, votes Top/Worst 5/0, recommandation Positive.
            record = {
                "date": date,
                "secteur": sector,
                "cycle": regime,
                "macro_score": macro_score,
                "signal_taux": rate_signal,
                "score_base": base_score[sector],
                "rang_base": base_rank[sector],
                "score_tilt": tilt_score[sector],
                "rang_tilt_macro": tilt_rank[sector],
                "score_global_macro": macro_adjusted_score[sector],
                "rang_global": overall_rank[sector],
                "top_count": int(top_count[sector]),
                "bottom_count": int(bottom_count[sector]),
                "recommendation": recommendations,
            }

            for pillar in ranks:
                record[f"score_{pillar.lower()}"] = pillars[pillar].at[date, sector]
                record[f"rang_{pillar.lower()}"] = ranks[pillar].at[date, sector]

            rows.append(record)

    # Retour : historique (une ligne par date et secteur), piliers, rangs et contexte.
    return (pd.DataFrame(rows), pillars, ranks, macro_context)


# 6. Exports des scores


def save_production(results, output_dir, config):
    """Sauvegarde uniquement les sorties courantes du modèle."""
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)

    prefix = config["output_prefix"]
    history = results["history"]
    pillars = results["pillars"]
    ranks = results["ranks"]

    pillar_dates = None

    for df in pillars.values():
        pillar_dates = (
            df.index if pillar_dates is None else pillar_dates.intersection(df.index)
        )

    if pillar_dates is not None and len(pillar_dates):
        pillar_date = pillar_dates.max()
        rows = []

        for sector in config["sectors"]:
            record = {
                "date": pillar_date,
                "secteur": sector,
            }

            for pillar in config["base_weights"]:
                record[f"score_{pillar.lower()}"] = pillars[pillar].at[
                    pillar_date, sector
                ]
                record[f"rang_{pillar.lower()}"] = ranks[pillar].at[
                    pillar_date, sector
                ]

            rows.append(record)

        save_table(pd.DataFrame(rows), directory / f"{prefix}_piliers_latest_available.xlsx")

    if history.empty:
        return

    latest_date = history["date"].max()
    latest = history[history["date"] == latest_date].copy()
    latest = latest.sort_values("rang_global", ascending=False)

    save_table(latest, directory / f"{prefix}_recommandations_latest.xlsx")


def save_all_outputs(results, output_dir, config):
    """Sauvegarde l'historique complet et les sorties courantes."""
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)

    prefix = config["output_prefix"]

    save_table(results["history"], directory / f"{prefix}_historique_modele.xlsx")

    save_production(results, directory, config)


def print_latest(results):
    """Affiche la dernière recommandation disponible."""
    history = results["history"]

    if history.empty:
        print("Aucun résultat calculé.")
        return

    latest_date = history["date"].max()
    latest = history[history["date"] == latest_date].copy()
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

    print(f"\nDate de la recommandation finale : " f"{latest_date.date()}")
    print(latest[cols].to_string(index=False))

    pillar_dates = None
    for df in results["pillars"].values():
        pillar_dates = (
            df.index if pillar_dates is None else pillar_dates.intersection(df.index)
        )

    if pillar_dates is not None and len(pillar_dates):
        pillar_date = pillar_dates.max()

        if pillar_date > latest_date:
            print(
                f"\nAttention : les piliers sont disponibles jusqu'au "
                f"{pillar_date.date()}, mais le régime macro permet une "
                f"recommandation finale seulement jusqu'au {latest_date.date()}."
            )
