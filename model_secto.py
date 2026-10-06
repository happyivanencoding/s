# -*- coding: utf-8 -*-
"""Interface unique pour les modèles sectoriels EU et US."""

from copy import deepcopy
from pathlib import Path

import pandas as pd

from data_io import (
    FICHIER_EXCEL_EU_PAR_DEFAUT,
    FICHIER_EXCEL_US_PAR_DEFAUT,
    FICHIER_MACRO_PAR_DEFAUT,
    ouvrir_workbooks,
)
from local_config import MODELE_EU, MODELE_US
import sector_core as core


UNIVERSES = {
    "EU": {
        "config": MODELE_EU,
        "excel_file": FICHIER_EXCEL_EU_PAR_DEFAUT,
    },
    "US": {
        "config": MODELE_US,
        "excel_file": FICHIER_EXCEL_US_PAR_DEFAUT,
    },
}


class SectorModel:
    """Configure, exécute et inspecte un modèle sectoriel."""

    def __init__(
        self,
        universe="EU",
        variables=None,
        excel_file=None,
        macro_file=None,
    ):
        universe = universe.upper()

        if universe not in UNIVERSES:
            raise ValueError("Universe must be EU or US.")

        settings = UNIVERSES[universe]

        self.universe = universe
        self._base_config = deepcopy(settings["config"])
        self.config = deepcopy(self._base_config)
        self.default_excel_file = settings["excel_file"]
        self.default_macro_file = FICHIER_MACRO_PAR_DEFAUT
        self.excel_file = excel_file
        self.macro_file = macro_file
        self.results = None

        self._available_variables = deepcopy(
            self.config["variables_backtest"]
        )
        self._variable_selection = {
            pillar: values.copy()
            for pillar, values in self._available_variables.items()
        }

        if variables:
            self._apply_variable_selection(variables)

    @property
    def sectors(self):
        """Retourne les secteurs du modèle."""
        return self.config["secteurs"].copy()

    @property
    def pillars(self):
        """Retourne les piliers du modèle."""
        return list(self.config["poids_base"])

    @property
    def benchmark_name(self):
        """Retourne le nom du benchmark du marché."""
        return self.config["volatilite"]["benchmark_nom"]

    def available_variables(self, pillar=None):
        """Retourne les variables disponibles avant sélection."""
        if pillar is None:
            return deepcopy(self._available_variables)

        self._validate_pillar(pillar)
        return self._available_variables[pillar].copy()

    def variables(self, pillar=None):
        """Retourne les variables actuellement utilisées."""
        if pillar is None:
            return deepcopy(self._variable_selection)

        self._validate_pillar(pillar)
        return self._variable_selection[pillar].copy()

    def variable_definitions(self, pillar=None):
        """Présente la source et les règles des variables disponibles."""
        pillars = self.pillars if pillar is None else [pillar]

        if pillar is not None:
            self._validate_pillar(pillar)

        rows = []

        for pillar_name in pillars:
            active = set(
                self._variable_selection[pillar_name]
            )
            historical = self._base_config[
                "variables_historiques"
            ].get(pillar_name, {})

            for variable in self._available_variables[pillar_name]:
                if variable in historical:
                    definition = historical[variable]
                    window = self._base_config[
                        "fenetre_historique"
                    ]
                    sector_windows = definition.get(
                        "fenetre_par_secteur",
                        {},
                    )

                    if sector_windows:
                        overrides = ", ".join(
                            f"{sector}={months}"
                            for sector, months
                            in sector_windows.items()
                        )
                        window = f"{window}; {overrides}"

                    rows.append(
                        {
                            "universe": self.universe,
                            "pillar": pillar_name,
                            "variable": variable,
                            "active": variable in active,
                            "source": "historical",
                            "sheet": definition["sheet"],
                            "column": definition["colonne"],
                            "direction": (
                                "higher"
                                if definition["ordre_rank"] == 1
                                else "lower"
                            ),
                            "window": window,
                            "cross_sectional_mix": definition[
                                "mix_rank_transversal"
                            ],
                        }
                    )
                    continue

                if pillar_name == "Momentum":
                    cfg = self._base_config["momentum"]

                    if variable == "momentum_6m_1m":
                        column = cfg["colonne_prix"]
                        window = f"{cfg['horizon_court']}M-1M"
                        source = "price"
                    elif variable == "momentum_12m_1m":
                        column = cfg["colonne_prix"]
                        window = f"{cfg['horizon_long']}M-1M"
                        source = "price"
                    else:
                        column = " / ".join(
                            [
                                cfg["colonne_revision_up"],
                                cfg["colonne_revision_down"],
                                cfg["colonne_revision_unchanged"],
                            ]
                        )
                        window = None
                        source = "earnings revisions"

                    rows.append(
                        {
                            "universe": self.universe,
                            "pillar": pillar_name,
                            "variable": variable,
                            "active": variable in active,
                            "source": source,
                            "sheet": cfg["sheet"],
                            "column": column,
                            "direction": "higher",
                            "window": window,
                            "cross_sectional_mix": True,
                        }
                    )
                    continue

                cfg = self._base_config["volatilite"]
                window = (
                    cfg["offset_volatilite"] + 1
                    if variable == "volatility_6m"
                    else cfg["offset_downside"] + 1
                )

                rows.append(
                    {
                        "universe": self.universe,
                        "pillar": pillar_name,
                        "variable": variable,
                        "active": variable in active,
                        "source": "sector returns",
                        "sheet": cfg["sheet_returns"],
                        "column": "derived",
                        "direction": "lower",
                        "window": f"{window} observations",
                        "cross_sectional_mix": False,
                    }
                )

        return pd.DataFrame(rows)

    def composition(self, pillar):
        """Présente les variables réellement utilisées par secteur."""
        self._validate_pillar(pillar)

        composition = self.config[
            "composition_piliers"
        ][pillar]
        default = composition["default"]
        by_sector = composition.get(
            "par_secteur",
            {},
        )

        rows = []

        for sector in self.sectors:
            variables = by_sector.get(
                sector,
                default,
            )

            rows.append(
                {
                    "sector": sector,
                    "variables": " | ".join(variables),
                }
            )

        return pd.DataFrame(rows)

    def summary(self):
        """Présente la composition active du modèle."""
        rows = []

        for pillar in self.pillars:
            rows.append(
                {
                    "universe": self.universe,
                    "pillar": pillar,
                    "weight": self.config["poids_base"][pillar],
                    "variables": " | ".join(
                        self._variable_selection[pillar]
                    ),
                }
            )

        return pd.DataFrame(rows)

    def run(self, force=False):
        """Exécute le modèle et conserve les résultats en mémoire."""
        if self.results is not None and not force:
            return self.results

        wb_sector, wb_macro = ouvrir_workbooks(
            self.excel_file,
            self.macro_file,
            fichier_excel_defaut=self.default_excel_file,
        )

        try:
            self.results = core.executer_modele(
                wb_sector,
                wb_macro,
                self.config,
            )
        finally:
            wb_sector.close()
            wb_macro.close()

        return self.results

    def variable_values(self, variable, date=None):
        """Retourne les scores sectoriels d'une variable pour une date."""
        results = self._ensure_results()

        if variable not in results["variable_scores"]:
            raise ValueError(
                f"Unknown variable: {variable}"
            )

        score = results["variable_scores"][variable]
        selected_date = self._select_date(score.index, date)

        return score.loc[selected_date].rename(variable)

    def pillar_values(self, pillar, date=None):
        """Retourne les scores sectoriels d'un pilier pour une date."""
        self._validate_pillar(pillar)
        results = self._ensure_results()
        score = results["pillars"][pillar]
        selected_date = self._select_date(score.index, date)

        return score.loc[selected_date].rename(pillar)

    def variable_table(self, pillar, date=None):
        """Compare les variables actives et le score final d'un pilier."""
        self._validate_pillar(pillar)
        results = self._ensure_results()

        variables = self._variable_selection[pillar]
        pillar_score = results["pillars"][pillar]
        selected_date = self._select_date(
            pillar_score.index,
            date,
        )

        table = pd.DataFrame(index=self.sectors)

        for variable in variables:
            score = results["variable_scores"][variable]

            if selected_date in score.index:
                table[variable] = score.loc[selected_date]
            else:
                table[variable] = pd.NA

        table[pillar] = pillar_score.loc[selected_date]
        table.index.name = "sector"

        return table

    def recommendations(self, date=None):
        """Retourne les recommandations finales pour une date."""
        history = self._ensure_results()["history"]

        if history.empty:
            return history.copy()

        dates = pd.DatetimeIndex(
            history["date"].drop_duplicates()
        ).sort_values()
        selected_date = self._select_date(dates, date)

        return (
            history[
                history["date"] == selected_date
            ]
            .sort_values(
                "rang_global",
                ascending=False,
            )
            .reset_index(drop=True)
        )

    def save_outputs(self, output_dir="output"):
        """Sauvegarde les sorties du modèle."""
        results = self._ensure_results()

        core.sauvegarder_sorties(
            results,
            Path(output_dir),
            self.config,
        )

    def show_latest(self):
        """Affiche la dernière recommandation disponible."""
        core.afficher_latest(
            self._ensure_results()
        )

    def _apply_variable_selection(self, selections):
        """Applique les variables choisies aux compositions des piliers."""
        for pillar, variables in selections.items():
            self._validate_pillar(pillar)

            if variables is None:
                continue

            variables = list(dict.fromkeys(variables))

            if not variables:
                raise ValueError(
                    f"{pillar} must contain at least one variable."
                )

            available = self._available_variables[pillar]
            unknown = [
                variable
                for variable in variables
                if variable not in available
            ]

            if unknown:
                raise ValueError(
                    f"Unknown variables for {pillar}: {unknown}"
                )

            self._update_pillar_composition(
                pillar,
                variables,
            )
            self._variable_selection[pillar] = variables
            self.config["variables_backtest"][pillar] = (
                variables.copy()
            )

            if pillar in self.config["variables_historiques"]:
                definitions = self.config[
                    "variables_historiques"
                ][pillar]
                self.config["variables_historiques"][pillar] = {
                    variable: definitions[variable]
                    for variable in variables
                }

    def _update_pillar_composition(self, pillar, variables):
        """Filtre la composition existante avec la sélection explicite."""
        base_composition = self._base_config[
            "composition_piliers"
        ][pillar]
        base_default = base_composition["default"]
        base_by_sector = base_composition.get(
            "par_secteur",
            {},
        )

        default = [
            variable
            for variable in base_default
            if variable in variables
        ]

        if not default:
            raise ValueError(
                f"No selected variable is valid for the default "
                f"{pillar} composition."
            )

        by_sector = {}

        for sector, sector_variables in base_by_sector.items():
            selected = [
                variable
                for variable in sector_variables
                if variable in variables
            ]

            if not selected:
                raise ValueError(
                    f"No selected variable is valid for "
                    f"{pillar} / {sector}."
                )

            by_sector[sector] = selected

        self.config["composition_piliers"][pillar] = {
            "default": default,
        }

        if by_sector:
            self.config["composition_piliers"][pillar][
                "par_secteur"
            ] = by_sector

    def _validate_pillar(self, pillar):
        if pillar not in self._available_variables:
            raise ValueError(
                f"Unknown pillar: {pillar}"
            )

    def _ensure_results(self):
        if self.results is None:
            self.run()

        return self.results

    @staticmethod
    def _select_date(index, date):
        dates = pd.DatetimeIndex(index).sort_values()

        if len(dates) == 0:
            raise ValueError("No date available.")

        if date is None:
            return dates[-1]

        target = (
            pd.Timestamp(date)
            .to_period("M")
            .to_timestamp("M")
        )
        available = dates[dates <= target]

        if len(available) == 0:
            raise ValueError(
                f"No date available before {target.date()}."
            )

        return available[-1]
