"""Configuration commune aux modèles Europe et US.

Modifier ici les chemins, les secteurs, les variables et les pondérations.
Par défaut, les fichiers Excel sont dans le même dossier que ce fichier.
"""

from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
EU_EXCEL_FILE = str(PROJECT_DIR / "Score_Sectoriel_EU_NEW.xlsm")
US_EXCEL_FILE = str(PROJECT_DIR / "Score_Sectoriel_US_NEW.xlsm")
MACRO_EXCEL_FILE = str(PROJECT_DIR / "maj cycle macro2.xlsx")
SECTOR_DEVIATION_EXCEL_FILE = str(PROJECT_DIR / "reco_secto_facto.xlsx")  # Classeur d'allocation ICB existant.

HISTORY_CONFIG = {
    "enabled": True,
    "file": "data_history.xlsx",
}

# Europe : univers et pondérations

EU_SECTORS = [
    "Materials",
    "ConsStaples",
    "Pers. Goods",
    "Fin",
    "HealthCare",
    "Indus",
    "Oil",
    "Tech",
    "Telco",
    "Utili",
    "Travel & leisure",
    "Media",
    "Auto",
]

PORTFOLIO_DELTA = 0.05
PORTFOLIO_TILT = 0.20

N_TOP = 3
N_WORST = 3
EU_HISTORY_WINDOW = 60

EU_BASE_WEIGHTS = {
    "Leverage": 1 / 6,
    "Margin": 1 / 6,
    "Value": 1 / 6,
    "Momentum": 1 / 6,
    "Growth": 1 / 6,
    "Volatility": 1 / 6,
}

EU_REGIME_WEIGHTS = {
    "C": {
        "Leverage": 0.10,
        "Margin": 0.00,
        "Value": 0.00,
        "Momentum": 0.30,
        "Growth": 0.20,
        "Volatility": 0.40,
    },
    "R": {
        "Leverage": 0.10,
        "Margin": 0.10,
        "Value": 0.10,
        "Momentum": 0.35,
        "Growth": 0.35,
        "Volatility": 0.00,
    },
    "E": {
        "Leverage": 1 / 6,
        "Margin": 1 / 6,
        "Value": 1 / 6,
        "Momentum": 1 / 6,
        "Growth": 1 / 6,
        "Volatility": 1 / 6,
    },
    "SD": {
        "Leverage": 0.00,
        "Margin": 0.00,
        "Value": 0.30,
        "Momentum": 0.10,
        "Growth": 0.30,
        "Volatility": 0.30,
    },
}

# rank_order : 0 pour une mesure faible favorable, 1 pour une mesure élevée.
# Les colonnes sont les débuts des blocs sectoriels des feuilles source.
EU_HISTORICAL_VARIABLES = {
    "Leverage": {
        "net_debt_ebitda": {
            "sheet": "Leverage_FMA",
            "column": "AF",
            "rank_order": 0,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
        "fcf_total_debt": {
            "sheet": "Leverage_FMA",
            "column": "FB",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
        "debt_equity": {
            "sheet": "Leverage_FMA",
            "column": "GR",
            "rank_order": 0,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
    },
    "Margin": {
        "operating_margin": {
            "sheet": "Margin_FMA",
            "column": "BV",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": True,
        },
        "net_margin": {
            "sheet": "Margin_FMA",
            "column": "DL",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": True,
        },
        "ebitda_margin": {
            "sheet": "Margin_FMA",
            "column": "GR",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": True,
        },
    },
    "Value": {
        "price_fcf": {
            "sheet": "Valuation_FMA_hist",
            "column": "DL",
            "rank_order": 0,
            "exclude_financials_from_mean": True,
            "mix_cross_sectional_rank": False,
            "sector_windows": {"Tech": 36},
        },
        "ev_ebitda": {
            "sheet": "Valuation_FMA_hist",
            "column": "FB",
            "rank_order": 0,
            "exclude_financials_from_mean": True,
            "mix_cross_sectional_rank": False,
            "sector_windows": {"Tech": 36},
        },
        "price_sales": {
            "sheet": "Valuation_FMA_hist",
            "column": "GR",
            "rank_order": 0,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
    },
    "Growth": {
        "eps_growth": {
            "sheet": "Growth_FMA",
            "column": "AF",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
        "sales_growth": {
            "sheet": "Growth_FMA",
            "column": "DL",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
        "ebitda_growth": {
            "sheet": "Growth_FMA",
            "column": "FB",
            "rank_order": 1,
            "exclude_financials_from_mean": True,
            "mix_cross_sectional_rank": False,
        },
    },
}

PILLAR_EXCLUSIONS = {
    "Growth": {
        "Fin": {"ebitda_growth"},
    },
    "Value": {
        "Fin": {"price_fcf"},
        "Tech": {"price_sales"},
    },
}

EU_PILLAR_ROUNDING = {
    "Growth": 14,
}

EU_MOMENTUM_CONFIG = {
    "sheet": "MOM_FMA",
    "price_column": "AT",
    "revision_up_column": "DL",
    "revision_down_column": "DZ",
    "revision_unchanged_column": "EN",
    "short_horizon": 6,
    "long_horizon": 12,
    "current_month_sector": "Tech",
}

EU_VOLATILITY_CONFIG = {
    "sheet_returns": "Returns_EQ",
    "returns_start_row": 7,
    "date_column": "P",
    "benchmark_returns_column": "Q",
    "benchmark_name": "FTSE World Europe",
    "returns_start_column": "R",
    "volatility_offset": 6,
    "offset_downside": 18,
}

EU_MACRO_CONFIG = {
    "sheet": "Europe",
    "start_row": 4,
    "date_column": "A",
    "score_column": "U",
    "regime_column": "AB",
}

EU_RATE_SIGNAL_CONFIG = {
    "sheet": "Cycle macro",
    "start_row": 4,
    "date_column": "A",
    "us10y_column": "R",
    "alpha_ewma": 0.715,
    "percentile_threshold": 0.85,
}

EU_NO_VOTE_RULES = [{ "sector": "Travel & leisure", "start_date": "2011-10-31", }]

EU_MACRO_VOTE_WEIGHT = 2
EU_RATE_OVERLAY_PILLARS = ["Leverage", "Value"]

EU_BACKTEST_VARIABLES = {
    "Leverage": ["net_debt_ebitda", "fcf_total_debt", "debt_equity"],
    "Margin": ["operating_margin", "net_margin", "ebitda_margin"],
    "Value": ["price_fcf", "ev_ebitda", "price_sales"],
    "Momentum": ["momentum_6m_1m", "momentum_12m_1m", "earnings_revision_ratio"],
    "Growth": ["eps_growth", "sales_growth", "ebitda_growth"],
    "Volatility": ["volatility_6m", "downside_volatility_18m"],
}

EU_PILLAR_COMPOSITION = {
    "Leverage": {"default": ["net_debt_ebitda", "fcf_total_debt", "debt_equity"]},
    "Margin": {"default": ["operating_margin", "net_margin", "ebitda_margin"]},
    "Value": {
        "default": ["price_fcf", "ev_ebitda", "price_sales"],
        "by_sector": {
            "Fin": ["ev_ebitda", "price_sales"],
            "Tech": ["price_fcf", "ev_ebitda"],
        },
    },
    "Growth": {
        "default": ["eps_growth", "sales_growth", "ebitda_growth"],
        "by_sector": {"Fin": ["eps_growth", "sales_growth"]},
    },
    "Momentum": {
        "default": ["momentum_6m_1m", "momentum_12m_1m", "earnings_revision_ratio"],
    },
    "Volatility": {
        "default": ["volatility_6m", "downside_volatility_18m"],
    },
}

EU_MACRO_ADJUSTED_WEIGHTS = {
    "Leverage": 1 / 7,
    "Margin": 1 / 7,
    "Value": 1 / 7,
    "Momentum": 1 / 7,
    "Growth": 1 / 7,
    "Volatility": 1 / 7,
    "Macro": 1 / 7,
}

EU_TOP_VOTE_PILLARS = list(EU_BASE_WEIGHTS)
EU_BOTTOM_VOTE_PILLARS = list(EU_BASE_WEIGHTS)


# États-Unis : univers et pondérations

US_SECTORS = [
    "Materials",
    "ConsStaples",
    "Retail",
    "Fin",
    "HealthCare",
    "Indus",
    "Oil",
    "Tech",
    "Telco",
    "Utili",
    "Travel & leisure",
    "Media",
]

N_TOP_US = 3
N_WORST_US = 3
US_HISTORY_WINDOW = 60

US_BASE_WEIGHTS = {
    "Leverage": 0.20,
    "Margin": 0.00,
    "Value": 0.20,
    "Momentum": 0.20,
    "Growth": 0.20,
    "Volatility": 0.20,
}

US_REGIME_WEIGHTS = {
    "C": {
        "Leverage": 0.20,
        "Margin": 0.00,
        "Value": 0.20,
        "Momentum": 0.00,
        "Growth": 0.20,
        "Volatility": 0.40,
    },
    "R": {
        "Leverage": 0.00,
        "Margin": 0.00,
        "Value": 0.00,
        "Momentum": 0.50,
        "Growth": 0.50,
        "Volatility": 0.00,
    },
    "E": {
        "Leverage": 0.20,
        "Margin": 0.00,
        "Value": 0.20,
        "Momentum": 0.20,
        "Growth": 0.20,
        "Volatility": 0.20,
    },
    "SD": {
        "Leverage": 0.10,
        "Margin": 0.00,
        "Value": 0.15,
        "Momentum": 0.30,
        "Growth": 0.15,
        "Volatility": 0.30,
    },
}

US_HISTORICAL_VARIABLES = {
    "Leverage": {
        "net_debt_ebitda": {
            "sheet": "Leverage_FMA",
            "column": "AD",
            "rank_order": 0,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": True,
        },
        "fcf_total_debt": {
            "sheet": "Leverage_FMA",
            "column": "EQ",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": True,
        },
    },
    "Margin": {
        "operating_margin": {
            "sheet": "Margin_FMA",
            "column": "BQ",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": True,
        },
        "net_margin": {
            "sheet": "Margin_FMA",
            "column": "DD",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": True,
        },
        "ebitda_margin": {
            "sheet": "Margin_FMA",
            "column": "GD",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": True,
        },
    },
    "Value": {
        "pe": {
            "sheet": "Valuation_FMA_hist",
            "column": "AD",
            "rank_order": 0,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
        "price_fcf": {
            "sheet": "Valuation_FMA_hist",
            "column": "DD",
            "rank_order": 0,
            "exclude_financials_from_mean": True,
            "mix_cross_sectional_rank": False,
            "sector_windows": {"Tech": 36},
        },
        "ev_ebitda": {
            "sheet": "Valuation_FMA_hist",
            "column": "EQ",
            "rank_order": 0,
            "exclude_financials_from_mean": True,
            "mix_cross_sectional_rank": False,
            "sector_windows": {"Tech": 36},
        },
        "price_sales": {
            "sheet": "Valuation_FMA_hist",
            "column": "GD",
            "rank_order": 0,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
    },
    "Growth": {
        "eps_growth": {
            "sheet": "Growth_FMA",
            "column": "AD",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
        "cash_flow_growth": {
            "sheet": "Growth_FMA",
            "column": "BQ",
            "rank_order": 1,
            "exclude_financials_from_mean": True,
            "mix_cross_sectional_rank": False,
        },
        "sales_growth": {
            "sheet": "Growth_FMA",
            "column": "DD",
            "rank_order": 1,
            "exclude_financials_from_mean": False,
            "mix_cross_sectional_rank": False,
        },
        "ebitda_growth": {
            "sheet": "Growth_FMA",
            "column": "EQ",
            "rank_order": 1,
            "exclude_financials_from_mean": True,
            "mix_cross_sectional_rank": False,
        },
    },
}

US_PILLAR_COMPOSITION = {
    "Leverage": {"default": ["net_debt_ebitda", "fcf_total_debt"]},
    "Margin": {"default": ["operating_margin", "net_margin", "ebitda_margin"]},
    "Value": {
        "default": ["price_fcf", "ev_ebitda", "price_sales"],
        "by_sector": {
            "Fin": ["pe", "price_sales"],
            "Tech": ["price_fcf", "ev_ebitda"],
        },
    },
    "Growth": {
        "default": ["cash_flow_growth", "ebitda_growth"],
        "by_sector": {"Fin": ["eps_growth", "sales_growth"]},
    },
    "Momentum": {
        "default": ["momentum_6m_1m", "momentum_12m_1m", "earnings_revision_ratio"],
    },
    "Volatility": {
        "default": ["volatility_6m", "downside_volatility_18m"],
    },
}

US_PILLAR_ROUNDING = {}

US_MOMENTUM_CONFIG = {
    "sheet": "MOM_FMA",
    "price_column": "AQ",
    "revision_up_column": "DD",
    "revision_down_column": "DQ",
    "revision_unchanged_column": "ED",
    "short_horizon": 6,
    "long_horizon": 12,
    "current_month_sector": "Tech",
    "scale_0_10": False,
}

US_VOLATILITY_CONFIG = {
    "sheet_returns": "Returns_EQ",
    "returns_start_row": 7,
    "date_column": "O",
    "benchmark_returns_column": "P",
    "benchmark_name": "WIUSAU Index",
    "returns_start_column": "Q",
    "volatility_offset": 6,
    "offset_downside": 18,
}

US_MACRO_CONFIG = {
    "sheet": "US",
    "start_row": 4,
    "date_column": "A",
    "score_column": "Q",
    "regime_column": "U",
}

US_RATE_SIGNAL_CONFIG = {
    "sheet": "Cycle macro",
    "start_row": 4,
    "date_column": "A",
    "us10y_column": "U",
    "alpha_ewma": 0.715,
    "percentile_threshold": 0.85,
}

US_MACRO_ADJUSTED_WEIGHTS = {
    "Leverage": 1 / 6,
    "Margin": 0.00,
    "Value": 1 / 6,
    "Momentum": 1 / 6,
    "Growth": 1 / 6,
    "Volatility": 1 / 6,
    "Macro": 1 / 6,
}

US_TOP_VOTE_PILLARS = ["Leverage", "Margin", "Value", "Momentum", "Growth", "Volatility"]

US_BOTTOM_VOTE_PILLARS = ["Leverage", "Value", "Momentum", "Growth", "Volatility"]

US_NO_VOTE_RULES = [{ "sector": "Travel & leisure", "start_date": "2011-10-31", }]
US_MACRO_VOTE_WEIGHT = 1
US_RATE_OVERLAY_PILLARS = ["Leverage", "Value"]

US_BACKTEST_VARIABLES = {
    "Leverage": ["net_debt_ebitda", "fcf_total_debt"],
    "Margin": ["operating_margin", "net_margin", "ebitda_margin"],
    "Value": ["pe", "price_fcf", "ev_ebitda", "price_sales"],
    "Momentum": ["momentum_6m_1m", "momentum_12m_1m", "earnings_revision_ratio"],
    "Growth": ["eps_growth", "cash_flow_growth", "sales_growth", "ebitda_growth"],
    "Volatility": ["volatility_6m", "downside_volatility_18m"],
}


# Configuration transmise au calcul de chaque marché.
EU_CONFIG = {
    "name": "EU",
    "output_prefix": "eu",
    "history_prefix": "",  # Préfixe des clés dans Historique_data pour distinguer les marchés.
    "excel_file": EU_EXCEL_FILE,
    "sectors": EU_SECTORS,
    "n_top": N_TOP,  # Nombre maximal de secteurs retenus côté Positive.
    "n_worst": N_WORST,  # Nombre maximal de secteurs retenus côté Negative.
    "history_window": EU_HISTORY_WINDOW,  # Mois antérieurs pour le rang historique, plus le mois courant.
    "base_weights": EU_BASE_WEIGHTS,  # Poids des piliers pour le score de base, avant le cycle.
    "regime_weights": EU_REGIME_WEIGHTS,  # Poids des piliers selon le régime C, R, E ou SD (tilt macro).
    "macro_adjusted_weights": EU_MACRO_ADJUSTED_WEIGHTS,  # Poids des rangs des piliers et du rang macro dans le score global.
    "historical_variables": EU_HISTORICAL_VARIABLES,  # Source, sens du rang et méthode H ou H+X de chaque variable.
    "pillar_composition": EU_PILLAR_COMPOSITION,  # Variables retenues par pilier, avec exceptions par secteur.
    "pillar_rounding": EU_PILLAR_ROUNDING,  # Arrondi appliqué aux scores des piliers avant classement.
    "momentum": {**EU_MOMENTUM_CONFIG, "scale_0_10": True},  # Prix et révisions de bénéfices ; rangs ramenés sur 10 en EU.
    "volatility": EU_VOLATILITY_CONFIG,  # Série de rendements et fenêtres de risque ; nom du benchmark.
    "macro": EU_MACRO_CONFIG,  # Feuille et colonnes source du score et du cycle macro.
    "rate_signal": EU_RATE_SIGNAL_CONFIG,  # Taux US 10 ans et seuils du signal On/Off.
    "top_vote_pillars": EU_TOP_VOTE_PILLARS,  # Piliers qui votent pour les secteurs Top.
    "bottom_vote_pillars": EU_BOTTOM_VOTE_PILLARS,  # Piliers qui votent pour les secteurs Worst.
    "no_vote_rules": EU_NO_VOTE_RULES,  # Secteurs et périodes où les votes Top/Worst sont neutralisés.
    "macro_vote_weight": EU_MACRO_VOTE_WEIGHT,  # Nombre de votes apportés par le rang du tilt macro.
    "rate_overlay_pillars": EU_RATE_OVERLAY_PILLARS,  # Piliers recevant un vote supplémentaire quand le taux est On.
    "backtest_variables": EU_BACKTEST_VARIABLES,  # Variables proposées pour les essais, sans remplacer la sélection en production.
}

US_CONFIG = {
    "name": "US",
    "output_prefix": "us",
    "history_prefix": "US_",  # Préfixe des clés dans Historique_data pour distinguer les marchés.
    "excel_file": US_EXCEL_FILE,
    "sectors": US_SECTORS,
    "n_top": N_TOP_US,  # Nombre maximal de secteurs retenus côté Positive.
    "n_worst": N_WORST_US,  # Nombre maximal de secteurs retenus côté Negative.
    "history_window": US_HISTORY_WINDOW,  # Mois antérieurs pour le rang historique, plus le mois courant.
    "base_weights": US_BASE_WEIGHTS,  # Poids des piliers pour le score de base, avant le cycle.
    "regime_weights": US_REGIME_WEIGHTS,  # Poids des piliers selon le régime C, R, E ou SD (tilt macro).
    "macro_adjusted_weights": US_MACRO_ADJUSTED_WEIGHTS,  # Poids des rangs des piliers et du rang macro dans le score global.
    "historical_variables": US_HISTORICAL_VARIABLES,  # Source, sens du rang et méthode H ou H+X de chaque variable.
    "pillar_composition": US_PILLAR_COMPOSITION,  # Variables retenues par pilier, avec exceptions par secteur.
    "pillar_rounding": US_PILLAR_ROUNDING,  # Arrondi appliqué aux scores des piliers avant classement.
    "momentum": US_MOMENTUM_CONFIG,  # Prix et révisions de bénéfices ; rangs bruts non ramenés sur 10 en US.
    "volatility": US_VOLATILITY_CONFIG,  # Série de rendements et fenêtres de risque ; nom du benchmark.
    "macro": US_MACRO_CONFIG,  # Feuille et colonnes source du score et du cycle macro.
    "rate_signal": US_RATE_SIGNAL_CONFIG,  # Taux US 10 ans et seuils du signal On/Off.
    "top_vote_pillars": US_TOP_VOTE_PILLARS,  # Piliers qui votent pour les secteurs Top.
    "bottom_vote_pillars": US_BOTTOM_VOTE_PILLARS,  # Piliers qui votent pour les secteurs Worst.
    "no_vote_rules": US_NO_VOTE_RULES,  # Secteurs et périodes où les votes Top/Worst sont neutralisés.
    "macro_vote_weight": US_MACRO_VOTE_WEIGHT,  # Nombre de votes apportés par le rang du tilt macro.
    "rate_overlay_pillars": US_RATE_OVERLAY_PILLARS,  # Piliers recevant un vote supplémentaire quand le taux est On.
    "backtest_variables": US_BACKTEST_VARIABLES,  # Variables proposées pour les essais, sans remplacer la sélection en production.
}
