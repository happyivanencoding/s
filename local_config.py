# -*- coding: utf-8 -*-
"""Configuration des modèles sectoriels Europe et US."""

FICHIER_EXCEL_EU = ""
FICHIER_EXCEL_US = ""
FICHIER_EXCEL_MACRO = ""

CONFIG_HISTORIQUE = {
    "actif": True,
    "fichier": "data_history.parquet",
}

SECTEURS = [
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

N_TOP = 3
N_WORST = 3
FENETRE_HISTORIQUE = 60

POIDS_BASE = {
    "Leverage": 1 / 6,
    "Margin": 1 / 6,
    "Value": 1 / 6,
    "Momentum": 1 / 6,
    "Growth": 1 / 6,
    "Volatility": 1 / 6,
}

POIDS_REGIME = {
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

VARIABLES_HISTORIQUES = {
    "Leverage": {
        "net_debt_ebitda": {
            "sheet": "Leverage_FMA",
            "colonne": "AF",
            "ordre_rank": 0,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
        "fcf_total_debt": {
            "sheet": "Leverage_FMA",
            "colonne": "FB",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
        "debt_equity": {
            "sheet": "Leverage_FMA",
            "colonne": "GR",
            "ordre_rank": 0,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
    },
    "Margin": {
        "operating_margin": {
            "sheet": "Margin_FMA",
            "colonne": "BV",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": True,
        },
        "net_margin": {
            "sheet": "Margin_FMA",
            "colonne": "DL",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": True,
        },
        "ebitda_margin": {
            "sheet": "Margin_FMA",
            "colonne": "GR",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": True,
        },
    },
    "Value": {
        "price_fcf": {
            "sheet": "Valuation_FMA_hist",
            "colonne": "DL",
            "ordre_rank": 0,
            "moyenne_sans_finance": True,
            "mix_rank_transversal": False,
            "fenetre_par_secteur": {"Tech": 36},
        },
        "ev_ebitda": {
            "sheet": "Valuation_FMA_hist",
            "colonne": "FB",
            "ordre_rank": 0,
            "moyenne_sans_finance": True,
            "mix_rank_transversal": False,
            "fenetre_par_secteur": {"Tech": 36},
        },
        "price_sales": {
            "sheet": "Valuation_FMA_hist",
            "colonne": "GR",
            "ordre_rank": 0,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
    },
    "Growth": {
        "eps_growth": {
            "sheet": "Growth_FMA",
            "colonne": "AF",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
        "sales_growth": {
            "sheet": "Growth_FMA",
            "colonne": "DL",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
        "ebitda_growth": {
            "sheet": "Growth_FMA",
            "colonne": "FB",
            "ordre_rank": 1,
            "moyenne_sans_finance": True,
            "mix_rank_transversal": False,
        },
    },
}

EXCLUSIONS_PILIER = {
    "Growth": {
        "Fin": {"ebitda_growth"},
    },
    "Value": {
        "Fin": {"price_fcf"},
        "Tech": {"price_sales"},
    },
}

ARRONDI_PILIER = {
    "Growth": 14,
}

CONFIG_MOMENTUM = {
    "sheet": "MOM_FMA",
    "colonne_prix": "AT",
    "colonne_revision_up": "DL",
    "colonne_revision_down": "DZ",
    "colonne_revision_unchanged": "EN",
    "horizon_court": 6,
    "horizon_long": 12,
    "secteur_mois_courant": "Tech",
}

CONFIG_VOLATILITE = {
    "sheet_retours": "Returns_EQ",
    "ligne_debut_retours": 7,
    "colonne_date": "P",
    "colonne_debut_retours": "R",
    "offset_volatilite": 6,
    "offset_downside": 18,
}

CONFIG_MACRO_EU = {
    "sheet": "Europe",
    "ligne_debut": 4,
    "colonne_date": "A",
    "colonne_score": "R",
    "colonne_regime": "V",
}

CONFIG_SIGNAL_TAUX = {
    "sheet": "Cycle macro",
    "ligne_debut": 4,
    "colonne_date": "A",
    "colonne_us10y": "R",
    "alpha_ewma": 0.715,
    "seuil_percentile": 0.85,
}

REGLES_SANS_VOTE_TOP_WORST = [
    {
        "secteur": "Travel & leisure",
        "date_debut": "2011-10-31",
    },
]

POIDS_VOTE_MACRO = 2
PILIERS_RATE_OVERLAY = [
    "Leverage",
    "Value",
]

VARIABLES_BACKTEST = {
    "Leverage": [
        "net_debt_ebitda",
        "fcf_total_debt",
        "debt_equity",
    ],
    "Margin": [
        "operating_margin",
        "net_margin",
        "ebitda_margin",
    ],
    "Value": [
        "price_fcf",
        "ev_ebitda",
        "price_sales",
    ],
    "Momentum": [
        "momentum_6m_1m",
        "momentum_12m_1m",
        "earnings_revision_ratio",
    ],
    "Growth": [
        "eps_growth",
        "sales_growth",
        "ebitda_growth",
    ],
    "Volatility": [
        "volatility_6m",
        "downside_volatility_18m",
    ],
}

# Europe

COMPOSITION_PILIERS = {
    "Leverage": {"default": ["net_debt_ebitda", "fcf_total_debt", "debt_equity"]},
    "Margin": {"default": ["operating_margin", "net_margin", "ebitda_margin"]},
    "Value": {
        "default": ["price_fcf", "ev_ebitda", "price_sales"],
        "par_secteur": {
            "Fin": ["ev_ebitda", "price_sales"],
            "Tech": ["price_fcf", "ev_ebitda"],
        },
    },
    "Growth": {
        "default": ["eps_growth", "sales_growth", "ebitda_growth"],
        "par_secteur": {"Fin": ["eps_growth", "sales_growth"]},
    },
}

POIDS_GLOBAL_MACRO = {
    "Leverage": 1 / 7,
    "Margin": 1 / 7,
    "Value": 1 / 7,
    "Momentum": 1 / 7,
    "Growth": 1 / 7,
    "Volatility": 1 / 7,
    "Macro": 1 / 7,
}

PILIERS_TOP_VOTE = list(POIDS_BASE)
PILIERS_BOTTOM_VOTE = list(POIDS_BASE)


# US

SECTEURS_US = [
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
FENETRE_HISTORIQUE_US = 60

POIDS_BASE_US = {
    "Leverage": 0.20,
    "Margin": 0.00,
    "Value": 0.20,
    "Momentum": 0.20,
    "Growth": 0.20,
    "Volatility": 0.20,
}

POIDS_REGIME_US = {
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

VARIABLES_HISTORIQUES_US = {
    "Leverage": {
        "net_debt_ebitda": {
            "sheet": "Leverage_FMA",
            "colonne": "AD",
            "ordre_rank": 0,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": True,
        },
        "fcf_total_debt": {
            "sheet": "Leverage_FMA",
            "colonne": "EQ",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": True,
        },
    },
    "Margin": {
        "operating_margin": {
            "sheet": "Margin_FMA",
            "colonne": "BQ",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": True,
        },
        "net_margin": {
            "sheet": "Margin_FMA",
            "colonne": "DD",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": True,
        },
        "ebitda_margin": {
            "sheet": "Margin_FMA",
            "colonne": "GD",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": True,
        },
    },
    "Value": {
        "pe": {
            "sheet": "Valuation_FMA_hist",
            "colonne": "AD",
            "ordre_rank": 0,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
        "price_fcf": {
            "sheet": "Valuation_FMA_hist",
            "colonne": "DD",
            "ordre_rank": 0,
            "moyenne_sans_finance": True,
            "mix_rank_transversal": False,
            "fenetre_par_secteur": {"Tech": 36},
        },
        "ev_ebitda": {
            "sheet": "Valuation_FMA_hist",
            "colonne": "EQ",
            "ordre_rank": 0,
            "moyenne_sans_finance": True,
            "mix_rank_transversal": False,
            "fenetre_par_secteur": {"Tech": 36},
        },
        "price_sales": {
            "sheet": "Valuation_FMA_hist",
            "colonne": "GD",
            "ordre_rank": 0,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
    },
    "Growth": {
        "eps_growth": {
            "sheet": "Growth_FMA",
            "colonne": "AD",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
        "cash_flow_growth": {
            "sheet": "Growth_FMA",
            "colonne": "BQ",
            "ordre_rank": 1,
            "moyenne_sans_finance": True,
            "mix_rank_transversal": False,
        },
        "sales_growth": {
            "sheet": "Growth_FMA",
            "colonne": "DD",
            "ordre_rank": 1,
            "moyenne_sans_finance": False,
            "mix_rank_transversal": False,
        },
        "ebitda_growth": {
            "sheet": "Growth_FMA",
            "colonne": "EQ",
            "ordre_rank": 1,
            "moyenne_sans_finance": True,
            "mix_rank_transversal": False,
        },
    },
}

COMPOSITION_PILIERS_US = {
    "Leverage": {"default": ["net_debt_ebitda", "fcf_total_debt"]},
    "Margin": {"default": ["operating_margin", "net_margin", "ebitda_margin"]},
    "Value": {
        "default": ["price_fcf", "ev_ebitda", "price_sales"],
        "par_secteur": {
            "Fin": ["pe", "price_sales"],
            "Tech": ["price_fcf", "ev_ebitda"],
        },
    },
    "Growth": {
        "default": ["cash_flow_growth", "ebitda_growth"],
        "par_secteur": {"Fin": ["eps_growth", "sales_growth"]},
    },
}

ARRONDI_PILIER_US = {}

CONFIG_MOMENTUM_US = {
    "sheet": "MOM_FMA",
    "colonne_prix": "AQ",
    "colonne_revision_up": "DD",
    "colonne_revision_down": "DQ",
    "colonne_revision_unchanged": "ED",
    "horizon_court": 6,
    "horizon_long": 12,
    "secteur_mois_courant": "Tech",
    "echelle_0_10": False,
}

CONFIG_VOLATILITE_US = {
    "sheet_retours": "Returns_EQ",
    "ligne_debut_retours": 7,
    "colonne_date": "O",
    "colonne_debut_retours": "Q",
    "offset_volatilite": 6,
    "offset_downside": 18,
}

CONFIG_MACRO_US = {
    "sheet": "US",
    "ligne_debut": 4,
    "colonne_date": "A",
    "colonne_score": "Y",
    "colonne_regime": "Z",
}

CONFIG_SIGNAL_TAUX_US = {
    "sheet": "Cycle macro",
    "ligne_debut": 4,
    "colonne_date": "A",
    "colonne_us10y": "U",
    "alpha_ewma": 0.715,
    "seuil_percentile": 0.85,
}

POIDS_GLOBAL_MACRO_US = {
    "Leverage": 1 / 6,
    "Margin": 0.00,
    "Value": 1 / 6,
    "Momentum": 1 / 6,
    "Growth": 1 / 6,
    "Volatility": 1 / 6,
    "Macro": 1 / 6,
}

PILIERS_TOP_VOTE_US = [
    "Leverage",
    "Margin",
    "Value",
    "Momentum",
    "Growth",
    "Volatility",
]

PILIERS_BOTTOM_VOTE_US = [
    "Leverage",
    "Value",
    "Momentum",
    "Growth",
    "Volatility",
]

REGLES_SANS_VOTE_TOP_WORST_US = [
    {
        "secteur": "Travel & leisure",
        "date_debut": "2011-10-31",
    },
]
POIDS_VOTE_MACRO_US = 1
PILIERS_RATE_OVERLAY_US = ["Leverage", "Value"]

VARIABLES_BACKTEST_US = {
    "Leverage": ["net_debt_ebitda", "fcf_total_debt"],
    "Margin": ["operating_margin", "net_margin", "ebitda_margin"],
    "Value": ["pe", "price_fcf", "ev_ebitda", "price_sales"],
    "Momentum": ["momentum_6m_1m", "momentum_12m_1m", "earnings_revision_ratio"],
    "Growth": ["eps_growth", "cash_flow_growth", "sales_growth", "ebitda_growth"],
    "Volatility": ["volatility_6m", "downside_volatility_18m"],
}


MODELE_EU = {
    "nom": "EU",
    "prefixe_sortie": "eu",
    "historique_prefixe": "",
    "fichier_excel": FICHIER_EXCEL_EU,
    "env_excel": "SCORE_SECTORIEL_EU_XLSM",
    "secteurs": SECTEURS,
    "n_top": N_TOP,
    "n_worst": N_WORST,
    "fenetre_historique": FENETRE_HISTORIQUE,
    "poids_base": POIDS_BASE,
    "poids_regime": POIDS_REGIME,
    "poids_global_macro": POIDS_GLOBAL_MACRO,
    "variables_historiques": VARIABLES_HISTORIQUES,
    "composition_piliers": COMPOSITION_PILIERS,
    "arrondi_pilier": ARRONDI_PILIER,
    "momentum": {**CONFIG_MOMENTUM, "echelle_0_10": True},
    "volatilite": CONFIG_VOLATILITE,
    "macro": CONFIG_MACRO_EU,
    "signal_taux": CONFIG_SIGNAL_TAUX,
    "piliers_top_vote": PILIERS_TOP_VOTE,
    "piliers_bottom_vote": PILIERS_BOTTOM_VOTE,
    "regles_sans_vote": REGLES_SANS_VOTE_TOP_WORST,
    "poids_vote_macro": POIDS_VOTE_MACRO,
    "piliers_rate_overlay": PILIERS_RATE_OVERLAY,
    "variables_backtest": VARIABLES_BACKTEST,
}

MODELE_US = {
    "nom": "US",
    "prefixe_sortie": "us",
    "historique_prefixe": "US_",
    "fichier_excel": FICHIER_EXCEL_US,
    "env_excel": "SCORE_SECTORIEL_US_XLSM",
    "secteurs": SECTEURS_US,
    "n_top": N_TOP_US,
    "n_worst": N_WORST_US,
    "fenetre_historique": FENETRE_HISTORIQUE_US,
    "poids_base": POIDS_BASE_US,
    "poids_regime": POIDS_REGIME_US,
    "poids_global_macro": POIDS_GLOBAL_MACRO_US,
    "variables_historiques": VARIABLES_HISTORIQUES_US,
    "composition_piliers": COMPOSITION_PILIERS_US,
    "arrondi_pilier": ARRONDI_PILIER_US,
    "momentum": CONFIG_MOMENTUM_US,
    "volatilite": CONFIG_VOLATILITE_US,
    "macro": CONFIG_MACRO_US,
    "signal_taux": CONFIG_SIGNAL_TAUX_US,
    "piliers_top_vote": PILIERS_TOP_VOTE_US,
    "piliers_bottom_vote": PILIERS_BOTTOM_VOTE_US,
    "regles_sans_vote": REGLES_SANS_VOTE_TOP_WORST_US,
    "poids_vote_macro": POIDS_VOTE_MACRO_US,
    "piliers_rate_overlay": PILIERS_RATE_OVERLAY_US,
    "variables_backtest": VARIABLES_BACKTEST_US,
}

