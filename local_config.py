# -*- coding: utf-8 -*-
"""
Configuration du modèle sectoriel Europe.

Ce fichier est versionné dans Git.
Les chemins Excel sont volontairement vides dans la version publique.

Pour un poste local, utiliser local_config_private.py pour renseigner
uniquement les chemins des fichiers Excel.
"""

FICHIER_EXCEL_EU = ""
FICHIER_EXCEL_MACRO = ""

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

SECTEURS_SANS_VOTE_TOP_WORST = {
    "Travel & leisure",
}

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
