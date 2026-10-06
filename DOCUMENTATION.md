# Modèle sectoriel EU / US

## Structure

- `local_config.py` : paramètres EU / US.
- `data_io.py` : lecture des fichiers Excel et gestion de l'historique.
- `sector_core.py` : calculs communs du modèle.
- `model_secto.py` : interface principale via `SectorModel`.
- `research_secto.py` : plots et backtests.
- `research_secto.ipynb` : interface de recherche.
- `audit_modele_eu.ipynb` / `audit_modele_us.ipynb` : contrôles détaillés.

## Production mensuelle

Après le refresh et la sauvegarde des fichiers Excel :

```python
from model_secto import SectorModel

SectorModel("EU").run_production()
SectorModel("US").run_production()
```

Les fichiers sont écrits dans `output/`.

Pour chaque marché :

- `*_recommandations_latest.csv` : recommandations les plus récentes.
- `*_piliers_latest_with_reco.csv` : scores des piliers avec recommandation.
- `*_piliers_latest_available.csv` : derniers scores de piliers disponibles.

`run_production()` relit toujours les fichiers source avant de lancer le calcul.

## Historique

Les données utilisées par le modèle sont conservées dans :

```text
data_history.parquet
```

L'historique est mis à jour de façon incrémentale.

Pour chaque série et chaque mois :

- si le mois n'existe pas encore, il est ajouté ;
- s'il existe déjà, la première observation enregistrée est conservée.

Une révision ultérieure du fichier source ne remplace donc pas une valeur déjà figée.

## Utilisation du modèle

```python
from model_secto import SectorModel

model = SectorModel("EU")
results = model.run()
```

Quelques méthodes utiles :

```python
model.summary()
model.variables("Value")
model.variable_definitions("Value")
model.composition("Value")
model.variable_table("Value")
model.recommendations()
```

## Test d'une autre composition

Il est possible de modifier les variables utilisées dans un pilier :

```python
model = SectorModel(
    universe="EU",
    variables={
        "Value": [
            "price_fcf",
            "ev_ebitda",
        ],
    },
)

results = model.run()
```

Les piliers non précisés gardent leur configuration par défaut.

Les règles spécifiques à certains secteurs restent appliquées. Une combinaison incompatible est refusée avant le calcul.

## Research

Pour les plots et backtests :

```python
from research_secto import SectorResearch

research = SectorResearch(model)

summary, backtests, fig, ax = research.backtest(
    pillar="Value",
    top_n=3,
    plot_series=["top", "benchmark"],
)
```

Le notebook `research_secto.ipynb` permet de modifier directement l'univers, les variables, les dates et les paramètres de backtest.

## Remarque

La dernière date disponible pour les piliers peut être plus récente que la dernière recommandation finale si les données macro ne sont pas encore disponibles pour le même mois.
