# Index technique du modèle sectoriel

## Document de référence

[Guide_production_modele_sectoriel.docx](docs/Guide_production_modele_sectoriel.docx)

| Partie du guide | Contenu |
| --- | --- |
| Production mensuelle | Entrées, lancement et lecture des résultats |
| Données et pipeline | Sources Excel, historique, Output et Dashboard |
| Méthode EU / US | Variables, scores, piliers, macro, taux et recommandations |
| Portefeuille et code | Performances, pondérations et fichiers de référence |
| Annexe A | Pipeline et appels entre actualiser_excel, data_utils, model_secto et excel_output |
| Annexes B et C | Sens des variables, H / H+X, macro, votes, score global et benchmark |

## Entrées de code

| Fichier | Rôle |
| --- | --- |
| Actualiser.bat | Commande de production et ouverture des résultats |
| actualiser_excel.py | Quatre étapes de production et bilan |
| local_config.py | Seule configuration persistante des sources et du modèle |
| model_secto.py | Interface SectorModel, pipeline et calculs |
| data_utils.py | Sources, dates, historique et conservation des fichiers Excel |
| excel_output.py | Output, portefeuille, copie des valeurs et conservation des formats du modèle XLSM |
| research_secto.py | Outils de recherche et backtests |

## Structure du modèle

model_secto.py contient six parties dans cet ordre :

1. SectorModel : choix du marché, production et lecture des résultats.
2. run_model : enchaînement des calculs d'un marché.
3. Sous-scores et piliers : mesures, rangs, Momentum et Volatility.
4. Macro et taux : contexte du mois.
5. Rangs et recommandations : scores agrégés, votes et choix des secteurs.
6. Exports : tableaux de scores en CSV.

Le schéma au début du fichier montre le parcours des données. actualiser_excel.py appelle le modèle puis excel_output.update_workbooks. La feuille Output est actualisée en valeurs seulement, sur une copie du XLSM existant ; les formats, les liaisons du Dashboard et la position des graphiques sont conservés. excel_output.py ajoute les nouveaux mois dans Reco_histo, rafraîchit les graphiques et exporte Output avec sa présentation, sans réappliquer NumberFormat ni reconstruire le Dashboard.

## Utilisation Python

```python
from model_secto import SectorModel

model = SectorModel("EU")
results = model.run()
model.variables("Value")
model.composition("Value")
model.variable_table("Value")
model.recommendations()
```

run() peut ajouter des observations à data_history.xlsx (feuille Historique_data). Ce n'est pas une lecture sans effet sur l'historique.

run_production() calcule et exporte les CSV et la grille Excel. Pour mettre à jour les deux XLSM et leur présentation, utiliser Actualiser.bat ou python actualiser_excel.py.

save_outputs() écrit aussi *_historique_modele.csv. Cet export complet n'est pas actualisé par le lancement mensuel standard.

## Recherche et audit

research_secto.ipynb utilise SectorResearch. audit_modele_eu.ipynb et audit_modele_us.ipynb donnent une lecture détaillée du modèle. Les backtests de recherche sont distincts du portefeuille pondéré de production.

Les notebooks sont conservés sans leurs résultats d'exécution dans le dépôt. Les figures et CSV générés dans output/ peuvent être recréés.

Le Word comprend une synthèse pratique et des annexes de méthodologie et d'architecture Python. Les correspondances précises de cellules et les paramètres techniques restent dans local_config.py et les fonctions Python.
