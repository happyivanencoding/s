# Index technique du modèle sectoriel

## Document de référence

[Guide_production_modele_sectoriel.docx](docs/Guide_production_modele_sectoriel.docx)

| Partie | Contenu |
| --- | --- |
| Page 1 | Structure du projet et vue d'ensemble du pipeline |
| Pages 2 à 6 | Préparation, lancement, contrôle, reprise et historique |
| Pages 7 à 10 | Fichiers de code, entrées, tableaux, graphiques et exports |
| Annexe A, pages 11 à 12 | Appels du pipeline et circulation des données |
| Annexe B, pages 13 à 23 | Variables EU / US, piliers, macro, taux, votes et recommandations |
| Pages 24 à 25 | Portefeuille, performances et fonctions à consulter |

## Entrées de code

| Fichier | Rôle |
| --- | --- |
| Actualiser.bat | Commande de production et ouverture des résultats |
| actualiser_excel.py | Quatre étapes de production et bilan |
| local_config.py | Seule configuration persistante des sources et du modèle |
| model_secto.py | Interface SectorModel, pipeline et calculs |
| data_utils.py | Sources, dates, historique et conservation des fichiers Excel |
| excel_output.py | Output, portefeuille, formats, graphiques et écriture dans les XLSM |
| research_secto.py | Outils de recherche et backtests |

## Structure du modèle

model_secto.py contient six parties dans cet ordre :

1. SectorModel : choix du marché, production et lecture des résultats.
2. run_model : enchaînement des calculs d'un marché.
3. Sous-scores et piliers : mesures, rangs, Momentum et Volatility.
4. Macro et taux : contexte du mois.
5. Rangs et recommandations : scores agrégés, votes et choix des secteurs.
6. Exports : tableaux de scores en CSV.

Le schéma au début du fichier montre le parcours des données. actualiser_excel.py appelle le modèle puis excel_output.update_workbooks ; les détails techniques Excel restent dans excel_output.py et data_utils.py.

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

run() peut ajouter des observations à data_history.xlsx. Ce n'est pas une lecture sans effet sur l'historique.

run_production() calcule et exporte les CSV et la grille Excel. Pour mettre à jour les deux XLSM et leur présentation, utiliser Actualiser.bat ou python actualiser_excel.py.

save_outputs() écrit aussi *_historique_modele.csv. Cet export complet n'est pas actualisé par le lancement mensuel standard.

## Recherche et audit

research_secto.ipynb utilise SectorResearch. audit_modele_eu.ipynb et audit_modele_us.ipynb donnent une lecture détaillée du modèle. Les backtests de recherche sont distincts du portefeuille pondéré de production.

Les notebooks sont conservés sans leurs résultats d'exécution dans le dépôt. Les figures et CSV générés dans output/ peuvent être recréés.

Les règles détaillées sont maintenues dans le Word, avec les noms des fonctions source.
