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
| actualiser_excel.py | Modes production, rebuild et replay ; ouverture des résultats de production |
| local_config.py | Seule configuration persistante des sources et du modèle |
| model_secto.py | Interface SectorModel, pipeline et calculs |
| data_utils.py | Sources, dates, historique et conservation des fichiers Excel |
| excel_output.py | Output, portefeuille, copie des valeurs et conservation des formats du modèle XLSM |
| sector_deviations.py | Mapping EU/US vers 19 supersecteurs, ratios et fichier d'allocation ICB |
| research_secto.py | Choix des variables, consultations et backtests |

## Structure du modèle

model_secto.py contient six parties dans cet ordre :

1. SectorModel : choix du marché et exécution du calcul.
2. run_model : enchaînement des calculs d'un marché.
3. Sous-scores et piliers : mesures, rangs, Momentum et Volatility.
4. Macro et taux : contexte du mois.
5. Rangs et recommandations : scores agrégés, votes et choix des secteurs.
6. Exports : tableaux de scores en Excel.

Le fichier macro actualisé fournit le score et le cycle en U / AB pour EU et en Q / U pour US. Chaque couple de colonnes est conservé sous une série distincte dans Historique_data, sans modifier les anciennes observations.

Le schéma au début du fichier montre le parcours des données. actualiser_excel.py appelle le modèle puis excel_output.update_workbooks. La feuille Output est actualisée en valeurs seulement, sur une copie du XLSM existant ; les formats, les liaisons du Dashboard et la position des graphiques sont conservés. excel_output.py ajoute les nouveaux mois dans Reco_histo, rafraîchit les graphiques et exporte Output avec sa présentation, sans réappliquer NumberFormat ni reconstruire le Dashboard.

## Utilisation Python

```python
from model_secto import SectorModel
from research_secto import SectorResearch

model = SectorModel("EU")
research = SectorResearch(model)
results = model.run()
research.variables("Value")
research.composition("Value")
research.variable_table("Value")
research.recommendations()
```

`SectorModel.run()` utilise `production` par défaut. En `rebuild`, il initialise un historique vierge dans le dossier demandé ; en `replay`, il lit les séries gelées sans ouvrir les fichiers Excel source. `run_model` et tous les calculs de sous-scores, piliers, votes et portefeuille sont communs aux trois modes.

Chaque appel à `run()` recalcule le modèle. Les consultations `SectorResearch` utilisent le dernier résultat disponible. Pour une sélection de variables, créer `SectorResearch(model, variables=...)` avant de lancer le modèle.

run_production() calcule et exporte les tableaux Excel. Pour mettre à jour les deux XLSM et leur présentation, lancer `python actualiser_excel.py` ; les classeurs EU et US s'ouvrent automatiquement après une production réussie.

Les recommandations détaillées sont dans `*_recommandations_latest.xlsx`. `*_piliers_latest_available.xlsx` conserve la dernière date commune des piliers, qui peut être plus récente que celle des recommandations. `*_performance_portefeuille.xlsx` contient les rendements mensuels ; les courbes d’Output montrent les cumuls.

research.save_outputs() écrit aussi *_historique_modele.xlsx. Cet export complet n'est pas actualisé par le lancement mensuel standard.

Le mode se choisit dans la variable `mode` en fin de `actualiser_excel.py` ; les paramètres `output_dir` et `history_file` se règlent dans l’appel `main(...)`. Aucune option de ligne de commande n’est nécessaire.

## Flux selon le mode

- `production` : chargement de l'historique existant, lecture d'Excel et ajout des nouveaux mois par `freeze_history`, puis mise à jour des XLSM. Les observations anciennes ne sont pas remplacées.
- `rebuild` : initialisation d'un historique vide à la destination choisie, extraction de toutes les valeurs disponibles dans Excel, calcul EU puis US et export de l'historique construit.
- `replay` : lecture seule du snapshot `Historique_data`. Les fonctions de données recherchent chaque série par `series_key` et mois ; aucun appel à `open_workbooks`. Le modèle se recalcule avec la configuration sélectionnée.

`rebuild` et `replay` utilisent `save_all_outputs` et `save_excel` pour exporter les fichiers EU et US, ainsi que `historique_sectoriel_eu_us.xlsx`. Le dossier de production et les graphiques XLSM sont réservés au mode production. Une variable qui n'a jamais été enregistrée dans le snapshot ne peut pas être inventée par `replay` : il faut enrichir les entrées avant de la tester.

## Fichier des déviations sectorielles

La production appelle `sector_deviations.update_sector_deviations(models)` après `excel_output.update_workbooks`. Le fichier `reco_secto_facto.xlsx` contient `secto_us`, `secto_eu`, `Signals_US`, `Signals_EU` et `Mapping` ; les dates affichées sont les dates d'allocation (date du signal + 1 mois, au premier du mois). Les formules des feuilles secto utilisent les signaux (+1 / 0 / -1), le taux de déviation et les parts indiquées dans Mapping. Les formules indiquent directement la colonne de signal associée ; les noms de secteurs de Mapping servent de repère. Changer un nom seul ne modifie donc pas la correspondance dans les formules.

Correspondances actuelles dans `Mapping` : EU Fin = Banks 60 % / Financial Services 20 % / Insurance 20 %, EU Materials = Basic Resources 50 % / Chemicals 50 %, EU Indus = Industrial Goods & Services 80 % / Construction & Materials 20 %. US Fin et US Materials sont chacun divisés en trois parts égales. Les autres secteurs suivent une correspondance directe, ou restent à zéro s'ils n'existent pas dans l'univers du marché.

Les expositions positives et négatives sont équilibrées sur le plus petit nombre de recommandations des deux côtés. Une ligne dont seul l'un des côtés existe est neutre. Chaque date et chaque marché doit donner une somme nulle sur les 19 colonnes. Le fichier est sauvegardé après contrôle des totaux et l'ancien état est copié dans `output/avant_actualisation`.

Le chemin du classeur existant est `SECTOR_DEVIATION_EXCEL_FILE` dans `local_config.py` ; l'amplitude et les parts se règlent dans l'onglet `Mapping` du classeur. Le script actualise seulement les signaux et prolonge les formules pour les nouveaux mois, sans reconstruire la présentation. `rebuild` et `replay` ne réécrivent pas le fichier sectoriel officiel.

## Recherche et audit

research_secto.ipynb utilise SectorResearch. audit_modele_eu.ipynb et audit_modele_us.ipynb donnent une lecture détaillée du modèle. Les backtests de recherche sont distincts du portefeuille pondéré de production.

Les notebooks sont conservés sans leurs résultats d'exécution dans le dépôt. Les figures et exports Excel générés dans output/ peuvent être recréés.

Le Word comprend une synthèse pratique et des annexes de méthodologie et d'architecture Python. Les correspondances précises de cellules et les paramètres techniques restent dans local_config.py et les fonctions Python.
