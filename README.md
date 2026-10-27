# Modèle sectoriel EU / US

Le projet calcule les scores et recommandations sectoriels EU et US. `python actualiser_excel.py` lance la production mensuelle et permet aussi de reconstruire ou de rejouer tout l’historique.

## Structure et pipeline

~~~mermaid
flowchart TD
    A["actualiser_excel.py"] --> B{"Mode"}
    B --> C["production : Excel + historique existant"]
    B --> D["rebuild : Excel + historique vierge"]
    B --> E["replay : historique seul"]
    C --> F["data_utils.py"]
    D --> F
    E --> F
    F --> G["model_secto.py"]
    G --> H["excel_output.py"]
    H --> I["Exports Excel"]
    C --> J["XLSM : Output, Dashboard, Reco_histo"]
~~~

| Fichier | Rôle |
| --- | --- |
| local_config.py | Chemins, secteurs, variables et paramètres du modèle |
| model_secto.py | Interface SectorModel et toutes les fonctions de calcul |
| data_utils.py | Lecture des données, historique mensuel et conservation des objets Excel |
| excel_output.py | Tableaux, performances et collage des valeurs dans les modèles Excel |
| sector_deviations.py | Répartition des recommandations sur les 19 supersecteurs ICB et export dédié |
| actualiser_excel.py | Sauvegarde, calcul EU et US, mise à jour Excel et bilan |
| research_secto.py et notebooks | Choix des variables, consultations, visualisations et backtests |

Le début de model_secto.py décrit ses six parties : interface, pipeline, sous-scores et piliers, macro et taux, rangs et recommandations, exports.
Les consultations et le choix des variables de recherche sont dans `SectorResearch` ; les variables se choisissent avant de lancer le modèle.

## Guide de production

Le guide pratique est [Guide_production_modele_sectoriel.docx](docs/Guide_production_modele_sectoriel.docx).

Le guide présente une synthèse opérationnelle en quatre pages et trois annexes : liens entre les fichiers Python, méthode détaillée des facteurs, agrégation des scores et benchmarks. Les adresses Excel restent dans le code.

## Trois modes d'exécution

| Mode | Données utilisées | Destination |
| --- | --- | --- |
| `production` (défaut) | Historique figé existant et nouveaux mois enregistrés dans Excel | `output/` et XLSM EU/US |
| `rebuild` | Toute l'histoire accessible dans les Excel actuels, sans historique initial | Dossier choisi, avec un nouveau `data_history.xlsx` |
| `replay` | `data_history.xlsx` uniquement ; aucun fichier source Excel ouvert | Dossier choisi, sans modifier l'historique |

Choisir `mode` et renseigner les paramètres de `main(...)` à la fin de `actualiser_excel.py`, puis lancer `python actualiser_excel.py`.

~~~python
if __name__ == "__main__":
    mode = "production"  # "production", "rebuild" ou "replay"
    main(
        mode=mode,
        output_dir=None,    # Dossier choisi pour rebuild/replay.
        history_file=None,  # Historique à lire en replay.
    )
    if mode == "production":
        for config in (EU_CONFIG, US_CONFIG):
            os.startfile(config["excel_file"])
~~~

Pour reconstruire l'historique, choisir `mode = "rebuild"` et `output_dir="output/rebuild_octobre"` dans `main(...)`.
Pour le rejouer sans ouvrir Excel, choisir `mode = "replay"`, `output_dir="output/replay_test"` et `history_file="output/rebuild_octobre/data_history.xlsx"`.

`rebuild` et `replay` demandent un dossier de sortie vide. Si `history_file=None` en replay, le modèle lit le `data_history.xlsx` du projet.

Les modes rebuild et replay écrivent les historiques de scores EU/US, les rendements et les grilles Output dans leur propre dossier, sans mettre à jour les XLSM. Ils partagent les mêmes fonctions de calcul que la production. Le résultat reconstruit reflète les valeurs historiques actuellement sauvegardées dans Excel et peut donc différer d'un historique de premières observations.

## Production mensuelle

1. Choisir un mois terminé et conserver une copie des entrées.
2. Rafraîchir les données FactSet / Bloomberg dans Excel, puis enregistrer les valeurs.
3. Mettre à jour le fichier macro, le taux US 10 ans de Cycle macro et les poids de Bench.
4. Fermer les deux classeurs sectoriels, le fichier macro, data_history.xlsx et reco_secto_facto.xlsx.
5. Depuis le dossier du projet, lancer `python actualiser_excel.py` ou exécuter ce fichier Python dans VS Code.
6. Lire output/actualisation.txt et contrôler Output, Dashboard, Reco_histo ainsi que reco_secto_facto.xlsx.
7. Conserver et diffuser la version validée.

Le lancement calcule les deux modèles et remplace les valeurs d'Output dans des copies des classeurs d'origine. Après une production réussie, les deux classeurs EU/US s'ouvrent automatiquement. Les formats, les formules du Dashboard et la disposition des graphiques sont ceux du modèle Excel et ne sont pas reconstruits par Python. Les graphiques sont rafraîchis et Reco_histo ne reçoit que les nouveaux mois. Aucun collage manuel n'est nécessaire.

La collecte fournisseur n'est pas déclenchée par ce lancement. Le programme ne filtre pas automatiquement les observations du mois en cours. Les premières observations enregistrées dans data_history.xlsx (feuille Historique_data) restent figées.

## Écarts sectoriels ICB

En production, le classeur existant `reco_secto_facto.xlsx` est actualisé à la racine du projet. Les onglets `secto_us` et `secto_eu` reproduisent les 19 supersecteurs des tableaux d'allocation, avec une date au premier jour du mois suivant le signal. `Signals_US`, `Signals_EU` et `Mapping` détaillent les décisions initiales et leur répartition.

La déviation par recommandation est actuellement de 5 % dans la cellule modifiable `Mapping!I2` du classeur. En Europe, Fin se répartit entre Banks / Financial Services / Insurance (60 / 20 / 20 %), Materials entre Basic Resources / Chemicals (50 / 50 %) et Indus entre Industrial Goods & Services / Construction & Materials (80 / 20 %). Aux États-Unis, Fin et Materials se répartissent également sur trois supersecteurs chacun. Les rubriques absentes du modèle restent à zéro.

Le total de chaque mois et de chaque marché est nul : si le nombre de Positive et Negative diffère, le budget commun est réparti entre les secteurs de chaque côté, sans dépasser l’amplitude définie dans `Mapping!I2`. S'il n'existe qu'un seul côté, aucune déviation n'est affectée. Les cellules du tableau sont calculées par des formules Excel. Le script met à jour les signaux et prolonge les formules existantes pour les nouveaux mois ; la présentation et le mapping restent dans le classeur.

Le chemin du fichier se change dans `local_config.py` avec `SECTOR_DEVIATION_EXCEL_FILE`. Son emplacement est indépendant des fichiers de sources EU/US et du dossier `output/`.

## Installation et sources

Installer les dépendances avec :

```bash
pip install -r requirements.txt
```

L'écriture des classeurs nécessite Excel pour Windows. Les fichiers source sont, par défaut, à côté du code :

- Score_Sectoriel_EU_NEW.xlsm
- Score_Sectoriel_US_NEW.xlsm
- maj cycle macro2.xlsx
- data_history.xlsx

La configuration macro suit les nouvelles colonnes : EU, score U et cycle AB ; US, score Q et cycle U. Les périodes de ces nouveaux signaux sont archivées séparément des anciens.

Les chemins et les paramètres se règlent dans local_config.py. Pour un essai Python ponctuel, SectorModel accepte aussi les arguments excel_file et macro_file.

## Résultats et sauvegarde

Les deux XLSM sont mis à jour à partir de leur propre mise en page ; leurs feuilles Output, Dashboard, Reco_histo et leurs liens doivent donc être déjà préparés dans le modèle. Les exports Excel (recommandations, performances et copie de la feuille Output avec ses formats) et le bilan sont dans output/. La dernière sauvegarde automatique est dans output/avant_actualisation/ ; elle ne remplace pas une archive mensuelle. Les copies de travail Excel sont créées dans le dossier temporaire du système et supprimées à la fin.

La synthèse d'Output affiche le mois du signal, les dernières données de piliers et le mois d'allocation suivant. Les scores indisponibles sont indiqués par n.a.

Le rendement du benchmark US provient de WIUSAU Index dans Returns_EQ. Le libellé historique S&P500 des graphiques a été remplacé ; les poids Bench US restent identifiés SP500_bench. Le rapport signale les rendements de benchmark manquants, les piliers qui bloquent un score Fin et la première interruption du cumul portefeuille. Si les valeurs du portefeuille existent dans Output mais pas dans le Dashboard, examiner les séries de graphique `Courbe_Ptf` et `Annee_Ptf`.

## Contenu du dépôt

Conserver le code, les notebooks sans sorties, les quatre fichiers Excel source, requirements.txt et le guide Word. data_history.xlsx fait partie des données à conserver : ses premières observations ne se reconstruisent pas à l'identique à partir de données révisées.

Le dossier output/, les caches Python, les environnements locaux et les fichiers temporaires Excel sont exclus par .gitignore. Les exports de recherche peuvent être supprimés puis recréés. Garder une archive externe des versions mensuelles validées.

Voir DOCUMENTATION.md pour l'index technique.
