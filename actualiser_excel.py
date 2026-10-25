# -*- coding: utf-8 -*-
"""Production mensuelle des classeurs EU et US.

ÉTAPES
1. Sauvegarder les fichiers avant modification.
2. Calculer les deux marchés avec model_secto.
3. Remplir et mettre en forme les classeurs avec excel_output.
4. Écrire le bilan des dates et des données manquantes.

FLUX
local_config + Excel -> data_utils -> model_secto
    -> excel_output -> Output + Reco_histo + Dashboard

Les données fournisseurs doivent être rafraîchies et enregistrées avant le lancement.
"""

from pathlib import Path
import shutil

import pandas as pd

from local_config import PROJECT_DIR, HISTORY_CONFIG, MACRO_EXCEL_FILE
from data_utils import normalize_metadata
from model_secto import SectorModel
from excel_output import update_workbooks


def backup_inputs(models, output_dir):
    """Garde une copie des deux classeurs et de l'historique avant le calcul."""
    file_paths = [m.excel_file for m in models]
    file_paths.append(PROJECT_DIR / HISTORY_CONFIG["file"])
    for file_path in file_paths:
        if (file_path.parent / ("~$" + file_path.name)).exists():
            raise RuntimeError(
                f"Fermer le fichier avant la mise à jour : {file_path.name}"
            )
    backup_dir = output_dir / "avant_actualisation"
    backup_dir.mkdir(exist_ok=True)
    for file_path in file_paths:
        if file_path.exists():
            copy_path = backup_dir / file_path.name
            shutil.copy2(file_path, copy_path)
            normalize_metadata(copy_path)


def model_summary(model, new_months):
    """Indique la date obtenue et les secteurs dont le score reste incomplet."""
    history = model.results["history"]
    date = history["date"].max()
    target_date = pd.Timestamp.today().normalize().replace(day=1) - pd.Timedelta(days=1)
    text = f"{model.universe} : date modèle {date:%d/%m/%Y}, {new_months} nouveau(x) mois archivé(s)."
    if date < target_date:
        text += f" Données en attente pour atteindre le {target_date:%d/%m/%Y}."
    latest = history[history["date"].eq(date)]
    missing = latest.loc[latest["score_global_macro"].isna(), "secteur"].tolist()
    if missing:
        text += " Scores agrégés indisponibles : " + ", ".join(missing) + "."
    return text


def main():
    output = PROJECT_DIR / "output"
    output.mkdir(exist_ok=True)
    models = [SectorModel("EU"), SectorModel("US")]

    # 1. Permettre un retour à l'état précédent en cas de problème.
    backup_inputs(models, output)
    macro = Path(MACRO_EXCEL_FILE)
    if macro.parent.resolve() == PROJECT_DIR:
        normalize_metadata(macro)

    # 2. Lire les entrées, calculer les scores et produire les exports courants.
    for model in models:
        print(f"Calcul du modèle {model.universe}...", flush=True)
        model.run_production(output_dir=output)

    # 3. Écrire Output, ajouter Reco_histo et actualiser les objets du Dashboard.
    added_months = update_workbooks(models, output)

    # 4. Garder un bilan court à lire avant de valider la production.
    rows = [model_summary(m, added_months[m.universe]) for m in models]
    text = "Mise à jour terminée à partir des données enregistrées.\n" + "\n".join(rows)
    (output / "actualisation.txt").write_text(text + "\n", encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
