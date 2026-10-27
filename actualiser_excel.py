"""Production mensuelle des classeurs EU et US.

ÉTAPES
1. Sauvegarder les fichiers avant modification.
2. Calculer les deux marchés avec model_secto.
3. Coller les résultats dans les classeurs modèles sans refaire les formats.
4. Écrire le bilan des dates et des données manquantes.
5. Ouvrir les deux classeurs après une production réussie.

FLUX
local_config + Excel -> data_utils -> model_secto
    -> excel_output -> Output + Reco_histo + Dashboard

Les données fournisseurs doivent être rafraîchies et enregistrées avant le lancement.
"""

from pathlib import Path
import os
import shutil

import pandas as pd

from local_config import (
    EU_CONFIG, US_CONFIG, PROJECT_DIR, HISTORY_CONFIG, MACRO_EXCEL_FILE, SECTOR_DEVIATION_EXCEL_FILE,
)
from data_utils import normalize_metadata, save_table
from model_secto import SectorModel, save_all_outputs
from excel_output import compute_portfolio, save_excel, update_workbooks
from sector_deviations import update_sector_deviations


def backup_inputs(models, output_dir):
    """Garde une copie des deux classeurs et de l'historique avant le calcul."""
    file_paths = [m.excel_file for m in models]
    file_paths.extend([PROJECT_DIR / HISTORY_CONFIG["file"], Path(SECTOR_DEVIATION_EXCEL_FILE)])
    for file_path in file_paths:
        if (file_path.parent / ("~$" + file_path.name)).exists():
            raise RuntimeError(f"Fermer le fichier avant la mise à jour : {file_path.name}")
    backup_dir = output_dir / "avant_actualisation"
    backup_dir.mkdir(exist_ok=True)
    for file_path in file_paths:
        if file_path.exists():
            copy_path = backup_dir / file_path.name
            shutil.copy2(file_path, copy_path)
            normalize_metadata(copy_path)


def _missing_financial_pillars(model, date, cycle):
    """Renvoie la liste des piliers financiers absents pour le cycle et la date donnés."""
    required = {
        name for name, weight in model.config["macro_adjusted_weights"].items()
        if name != "Macro" and weight != 0
    }
    required.update(
        name for name, weight in model.config["regime_weights"][cycle].items()
        if weight != 0
    )
    return [
        name for name, values in model.results["pillars"].items()
        if name in required and pd.isna(values.at[date, "Fin"])
    ]


def _portfolio_summary(results, date):
    """Décrit le premier mois qui interrompt la série cumulée du portefeuille."""
    monthly, cumulative = compute_portfolio(results, date)
    missing_months = monthly.index[monthly["portfolio"].isna()]
    if len(missing_months) == 0:
        if pd.notna(cumulative["portfolio"].iloc[-1]):
            return " Courbe portefeuille : données cumulées disponibles."
        return ""

    first_month = missing_months[0]
    sectors = results["returns"].columns
    weight_history = results["benchmark_weights"].sort_index()
    weights = weight_history.loc[:first_month].iloc[-1].reindex(sectors)
    returns = results["returns"].loc[first_month].reindex(sectors)
    missing_weights = weights.index[weights.isna()].tolist()
    missing_returns = returns.index[returns.isna()].tolist()

    text = f" Courbe portefeuille interrompue dès le {first_month:%d/%m/%Y}."
    if missing_weights:
        text += " Poids Bench absents : " + ", ".join(missing_weights) + "."
    if missing_returns:
        text += " Rendements sectoriels absents : " + ", ".join(missing_returns) + "."
    return text


def model_summary(model, new_months):
    """Résume le mois produit, les scores absents et la disponibilité des performances."""
    history = model.results["history"]
    date = history["date"].max()
    target_date = pd.Timestamp.today().normalize().replace(day=1) - pd.Timedelta(days=1)
    text = f"{model.universe} : date modèle {date:%d/%m/%Y}, {new_months} nouveau(x) mois archivé(s)."
    if date < target_date:
        text += f" Données en attente pour atteindre le {target_date:%d/%m/%Y}."

    latest = history[history["date"].eq(date)]
    missing_scores = latest.loc[latest["score_global_macro"].isna(), "secteur"].tolist()
    if missing_scores:
        text += " Scores agrégés indisponibles : " + ", ".join(missing_scores) + "."
        if "Fin" in missing_scores:
            cycle = latest.loc[latest["secteur"].eq("Fin"), "cycle"].iloc[0]
            unavailable = _missing_financial_pillars(model, date, cycle)
            if unavailable:
                text += " Fin : pilier(s) absent(s) " + ", ".join(unavailable) + "."

    benchmark = model.results["benchmark_returns"]["benchmark"]
    if date not in benchmark.index or pd.isna(benchmark.at[date]):
        text += f" Rendement du benchmark {model.benchmark_name} absent au {date:%d/%m/%Y}."
    return text + _portfolio_summary(model.results, date)


def run_production():
    """Actualise les classeurs et les sorties mensuelles avec l'historique existant."""
    output = PROJECT_DIR / "output"
    output.mkdir(exist_ok=True)
    models = [SectorModel("EU"), SectorModel("US")]

    backup_inputs(models, output)
    macro = Path(MACRO_EXCEL_FILE)
    if macro.parent.resolve() == PROJECT_DIR:
        normalize_metadata(macro)

    for model in models:
        print(f"Calcul du modèle {model.universe}...", flush=True)
        model.run_production(output_dir=output)

    added_months = update_workbooks(models)
    deviation_file = update_sector_deviations(models)  # Repartir les recommandations sur les 19 supersecteurs.
    print(f"Déviations sectorielles : {deviation_file.name}")
    rows = [model_summary(model, added_months[model.universe]) for model in models]
    text = "Mise à jour terminée à partir des données enregistrées.\n" + "\n".join(rows)
    (output / "actualisation.txt").write_text(text + "\n", encoding="utf-8")
    print(text)


def run_history(mode, output_dir, history_file=None):
    """Exporte rebuild ou replay sans toucher aux XLSX ni à l'output mensuel."""
    if output_dir is None:
        raise ValueError("Indiquer output_dir pour rebuild ou replay.")
    destination = Path(output_dir).expanduser().resolve()
    if destination in (PROJECT_DIR.resolve(), (PROJECT_DIR / "output").resolve()):
        raise ValueError("Choisir un dossier de sortie distinct du projet et de output/.")
    if destination.exists() and any(destination.iterdir()):
        raise FileExistsError(f"Le dossier des résultats doit être vide : {destination}")

    snapshot = destination / HISTORY_CONFIG["file"]
    if mode == "rebuild":
        if history_file is not None:
            raise ValueError("rebuild repart d'un historique vide : ne pas fournir history_file.")
        if snapshot.exists():
            raise FileExistsError(f"Historique déjà présent : {snapshot}")
        history_file = snapshot
    else:
        history_file = Path(history_file or PROJECT_DIR / HISTORY_CONFIG["file"]).resolve()
        if not history_file.is_file():
            raise FileNotFoundError(f"Historique introuvable : {history_file}")
        if history_file.parent == destination:
            raise ValueError("En replay, le dossier des résultats doit être distinct du snapshot.")

    destination.mkdir(parents=True, exist_ok=True)
    histories = []
    for market in ("EU", "US"):
        model = SectorModel(market)
        print(f"{mode} : {market}...", flush=True)
        results = model.run(
            mode=mode,
            history_file=history_file,
            fresh_history=(mode == "rebuild" and market == "EU"),
        )

        save_all_outputs(results, destination, model.config)
        save_excel(results, destination, model.config)

        table = results["history"].copy()
        table.insert(0, "marche", market)
        histories.append(table)

    history = pd.concat(histories, ignore_index=True)
    save_table(history, destination / "historique_sectoriel_eu_us.xlsx")
    print(f"{mode} terminé : {destination}")


def main(mode="production", output_dir=None, history_file=None):
    if mode == "production":
        if output_dir is not None or history_file is not None:
            raise ValueError("production utilise les entrées et le dossier output du projet.")
        run_production()
    elif mode in ("rebuild", "replay"):
        run_history(mode, output_dir, history_file)
    else:
        raise ValueError(f"Mode inconnu : {mode}")


if __name__ == "__main__":
    mode = "production"  # "production", "rebuild" ou "replay"
    main(
        mode=mode,
        output_dir=None,    # Dossier de résultats pour rebuild et replay.
        history_file=None,  # Snapshot à lire en replay (None = data_history.xlsx).
    )
    if mode == "production":
        for config in (EU_CONFIG, US_CONFIG):
            os.startfile(config["excel_file"])
