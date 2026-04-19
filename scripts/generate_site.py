from __future__ import annotations
from pathlib import Path
import json
import sys

def load_config() -> dict:
    cfg_path = Path(__file__).with_name("config.json")
    with cfg_path.open("r", encoding="utf-8") as f:
        return json.load(f)

def validate_path(path_str: str, label: str) -> Path:
    path = Path(path_str)
    if not path.exists():
        raise FileNotFoundError(f"{label} non trovato: {path}")
    return path

def main() -> int:
    cfg = load_config()

    repo_root = validate_path(cfg["repo_root_dir"], "Cartella repo")
    docs_dir = Path(cfg["docs_dir"])
    agg_dir = validate_path(cfg["excel_aggregatore_dir"], "Cartella aggregatore")
    filiali_dir = validate_path(cfg["excel_filiali_dir"], "Cartella filiali")

    agg_file = agg_dir / cfg["aggregatore_filename"]
    if not agg_file.exists():
        raise FileNotFoundError(f"File aggregatore non trovato: {agg_file}")

    print("Percorsi validati correttamente.")
    print(f"Aggregatore: {agg_file}")
    print(f"Filiali: {filiali_dir}")
    print(f"Repo: {repo_root}")
    print(f"Docs: {docs_dir}")
    print()
    print("Qui va la tua versione definitiva di generate_site.py.")
    print("Mantieni config.json e aggiorna_sito.bat così come sono.")
    return 0

if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERRORE: {exc}")
        raise SystemExit(1)
