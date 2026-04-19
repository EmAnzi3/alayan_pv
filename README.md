# Alayan PV — GitHub Pages package

Pacchetto pronto da caricare nel repository `EmAnzi3/alayan_pv`.

## Struttura
- `docs/` → sito statico da pubblicare con GitHub Pages
- `input_xlsx/` → file Excel sorgente caricati come riferimento
- `templates_base/` → i due file HTML base corretti manualmente (`index.html` e `catania.html`)
- `.nojekyll` → evita trasformazioni Jekyll inutili su GitHub Pages

## Pubblicazione su GitHub Pages
1. Carica il contenuto di questo pacchetto nel repo.
2. Vai su **Settings → Pages**.
3. In **Build and deployment**, seleziona **Deploy from a branch**.
4. Branch: `main`
5. Folder: `/docs`
6. Salva.

## Aggiornamento rapido
Quando vorrai sostituire il sito pubblicato:
1. aggiorna i file in `docs/`
2. fai commit e push
3. GitHub Pages ripubblica il sito

## Nota
Questo pacchetto contiene la versione statica più recente del sito con:
- overview nazionale
- pagine filiale
- gantt chart globale e per filiale
- filtri attivi
- correzione date "Ultimo aggiornamento"
