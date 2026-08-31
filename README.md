# Aggiornamento sito Alayan PV

## File principali
- `scripts/config.json`
- `scripts/generate_site.py`
- `scripts/aggiorna_sito.bat`
- `scripts/check_before_publish.ps1`

## Cosa fa

Il sito viene aggiornato manualmente.

1. `aggiorna_sito.bat` legge gli Excel presenti nelle cartelle OneDrive configurate.
2. `generate_site.py` rigenera gli HTML nella cartella `docs`.
3. Il sito viene controllato localmente.
4. Commit e push vengono eseguiti manualmente con GitHub Desktop.

## Uso

Eseguire:

`scripts\aggiorna_sito.bat`

Al termine dello script, verificare il risultato prima di eseguire commit e push.

<!-- MAINTENANCE-STANDARD:START -->
## Manutenzione repository

- Stato operativo: `CURRENT_STATE.md`
- Istruzioni per ChatGPT/Codex: `AGENTS.md`
- Storico modifiche: `CHANGELOG.md`
- Controllo pre-pubblicazione: `.\scripts\check_before_publish.ps1`

Comando consigliato prima del commit:

`powershell
.\scripts\check_before_publish.ps1
git status
git diff --check
`
<!-- MAINTENANCE-STANDARD:END -->
