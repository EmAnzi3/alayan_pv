# Automazione aggiornamento sito Alayan PV

## File inclusi
- scripts/config.json
- scripts/aggiorna_sito_auto.bat
- scripts/run_task_scheduler.ps1
- scripts/task_scheduler_hourly.xml

## Cosa fa
1. legge gli Excel dal OneDrive
2. rigenera gli HTML nel repo locale
3. esegue `git add`
4. esegue commit solo se ci sono modifiche
5. fa `git push origin main`

## Prerequisiti
- `generate_site.py` già presente e funzionante in `scripts/`
- Git installato e disponibile nel PATH
- repo locale già collegato a GitHub
- OneDrive sincronizzato in locale

## Uso manuale
Doppio clic su:
`scripts\aggiorna_sito_auto.bat`

## Uso automatico con Utilità di pianificazione
### Metodo semplice
1. Apri `Utilità di pianificazione`
2. `Importa attività...`
3. seleziona `scripts\task_scheduler_hourly.xml`
4. controlla che l'utente sia quello giusto
5. salva

### Metodo manuale
1. Crea attività
2. Nome: `Alayan PV - Aggiornamento automatico`
3. Trigger: `Giornaliero`
4. Ripeti attività ogni: `1 ora`
5. Azione:
   - Programma/script: `powershell.exe`
   - Argomenti:
     `-NoProfile -ExecutionPolicy Bypass -File "C:\Users\anzillotti\OneDrive - CGT Edilizia S.p.a\Documenti\GitHub\alayan_pv\scripts\run_task_scheduler.ps1"`
   - Avvia in:
     `C:\Users\anzillotti\OneDrive - CGT Edilizia S.p.a\Documenti\GitHub\alayan_pv`

## Nota pratica
Se vuoi ridurre il rischio di aggiornamenti troppo frequenti:
- cambia la ripetizione da 1 ora a 3 ore
- oppure usa 3 trigger fissi: 08:00, 13:00, 18:00
