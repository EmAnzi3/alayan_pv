# Alayan PV - generatore sito

## Come usarlo
1. Copia la cartella `scripts` dentro il repo locale `alayan_pv`.
2. Verifica che `docs/` esista nel repo.
3. Fai doppio clic su `scripts/aggiorna_sito.bat`.
4. Quando la generazione è finita, apri GitHub Desktop e fai commit + push.

## Cosa genera
- `docs/index.html`
- `docs/filiali/*.html`

## Regole già incluse
- conteggio agrivoltaico leggendo la parola `agrivoltaico` nelle note
- filtro area -> filiale dipendente nella overview
- Gantt chart in overview e nelle singole filiali
- correzione date sporche tipo `01 Jan 1970`
- link overview <-> filiali
