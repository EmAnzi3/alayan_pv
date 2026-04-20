$repo = "C:\Users\anzillotti\OneDrive - CGT Edilizia S.p.a\Documenti\GitHub\alayan_pv"
$bat  = Join-Path $repo "scripts\aggiorna_sito_auto.bat"

if (!(Test-Path $repo)) {
    Write-Host "Repo non trovato: $repo"
    exit 1
}
if (!(Test-Path $bat)) {
    Write-Host "BAT non trovato: $bat"
    exit 1
}

Start-Process -FilePath $bat -WorkingDirectory $repo -Wait -NoNewWindow
