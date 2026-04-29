# Bygga Windows-app

Det här projektet byggs med PyInstaller till en fristående Windows-appmapp.
Builden körs i en temporär lokal mapp för att undvika låsningar från OneDrive.

## Snabbbygge

Kör i CMD från projektroten:

```bat
build_windows.bat
```

Resultat:

- `release\Allokering\Allokering.exe`
- `release\Allokering-12.1.0-win64.zip`
- `release\Allokering-12.1.0-Setup.exe` om Inno Setup 6 finns installerat

Zip-filen kan delas till användare utan Python installerat. Den innehåller appen,
`Installera Allokering.bat`, avinstallation och en kort användar-README.

Om Inno Setup 6 finns installerat skapas även en riktig `Setup.exe`.

## Installer

Inno Setup-mallen finns i:

```text
packaging\windows\Allokering.iss
```

När Inno Setup 6 finns installerat bygger `build_windows.bat` en riktig `Setup.exe`
från den redan byggda `release\Allokering`-mappen. Det går också att köra:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File packaging\windows\build_setup.ps1
```

Installeraren är per-user och kräver inte administratörsrättigheter. Den installeras
i användarens `%LOCALAPPDATA%\Allokering`, skapar genvägar och registrerar
avinstallation.

## Uppdateringar

Appen har `Hjälp -> Sök efter uppdateringar` och gör även en tyst kontroll vid
start. Den läser senaste GitHub Release från `EmirKadr/allokering`, letar efter en
asset som slutar på `Setup.exe`, laddar ner den och startar installeraren.
Eftersom installeraren är per-user behövs inga admin-rättigheter vid uppdatering.
När användaren godkänner en uppdatering körs installeraren tyst.

Versionsnumret finns i `app_info.py`. Höj `APP_VERSION`, skapa en tagg som
`v12.2.0` och pusha taggen för att skapa en ny release-build.

Se `RELEASE.md` för hela releaseprocessen.

## GitHub artifact

Workflowen `.github/workflows/windows-release.yml` bygger Windows-paketet manuellt
via GitHub Actions (`workflow_dispatch`) eller när en tagg som `v12.1.0` pushas.
Den laddar upp zippen och `Setup.exe` som artifacts. Vid tagg-push laddas samma
filer också upp på GitHub Release så appens updater kan hitta dem.
