# Releaseprocess

Den här filen beskriver hur vi publicerar en ny version av Allokering så att
installeraren hamnar på GitHub Releases och användare kan uppdatera appen.

## Viktig regel

Skapa inte en ny release för varje ändring.

Vanliga kodändringar ska normalt bara commitas och pushas till `main`. En release
ska bara göras när Emir uttryckligen ber om det, till exempel:

- "gör en release"
- "släpp version 12.2.0"
- "tagga och publicera ny version"
- "nu ska kollegan få en uppdatering"

## Vad som händer vid release

När en tagg som `v12.2.0` pushas startar GitHub Actions-workflowen
`.github/workflows/windows-release.yml`.

Workflowen bygger:

- `Allokering-12.2.0-win64.zip`
- `Allokering-12.2.0-Setup.exe`

Vid tagg-push laddas filerna även upp på GitHub Release. Appens updater läser
senaste GitHub Release och letar efter `Setup.exe`. Om versionen där är högre än
användarens installerade version får användaren frågan att uppdatera.

## Steg för ny release

Byt ut `12.2.0` mot versionsnumret som ska släppas.

1. Kontrollera att alla ändringar är klara.
2. Höj versionsnumret i `app_info.py`:

   ```py
   APP_VERSION = "12.2.0"
   APP_VERSION_DISPLAY = "12.2"
   ```

3. Bygg och smoke-testa installeraren lokalt:

   ```bat
   build_windows.bat
   ```

   Förväntade filer:

   ```text
   release\Allokering-12.2.0-win64.zip
   release\Allokering-12.2.0-Setup.exe
   ```

4. Committa versionshöjningen och eventuella ändringar:

   ```bat
   git status --ignore-submodules=all
   git add .
   git commit -m "Release 12.2.0"
   git push
   ```

5. Skapa och pusha release-taggen:

   ```bat
   git tag v12.2.0
   git push origin v12.2.0
   ```

6. Kontrollera GitHub Actions och GitHub Release:

   - Actions ska bli grön.
   - Releasen `v12.2.0` ska ha `Setup.exe` och zip som assets.

## Efter release

Användare får uppdateringen genom:

- automatisk kontroll vid appstart
- eller `Hjälp -> Sök efter uppdateringar`

Installeraren är per-user och kräver inte administratörsrättigheter. När
användaren godkänner uppdateringen laddar appen ner `Setup.exe`, startar den tyst
och stänger appen medan uppdateringen installeras.

## Gör inte detta

- Skapa inte tagg/release för små mellanändringar.
- Pusha inte en tagg utan att först höja `APP_VERSION`.
- Återanvänd inte samma versionsnummer för olika installerare.
- Force-pusha inte release-taggar utan att uttryckligen diskutera det först.
- Ladda inte upp en lokal installer manuellt om den inte matchar exakt version i
  `app_info.py`.
