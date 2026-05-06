# Lokal Analytics

Det har ar en agar-guide for den lokala analytics-dashboarden.

Analytics syns inte i appens UI for anvandarna. Appen sparar i stallet anonyma events till filer som bara du behover lasa i dashboarden.

## Vad som sparas

Appen sparar bland annat:

- unika installationer via `distinct_id` och `install_id`
- appstarter via `app_started`
- appstangningar via `app_closed`
- feature-anvandning via `feature_usage`
- indatauppladdningar via `input_selected` och `input_loaded`

Viktigast for dig ar:

- `feature_usage` med `feature` och `action`
- `input_selected` med `file_type`, `source` och `extension`
- `app_closed` med `session_seconds`

Appen sparar inte filinnehall, ordernummer, kundnamn eller annan affarsdata i analytics.

## Var datan hamnar

Som standard sparas events i:

```text
%APPDATA%\allokering\analytics
```

Varje installation skriver till en egen `.jsonl`-fil. Det gor losningen enkel att lasa och mindre kanslig an en gemensam databasfil.

## Starta dashboarden

1. Installera dashboard-beroenden:

   ```bat
   python -m pip install -r requirements-analytics.txt
   ```

2. Starta dashboarden:

   ```bat
   start_analytics_dashboard.bat
   ```

3. Dashboarden oppnas i webblasaren lokalt pa din dator.

## Det du kan se direkt

Dashboarden visar bland annat:

- `Unika anvandare`
- `Filuppladdningar`
- `Feature-korningar`
- `Snitt oppen tid`
- popularaste funktioner
- mest uppladdade filtyper
- filuppladdningar per anvandare
- senaste events

## Om du vill samla flera anvandare utan server

Om du senare vill att flera anvandare ska skriva till samma stalle kan du satta en delad mapp i [app_info.py](/c:/Users/emikad/OneDrive%20-%20Dole%20Nordic%20AB/Skrivbordet/projects/allokering/app_info.py):

```py
ANALYTICS_LOCAL_STORAGE_DIR = r"\\\\server\\share\\allokering-analytics"
```

Du kan ocksa anvanda en synkad mapp, till exempel OneDrive eller SharePoint, men en vanlig natverksmapp ar oftast stabilast.

Nar den raden finns med i den version du distribuerar kommer klienterna att skriva events dit, och dashboarden kan lasa samma mapp.

## Bra att veta

- `install_id` ar anonymt men stabilt per installation.
- `session_seconds` visar hur lange appen varit oppen under en session.
- om dashboarden ar tom: oppna appen, ladda upp en fil, kor en funktion och stang appen igen
