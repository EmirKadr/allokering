# Klassificering V2 Paritetsmatris

Källa: `C:\artikelplacering\Artikelplacering\classifier.py`  
Mål: Chromium-baserad implementation i `C:\allokering\webapp`.

## Skärmar / vyer

| Område | Status | Kommentar |
|---|---|---|
| Name | Implementerad | `test_name` i v2-startflöde |
| Categories | Implementerad | `name|description`, ny kategori under körning |
| Source | Implementerad | `auto/images/item_attribute` |
| AI Settings | Implementerad | `api_url`, `model`, `api_key` |
| Filter | Delvis | Grunddataflöde finns, full filter-paritet kvar |
| Classify | Implementerad | Klassificera, hoppa över, gå tillbaka |
| AI Job Kanban | Implementerad (grund) | Kolumner, drag/drop, Ctrl-multiselect, snabb reclass |
| Done | Implementerad | Export ZIP/Excel, avsluta i förtid |

## Data-/domänobjekt

| Objekt | Status | Kommentar |
|---|---|---|
| `test_name` | Implementerad | |
| `categories` | Implementerad | inkl. `description` |
| `csv_data` | Implementerad | URL + lokal bild |
| `categorized` | Implementerad | |
| `results` | Implementerad | `reason` + source |
| `cat_knowledge` | Implementerad | |
| `cat_example_articles` | Implementerad | |

## API-v2

| Endpointgrupp | Status |
|---|---|
| Config + datafiler | Implementerad |
| Session start/state/delete | Implementerad |
| Actions (classify/skip/back/finish/add/reclassify/knowledge) | Implementerad |
| AI controls (start/pause/resume/stop/start-step2/analyze-all) | Implementerad |
| Import/export ZIP | Implementerad |
| Import/export Excel | Implementerad |
| Eventstream/loggar | Implementerad |

## Tangentflöden

| Funktion | Status |
|---|---|
| `1-9`, `0` kategori | Implementerad |
| Vänsterpil = tillbaka | Implementerad |
| Högerpil = hoppa över | Implementerad |

## Kända gap mot full 1:1

1. Full visuell paritet med PyQt-layout är inte pixel-identisk.
2. Fullständig prompt/logikparitet för steg 1/2 i AI är delvis förenklad i v2-backend.
3. Kanban contextmenu/reclassify-dialog är webbanpassad och enklare än originaldialogen.
4. Full regressionsjämförelse mot exakt original-output på alla dataset återstår.

