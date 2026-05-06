# Allokering filkunskap

Det har dokumentet samlar filkunskapen bakom CLI:t for `kor allokering` och for framtida Stigamo-verktyg.

## Kallstatus

- `Observerat i fil`: headers och sample-data hittade lokalt i exportfilerna.
- `Kodifierat i logik`: kolumnalias och normalisering som finns i CLI-koden eller originalkoden.
- `Wiki-beskrivet`: vy, omrade, datatyp och flodesroll som redan dokumenterats i Stigamo-wikin.
- `Osakert/infererat`: sadant som inte uttryckligen stod i fil eller wiki men drogs som forsiktig slutsats.

## Filmatris

| Family ID | Vy eller tabell | Omrade | Kalla | Datatyp | Flodessteg | CLI v1 |
| --- | --- | --- | --- | --- | --- | --- |
| `v_ask_customer_order_details_all` | Detalj Kundorder (Alla) | 3. Orderhantering | ASK | Orderrader pa detaljniva | Order och plock | Ja |
| `v_ask_article_buffertpallet` | Buffertpall | 2. Artikelhantering | ASK | Buffertpallar och pallinnehall | Buffert/lager | Ja |
| `v_ask_item_summary_stock_automation` | Saldo Inkl. Automation | 2. Artikelhantering | ASK | Saldo per artikel | Buffert/lager och automation | Ja |
| `item_option` | Item Option | Underhallstabeller | ASK | Artikelregler och artikeloptionsdata | Masterdata/artikelregler | Ja |
| `v_ask_order_overview` | Orderöversikt | 3. Orderhantering | ASK | Orderhuvud och oversiktsdata | Orderhuvud och dispatchforberedelse | Nej |
| `v_ask_booking_putaway` | Ej Inlagrade Artiklar | 1. Varumottagning | ASK | Mottaget men ej inlagrat gods | Inleverans och ej inlagrat | Nej |
| `v_ask_receive_log` | Varumottagningslogg | Loggar | ASK | Mottagningshistorik | Inleverans | Nej |
| `v_ask_trans_log` | Translogg | Loggar | ASK | Interna lagerforflyttningar | Interna rorelser | Nej |
| `v_ask_pick_log_full` | Plocklogg Full | Loggar | ASK | Plockhandelser | Order och plock | Nej |
| `v_ask_correct_log` | Saldojusteringar & Inventeringsavvikelser | Loggar | ASK | Korrigerings- och inventeringslogg | Korrigeringar | Nej |
| `v_ask_assignment_move` | Palluppdrag | 2. Artikelhantering | ASK | Arbetsuppdrag for pallflyttar | Interna rorelser och pafyllning | Nej |
| `v_ask_dispatch_pallet` | Dispatchpallar | Underhallstabeller | ASK | Utgaende pallar kopplade till leveranser | Dispatch/utleverans | Nej |
| `dimension` | Dimensioner | Underhallstabeller | ASK | Dimensionsregister | Masterdata/plats- och pallregler | Nej |
| `item` | Artiklar (Item) | Underhallstabeller | ASK | Artikelmasterdata | Masterdata/artikelregler | Nej |
| `location` | Lagerplatser | Underhallstabeller | ASK | Lagerplatsmaster | Masterdata/platsregler | Nej |
| `main_category` | Huvudkategori | Register | ASK | Kategorimappning | Masterdata/kategori | Nej |
| `pallet_type` | Palltyp | Register | ASK | Palltypsregister | Masterdata/plats- och pallregler | Nej |
| `v_ask_order_log` | Orderlogg | Loggar | ASK | Orderhuvud och orderhistorik | Orderhuvud och historik | Nej |
| `v_ask_pick_location_log` | Plockplatslogg | Loggar | ASK | Historik over plockplatsbyten | Platsforandringar | Nej |
| `asw_order` | ASW Order | Integration | ASW | Order- och radinformation fran ASW | Order och plock | Nej |
| `dblog_pick_log` | Arkiv Plocklogg | X. Arkiv | ASK | Arkiverad plocklogg | Order och plock / arkiv | Nej |
| `prognos_excel` | Kundprognos | Kundinput | Kund-Excel | Planerings- och behovsprognos | Prognos och planering | Nej |
| `campaign_excel` | Kampanjvolymer | Kundinput | Kund-Excel | Kampanj- och volymplanering | Prognos och planering | Nej |

## Excel-regler

### Prognos Excel

Originalregeln i `read_prognos_xlsx` ar:

1. Ta bort raderna 0, 1 och 3 om de finns.
2. Ta bort forsta kolumnen.
3. Anvand forsta kvarvarande rad som headers.
4. Matcha kandidater for `Artikelnummer`, `Beskrivning`, `Antal styck`, `Antal rader` och `Antal butiker`.

### Kampanj Excel

Originalregeln i `read_campaign_xlsx` ar:

1. Ta bort rad 4 om den finns.
2. Ta bort raderna 0, 1 och 2.
3. Behall bara kolumner upp till index 6.
4. Droppa kolumnerna 5, 4, 3, 1 och 0 om de finns.
5. Normalisera till `Artikelnummer` och `Antal styck`.

## Flodeskarta

`inleverans -> ej inlagrat -> buffert/lager -> order och plock -> interna rorelser -> korrigeringar -> dispatch/utleverans -> prognos framat`

## v_ask_customer_order_details_all

**Namn:** Detalj Kundorder (Alla)

**Kort svar:** Kundorder pa radniva: artikel, behov, status, zon och plockkontext.

**Kalla:** ASK

**Vy/tabell:** Detalj Kundorder (Alla)

**ASK-omrade:** 3. Orderhantering

**Datatyp:** Orderrader pa detaljniva

**Flodessteg:** Order och plock

**Anvands i CLI v1:** Ja

### Observerat i fil

`Status`, `Beskrivning`, `Struktur`, `Kund`, `Kund.1`, `Artikel`, `Artikel.1`, `Plockplats`, `Plock`, `Plockat`, `Beställt`, `Timestamp`, `Diff`, `Zon`, `Bin Typ`, `Användare`, `Rad`, `Order nr`, `Rel`, `Batch`, `Orderstart`, `Order`, `Circa`, `Lager`, `Bolag`, `Meddelande`, `Robot`, `Robot Artikel`, `Pack klass`, `Kopplat inköp`, `Inköpsrad`, `Utgångsdatum`, `Anskaffningsmetod`, `Är plockad`, `RecordId`, `Orderdatum`, `Index num`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Artikel` | Artikeln som ska allokeras eller plockas. |
| `Beställt` | Efterfragat antal som allokeringsmotorn arbetar mot. |
| `Status` | Radstatus; 35 filtreras bort i allokeringen. |
| `Zon` | Original zonkod som senare reklassas till CLI-resultat. |
| `Order nr` | Order-id som bevaras i utdata och parity-jamforelse. |

### Kodifierat i logik

- `artikel`: `artikel`, `artikelnummer`, `sku`, `article`, `artnr`, `art.nr`
- `qty`: `beställt`, `antal`, `qty`, `quantity`, `bestalld`, `order qty`
- `status`: `status`, `radstatus`, `orderstatus`, `state`
- `ordid`: `ordernr`, `order nr`, `order number`, `kund`, `kundnr`
- `radid`: `radnr`, `rad nr`, `line id`, `rad`, `struktur`, `radsnr`

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_customer_order_details_all-20260317145125.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\v_ask_customer_order_details_all-20260327111808.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## v_ask_article_buffertpallet

**Namn:** Buffertpall

**Kort svar:** Aktuella buffertpallar, deras plats, kvantitet, status, tid och palltyp.

**Kalla:** ASK

**Vy/tabell:** Buffertpall

**ASK-omrade:** 2. Artikelhantering

**Datatyp:** Buffertpallar och pallinnehall

**Flodessteg:** Buffert/lager

**Anvands i CLI v1:** Ja

### Observerat i fil

`Lagerplats`, `Pallid`, `Artikel`, `Beskrivning`, `Palltyp`, `Antal`, `Status`, `Vikt`, `Timestamp`, `Detalj`, `Kontrolind`, `Antal per lav`, `Pack klass`, `Datum/tid`, `Batch`, `Utgångsdatum`, `Lager`, `Bolag`, `Beskrivning.1`, `Rfid`, `Mixpallsnummer`, `Reservation`, `Platstyp`, `Produktionsdatum`, `Inköpsnr`, `Robot ind`, `Krangång`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Lagerplats` | Styr platsregler och skiljer AUTOSTORE mot ovriga pallar. |
| `Pallid` | Kallan som skrivs ut i allokeringsresultatet. |
| `Antal` | Tillganglig kvantitet pa pallen. |
| `Status` | Maste vara 29, 30 eller 32 for att anvandas. |
| `Datum/tid` | Anvands for FIFO-sortering inom artikel. |

### Kodifierat i logik

- `artikel`: `artikel`, `article`, `artnr`, `art.nr`, `artikelnummer`
- `qty`: `antal`, `qty`, `quantity`, `pallantal`, `colli`, `units`
- `loc`: `lagerplats`, `plats`, `location`, `bin`, `hyllplats`
- `dt`: `datum/tid`, `datum`, `mottagen`, `received`, `inleverans`, `inleveransdatum`, `timestamp`, `arrival`
- `id`: `pallid`, `pall id`, `id`, `sscc`, `etikett`, `batch`, `lpn`
- `status`: `status`, `pallstatus`, `state`

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_article_buffertpallet-20260317145136.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\v_ask_article_buffertpallet-20260327111741.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: place-analysis-data
- Observerad sample-fil

## v_ask_item_summary_stock_automation

**Namn:** Saldo Inkl. Automation

**Kort svar:** Samlad saldobild per artikel med plock, buffert, automation, kran och autoplock.

**Kalla:** ASK

**Vy/tabell:** Saldo Inkl. Automation

**ASK-omrade:** 2. Artikelhantering

**Datatyp:** Saldo per artikel

**Flodessteg:** Buffert/lager och automation

**Anvands i CLI v1:** Ja

### Observerat i fil

`Artikel`, `Beskrivning`, `Vikt netto`, `Plockplats`, `Plocksaldo`, `Kundorder`, `Diff`, `Buffertsaldo`, `Saldo automation`, `Automation diff`, `Saldo kran`, `Kran diff`, `Saldo autoplock`, `Autoplock diff`, `Utbeställt`, `Inlagring`, `Ankommande`, `Pågående plock`, `För och under plock`, `Saldo`, `Diff2`, `Externt`, `Lager`, `Bolag`, `Batchartikel`, `Klassificering`, `Robot`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Artikel` | Nyckel per artikel i saldonormaliseringen. |
| `Plockplats` | Forsta icke-tomma plockplatsen sparas per artikel. |
| `Plocksaldo` | Summeras per artikel i CLI-normaliseringen. |
| `Saldo automation` | Visar lagersaldo i automationen men anvands inte direkt i v1. |
| `Saldo autoplock` | Viktig tolkningssignal i nuvarande WMS-underlag. |

### Kodifierat i logik

- `artikel`: `artikel`, `artnr`, `art.nr`, `artikelnummer`, `sku`, `article`
- `plocksaldo`: `plocksaldo`, `plock saldo`, `plock-saldo`, `saldo`, `pick saldo`, `pick qty`, `tillgängligt plock`, `tillgangligt plock`, `available pick`, `plock`
- `plockplats`: `plockplats`, `huvudplock`, `mainpick`, `hyllplats`, `bin`, `location`, `lagerplats`

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_item_summary_stock_automation-20260317145351.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\v_ask_item_summary_stock_automation-20260327111753.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## item_option

**Namn:** Item Option

**Kort svar:** Artikelvisa regler som plockzon, robotplock och ej staplingsbar.

**Kalla:** ASK

**Vy/tabell:** Item Option

**ASK-omrade:** Underhallstabeller

**Datatyp:** Artikelregler och artikeloptionsdata

**Flodessteg:** Masterdata/artikelregler

**Anvands i CLI v1:** Ja

### Observerat i fil

`Artikel`, `Lager`, `Bolag`, `Nollinventering`, `Hantera saldo`, `Timestamp`, `Klassificering`, `Skapad`, `Plockzon`, `Vikt tolerans`, `Zon klass`, `Tillåt uppdrag`, `Ignorera EAN`, `Splitta helpalls gen. med`, `Automatiserat robotplock`, `Ej staplingsbar`, `Alltid farligt gods`, `Helpalls avvikelse %`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Artikel` | Nyckel som matchas mot allokeringsresultatet. |
| `Plockzon` | Visar affarsregel kring var artikeln normalt plockas. |
| `Automatiserat robotplock` | Signal om robotkoppling i kallsystemet. |
| `Ej staplingsbar` | Skrivs in i resultatet och paverkar pallplatsrapporten. |

### Kodifierat i logik

- `artikel`: `artikel`, `artikelnummer`, `sku`, `article`, `artnr`, `art.nr`
- `staplingsbar`: `staplingsbar`, `staplings bar`, `staplbar`, `stackable`, `ej staplingsbar`, `ejstaplingsbar`, `ej_staplingsbar`, `non stackable`

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\item_option-20260317145203.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## v_ask_order_overview

**Namn:** Orderöversikt

**Kort svar:** Orderhuvud med avganger, kund, vikt, rader, transportor och sandningsnummer.

**Kalla:** ASK

**Vy/tabell:** Orderöversikt

**ASK-omrade:** 3. Orderhantering

**Datatyp:** Orderhuvud och oversiktsdata

**Flodessteg:** Orderhuvud och dispatchforberedelse

**Anvands i CLI v1:** Nej

### Observerat i fil

`Ordernr`, `Status`, `Land`, `Struktur`, `Trans nr`, `Transportör`, `Prio`, `Starttid`, `Orderdatum`, `Leveransdatum`, `Ursprungsdatum`, `Laststarttid`, `Avgångstid`, `Yta`, `Användare`, `Timestamp`, `Rader`, `Antal`, `Zon`, `Multi`, `SPC`, `Robot info`, `Ordertyp`, `Lager`, `Volym`, `Vikt`, `Kund nr`, `Kund`, `Avgångsnr.`, `Multi index`, `Vagn`, `Sändningsnr`, `Produkt`, `TransportProdukt`, `Avgång`, `Bolag`, `Alt adress`, `Kund Adr`, `Butiks nr`, `Kund ref`, `Inköpsnr`, `Meddelande`, `Orderflagga`, `Kollisnitt`, `Brand`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Ordernr` | Orderhuvudets id. |
| `Orderdatum` | Skapar tidskontext per orderhuvud. |
| `Sändningsnr` | Kopplar order till utgaende leverans. |
| `Transportör` | Visar vald transportor pa huvudniva. |
| `Ordertyp` | Viktig affarsklassificering i flera kontroller. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_order_overview-20260317145114.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## v_ask_booking_putaway

**Namn:** Ej Inlagrade Artiklar

**Kort svar:** Pall- och inkopsrader som ar mottagna men fortfarande saknar slutlig inlagring.

**Kalla:** ASK

**Vy/tabell:** Ej Inlagrade Artiklar

**ASK-omrade:** 1. Varumottagning

**Datatyp:** Mottaget men ej inlagrat gods

**Flodessteg:** Inleverans och ej inlagrat

**Anvands i CLI v1:** Nej

### Observerat i fil

`Status`, `Prioritet`, `Pall nr`, `Batch nr`, `Artikel`, `Artikel.1`, `Område`, `Ändrad`, `Antal`, `Vikt`, `Användare`, `Kö`, `Utgång`, `SSCC`, `Bolag`, `Lager`, `Antal per lav`, `Detalj`, `Kontrolind`, `Viktartikel`, `Palltyp`, `Pack klass`, `Mix Pall`, `Inköpsnr`, `Crossdock Ind`, `Produktionsdatum`, `Robot Ind`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Pall nr` | Pallens identitet innan inlagring. |
| `Inköpsnr` | Inkopskoppling till inkommande flode. |
| `Artikel` | Vilken artikel pallen innehaller. |
| `Status` | Visar var raden befinner sig i inlagringsprocessen. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_booking_putaway-20260317145232.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## v_ask_receive_log

**Namn:** Varumottagningslogg

**Kort svar:** Historik over inkommande gods med inkop, mottaget antal, pall och leverantor.

**Kalla:** ASK

**Vy/tabell:** Varumottagningslogg

**ASK-omrade:** Loggar

**Datatyp:** Mottagningshistorik

**Flodessteg:** Inleverans

**Anvands i CLI v1:** Nej

### Observerat i fil

`Typ`, `Inköpsnr`, `Mottagning`, `Sändningsnummer`, `Rad`, `Leverantör`, `Leverantör.1`, `Artikel`, `Beskrivning`, `Pallid`, `Beställt`, `Mottaget`, `Pall Typ`, `Ordertyp`, `Ankomst`, `Status`, `Avvikelsekod`, `Användare`, `Vikt`, `Host`, `Ändrad`, `Bolag`, `Lager`, `Batch`, `Område`, `Utgångsdatum`, `Pris`, `SSCC`, `Release`, `Container`, `Pack Class`, `Förbehåll`, `Rowid`, `Mix Pall`, `Record Id`, `Lev batch`, `VAS`, `Crossdock`, `Serie Nummer`, `Leverantörsartikel`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Inköpsnr` | Kopplar loggen till specifik inkommande order. |
| `Mottaget` | Faktiskt mottagen kvantitet. |
| `Pallid` | Pallsparet i mottagningsflodet. |
| `Sändningsnummer` | Transportens eller ankomstens identitet. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_receive_log-20260317145157.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## v_ask_trans_log

**Namn:** Translogg

**Kort svar:** Flyttlogg for pall och artikel mellan olika lagerplatser.

**Kalla:** ASK

**Vy/tabell:** Translogg

**ASK-omrade:** Loggar

**Datatyp:** Interna lagerforflyttningar

**Flodessteg:** Interna rorelser

**Anvands i CLI v1:** Nej

### Observerat i fil

`Typ`, `Status`, `Pallid`, `SSCC`, `Artikel`, `Batch`, `Från`, `Till`, `Antal`, `Användare`, `Timestamp`, `Bolag`, `Lager`, `Host`, `Rowid`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Från` | Kallplats for intern flytt. |
| `Till` | Malplats for intern flytt. |
| `Pallid` | Pallen som flyttas. |
| `Artikel` | Artikeln pa pallen. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_trans_log-20260317170854.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## v_ask_pick_log_full

**Namn:** Plocklogg Full

**Kort svar:** Detaljerad plocklogg pa radniva med avvikelser, kvantiteter, pall och anvandare.

**Kalla:** ASK

**Vy/tabell:** Plocklogg Full

**ASK-omrade:** Loggar

**Datatyp:** Plockhandelser

**Flodessteg:** Order och plock

**Anvands i CLI v1:** Nej

### Observerat i fil

`Status`, `Typ`, `Avvikelsekod`, `Zon`, `Ordernr`, `Kundreferens`, `Multinr`, `Linjenr`, `Relnr`, `Kundnr`, `Kund`, `Lokation`, `Artikelnr`, `Artikel`, `Beställt`, `Plockat`, `Pallid`, `Host`, `Användare`, `Datum`, `Ändrad`, `Plockpallsnr`, `Transportörnr`, `Transportör`, `Order`, `Vikt`, `Batch`, `Lager`, `Bolag`, `Ansvarig inköpare`, `Volym`, `Vikt gross`, `Alt adressnr`, `Radid`, `Pris`, `Valuta`, `Beställare`, `Beställargrupp`, `Land`, `Vagn`, `Unnr`, `Adr klass`, `Original Prioritet`, `Ordertyp`, `Namn`, `Avgångsnr`, `Struktur`, `Serial nr`, `Kund ordernum`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Artikelnr` | Plockad artikel i WMS-loggen. |
| `Beställt` | Orderradens behov i plockogonblicket. |
| `Plockat` | Faktiskt plockad kvantitet. |
| `Plockpallsnr.` | Utgaende plockpall i plockflodet. |
| `Transportör` | Transportkoppling for analys mot sandning och dispatch. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_pick_log_full-20260317170910.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\v_ask_pick_log_full-20260327111807.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Analyseramottagnarader\v_ask_pick_log_full-20260416073826.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Analyseramottagnarader\v_ask_pick_log_full-20260416082044.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Analyseramottagnarader\v_ask_pick_log_full-20260422075403.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: analyze-receiving-lines-data
- Observerad sample-fil

## v_ask_correct_log

**Namn:** Saldojusteringar & Inventeringsavvikelser

**Kort svar:** Logg for saldojusteringar, avvikelsekoder och interna kommentarer.

**Kalla:** ASK

**Vy/tabell:** Saldojusteringar & Inventeringsavvikelser

**ASK-omrade:** Loggar

**Datatyp:** Korrigerings- och inventeringslogg

**Flodessteg:** Korrigeringar

**Anvands i CLI v1:** Nej

### Observerat i fil

`Typ`, `Pallid`, `Artikel`, `Beskrivning`, `Batch`, `Avvikelsekod`, `ERP-Kod`, `Antal`, `Lagerplats`, `Status`, `Användare`, `Ändrad`, `Dator`, `Lager`, `Bolag`, `Anledning`, `Index`, `Anvsvarig inköpare`, `Intern Kommentar`, `Rowid`, `Aggregerings Id`, `Lev nr`, `Leverantör`, `Robot Ind`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Avvikelsekod` | Klassificerar korrigeringens typ. |
| `Antal` | Justerad kvantitet. |
| `Lagerplats` | Platsen som paverkades. |
| `Anledning` | Manuell eller regelstyrd orsak till justeringen. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_correct_log-20260317145302.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## v_ask_assignment_move

**Namn:** Palluppdrag

**Kort svar:** Aktiva eller historiska palluppdrag med prioritet, ko, kund och plats.

**Kalla:** ASK

**Vy/tabell:** Palluppdrag

**ASK-omrade:** 2. Artikelhantering

**Datatyp:** Arbetsuppdrag for pallflyttar

**Flodessteg:** Interna rorelser och pafyllning

**Anvands i CLI v1:** Nej

### Observerat i fil

`Pallid`, `Status`, `Ordernr`, `Yta`, `Kund`, `Kund.1`, `Antal`, `Artikel`, `Beskrivning`, `Packklass`, `Ändrad`, `Användare`, `Kö`, `Prioritet`, `Lagerplats`, `Zon`, `Bolag`, `Lager`, `Batch`, `Sändningsnr`, `Utgångsdatum`, `Pågående Plock`, `Plockplats Diff`, `SSCC`, `Avgång`, `Krangång`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Pallid` | Pallen som uppdraget galler. |
| `Kö` | Arbetsko eller profil som styr truckflodet. |
| `Prioritet` | Visar uppdragets vikt i operativ ko. |
| `Lagerplats` | Nuvarande eller malrelaterad plats i uppdraget. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_assignment_move-20260317182655.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: refill-manual-picking-mg
- Observerad sample-fil

## v_ask_dispatch_pallet

**Namn:** Dispatchpallar

**Kort svar:** Utgaende pallar med plockpall, kund, sandning, transportor, kolli och volymdata.

**Kalla:** ASK

**Vy/tabell:** Dispatchpallar

**ASK-omrade:** Underhallstabeller

**Datatyp:** Utgaende pallar kopplade till leveranser

**Flodessteg:** Dispatch/utleverans

**Anvands i CLI v1:** Nej

### Observerat i fil

`Plockpallsnr.`, `Set-nr.`, `Staplingsbar`, `Palltyp`, `Pallbeskrivning`, `Pallplacering`, `Ordernr`, `Kundnr.`, `Kund`, `SRS Id`, `Transportörsnr.`, `Transportör`, `Leveransdatum`, `Status`, `SSCC`, `Plats`, `Ändrad`, `Användare`, `Bolag`, `Pappapallsnr`, `Sändningsnr`, `Rader`, `Kolli`, `Flakmeter`, `Bredd`, `Längd`, `Höjd`, `Vikt`, `Externt id`, `RFID`, `Butiknr.`, `Avgångsnr.`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Plockpallsnr.` | Identifierar utgaende plockpall. |
| `Ordernr` | Koppling till orderhuvud eller ordermangd. |
| `Sändningsnr` | Leveransens sandningsidentitet. |
| `Transportör` | Vald transportor i dispatchflodet. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\v_ask_dispatch_pallet-20260316130458.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: ask-export-files-and-customer-inputs
- Observerad sample-fil

## dimension

**Namn:** Dimensioner

**Kort svar:** Langd, bredd och hojd per dimensions-id for artiklar, platser och palltyper.

**Kalla:** ASK

**Vy/tabell:** Dimensioner

**ASK-omrade:** Underhallstabeller

**Datatyp:** Dimensionsregister

**Flodessteg:** Masterdata/plats- och pallregler

**Anvands i CLI v1:** Nej

### Observerat i fil

`Dimension Id`, `Beskrivning`, `Längd`, `Bredd`, `Höjd`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Dimension Id` | Nyckel som kopplar andra register till dimensionstabellen. |
| `Längd` | Fysisk langd. |
| `Bredd` | Fysisk bredd. |
| `Höjd` | Fysisk hojd. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\dimension-20260422053913.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: place-analysis-data
- Observerad sample-fil

## item

**Namn:** Artiklar (Item)

**Kort svar:** Artikelmaster med plockplats, palltyp, robotflagga, kategori och per-pall-regler.

**Kalla:** ASK

**Vy/tabell:** Artiklar (Item)

**ASK-omrade:** Underhallstabeller

**Datatyp:** Artikelmasterdata

**Flodessteg:** Masterdata/artikelregler

**Anvands i CLI v1:** Nej

### Observerat i fil

`Artikel`, `Säljstatus`, `Beskrivning`, `Enhet`, `Antal/enhet`, `Viktartikel`, `Utg datum`, `Batchartikel`, `Serie`, `Plockplats`, `Palltyp`, `Ändrad`, `Lagerställe`, `Bolag`, `Per pall`, `Per lav`, `Temperatur`, `Påfyllningspunkt`, `Lagerdagar`, `Utgångsdagar`, `UN nummer`, `Vikt brutto`, `Vikt netto`, `EAN13`, `Inv datum`, `Antal pack`, `Prod typ`, `Lav höjd`, `Pack klass`, `Robot`, `Vikt auto`, `Ansvarig`, `Ursprungsland`, `Intrastat`, `Intrastat besk`, `Struktur`, `Volym`, `Y-kod`, `Flampunkt`, `Supply type`, `Beg. mängd`, `Säsong`, `Lev. batch`, `Extern artikel grupp`, `Kategori`, `Krankluster`, `Staplingsbar`, `Plock Instruktion`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Artikel` | Primar artikelnyckel i masterdatat. |
| `Plockplats` | Normal eller rekommenderad plockplats. |
| `Palltyp` | Koppling till palltypsregister. |
| `Per pall` | Antal per pall i masterdata. |
| `Robot` | Flagga om artikeln ar robotrelevant. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\item-20260422055339.csv`

### Kallor

- Wiki: place-analysis-data
- Observerad sample-fil

## location

**Namn:** Lagerplatser

**Kort svar:** Platsregister med koordinater, klassning, max pall och dimensionskoppling.

**Kalla:** ASK

**Vy/tabell:** Lagerplatser

**ASK-omrade:** Underhallstabeller

**Datatyp:** Lagerplatsmaster

**Flodessteg:** Masterdata/platsregler

**Anvands i CLI v1:** Nej

### Observerat i fil

`Lagerplats`, `Typ`, `Detalj`, `Clearing`, `X-koordinat`, `Y-koordinat`, `Timestamp`, `Multi`, `Sekvens`, `Checksiffra`, `Visa saldo`, `Påfyllning`, `Max pall`, `Påfyllnings offset`, `Auto plac`, `Klassificering`, `Farligt gods`, `Dimension`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Lagerplats` | Platsens unika id. |
| `Typ` | Typ av lagerplats eller beteende. |
| `X-koordinat` | Platsens koordinat i layouten. |
| `Dimension` | Kopplar platsen till dimensionsregistret. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\location-20260422053859.csv`

### Kallor

- Wiki: place-analysis-data
- Wiki: pick-location-and-pallet-rules
- Observerad sample-fil

## main_category

**Namn:** Huvudkategori

**Kort svar:** Mappning mellan huvudkategori och kategori per bolag.

**Kalla:** ASK

**Vy/tabell:** Huvudkategori

**ASK-omrade:** Register

**Datatyp:** Kategorimappning

**Flodessteg:** Masterdata/kategori

**Anvands i CLI v1:** Nej

### Observerat i fil

`Huvudkategori`, `Kategori`, `Bolag`, `Timestamp`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Huvudkategori` | Overgripande kategori. |
| `Kategori` | Underkategori eller detaljkategori. |
| `Bolag` | Bolagsspecifik kategoriindelning. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\main_category-20260422055343.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: place-analysis-data
- Observerad sample-fil

## pallet_type

**Namn:** Palltyp

**Kort svar:** Palltyp med dimension, vikt, placering, SRS och bottenpallsprioritet.

**Kalla:** ASK

**Vy/tabell:** Palltyp

**ASK-omrade:** Register

**Datatyp:** Palltypsregister

**Flodessteg:** Masterdata/plats- och pallregler

**Anvands i CLI v1:** Nej

### Observerat i fil

`Palltyp`, `Beskrivning`, `Timestamp`, `Bit typ`, `Pall besk kort`, `COOP Besk`, `Höjd`, `Vikt`, `Flakmeter`, `Dimension`, `Nivå`, `SRS`, `Placering`, `Vagn`, `Rfid`, `Inköpstyp`, `Ordning`, `Artikel dimensioner`, `Max vikt`, `Boka alltid`, `Boka aldrig`, `Bottenpall Prio`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Palltyp` | Palltypens kod. |
| `Dimension` | Koppling till dimensionsregistret. |
| `Max vikt` | Tillaten maxvikt. |
| `Bottenpall Prio` | Prioritet som styr bottenpallshantering. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\Downloads\pallet_type-20260504125835.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: pallet-type-view
- Observerad sample-fil

## v_ask_order_log

**Namn:** Orderlogg

**Kort svar:** Orderhistorik pa huvudniva med kundreferenser, leveransvillkor och sandningsgrupper.

**Kalla:** ASK

**Vy/tabell:** Orderlogg

**ASK-omrade:** Loggar

**Datatyp:** Orderhuvud och orderhistorik

**Flodessteg:** Orderhuvud och historik

**Anvands i CLI v1:** Nej

### Observerat i fil

`Sändningsnr`, `Ordernr`, `Bolag`, `Ändrad`, `Transportörsnr.`, `Kundnr.`, `Lev. datum`, `Order datum`, `Kund adressnr.`, `Kund ref`, `Leverans text`, `Kund ordernr.`, `Skeppnings markering`, `Extra lev.`, `STORE_ADR_NUM`, `STORE_DATE`, `PACK_GROUP1`, `PACK_GROUP2`, `PACK_GROUP3`, `PACK_GROUP4`, `PACK_GROUP5`, `PACK_GROUP6`, `Ordertyp`, `Vår ref`, `Leveransvilkor`, `Sändnings-grupp id`, `Transaktionstyp`, `Utskrifts lev. text`, `Transport-produkt`, `Transport returprodukt`, `Transport förboknings-id`, `Tull`, `Original Prio`, `Bokningsbekräftelse`, `Dokumentspråk`, `Kundinstruktioner`, `Kund ref2`, `Skapad`, `Brand`, `Försäljningsorganisation`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Sändningsnr` | Sandningsidentitet pa orderhuvudsniva. |
| `Ordernr` | Orderhuvudets id. |
| `Lev. datum` | Leveransdatum i orderflodet. |
| `Försäljningsorganisation` | Overliggande saljsammanhang eller organisation. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\v_ask_order_log-20260327111749.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: place-analysis-data
- Observerad sample-fil

## v_ask_pick_location_log

**Namn:** Plockplatslogg

**Kort svar:** Visar tidigare och ny plats for artiklar nar plockplats flyttas eller andras.

**Kalla:** ASK

**Vy/tabell:** Plockplatslogg

**ASK-omrade:** Loggar

**Datatyp:** Historik over plockplatsbyten

**Flodessteg:** Platsforandringar

**Anvands i CLI v1:** Nej

### Observerat i fil

`Artikel`, `Beskrivning`, `Lager`, `Bolag`, `Batch`, `Tidigare plats`, `Ny plats`, `Användare`, `Host`, `Timestamp`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Artikel` | Artikeln vars plockplats andrats. |
| `Tidigare plats` | Forra plockplatsen. |
| `Ny plats` | Nya plockplatsen. |
| `Användare` | Vem som gjorde andringen. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Platsanalys\data\v_ask_pick_location_log-20260327111817.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: place-analysis-data
- Observerad sample-fil

## asw_order

**Namn:** ASW Order

**Kort svar:** Orderrader fran ASW med kund, artikel, antal, leveransadress, sandning och logistikattribut.

**Kalla:** ASW

**Vy/tabell:** ASW Order

**ASK-omrade:** Integration

**Datatyp:** Order- och radinformation fran ASW

**Flodessteg:** Order och plock

**Anvands i CLI v1:** Nej

### Observerat i fil

`Status`, `Ordernr`, `Rad`, `Kund`, `Transportör`, `Orderdatum`, `Artikel`, `Antal`, `Ordertext`, `Radtext`, `Skickat`, `Mottaget`, `Leveransdatum`, `Rowid`, `Skapad av`, `Lager`, `Bolag`, `Batch`, `Meddelande`, `Inköpsnr`, `Inköpsrad`, `Extra lev.`, `Ordertyp`, `Adr num`, `Kund ordernr.`, `Adress 1`, `Adress 2`, `Adress 3`, `Adress 4`, `Post nr`, `Adress namn`, `Kund ref`, `Land`, `EMail`, `Telefon`, `Pickup point`, `Pris`, `Valuta`, `Prioritet`, `Skall`, `Allokeringsmetod`, `Utgångsdatum`, `Custom store num`, `Custom store name`, `Skeppnings markering`, `Leverans text`, `Butiks nr`, `Butiksnamn`, `Butiksaddress 1`, `Butiksaddress 2`, `Butiksaddress 4`, `Butikspostnummer`, `PACK_GROUP1`, `PACK_GROUP2`, `PACK_GROUP3`, `PACK_GROUP4`, `PACK_GROUP5`, `PACK_GROUP6`, `STORE_DATE`, `Leveransvilkor`, `Sändningsnr`, `Sändnings-grupp id`, `Transaktionstyp`, `GTIN`, `Utskrifts lev. text`, `Produkt`, `Stat`, `GLN`, `Returprodukt`, `Plockstruktur`, `Butiksmail`, `Butikstelefon`, `Butiksland`, `Transport förboknings-id`, `Serie`, `Bokningsbekräftelse`, `RecordId`, `Dokumentspråk`, `Kundinstruktioner`, `Kund ref2`, `Brand`, `Försäljningsorganisation`, `Rutt laststart`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Ordernr` | Affarssystemets order-id. |
| `Rad` | Orderrad inom ordern. |
| `Artikel` | Bestalld artikel. |
| `Antal` | Bestalld kvantitet. |
| `Sändningsnr` | Sandningskoppling mellan ASW och WMS. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Analyseramottagnarader\asw_order-20260416115204.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Analyseramottagnarader\asw_order-20260416115236.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Analyseramottagnarader\asw_order-20260422080114.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: analyze-receiving-lines-data
- Observerad sample-fil

## dblog_pick_log

**Namn:** Arkiv Plocklogg

**Kort svar:** Arkiverad plocklogg pa radniva, anvandbar for historiska analyser och felsokning.

**Kalla:** ASK

**Vy/tabell:** Arkiv Plocklogg

**ASK-omrade:** X. Arkiv

**Datatyp:** Arkiverad plocklogg

**Flodessteg:** Order och plock / arkiv

**Anvands i CLI v1:** Nej

### Observerat i fil

`Typ`, `Ordertyp`, `Ordernr`, `Rad`, `Kund`, `Artikel`, `Starttid`, `Avvikelsekod`, `Transportör`, `Vikt`, `Beställt`, `Plockat`, `Användare`, `Status`, `Timestamp`, `Lager`, `Multi`, `Adr num`, `Release`, `Plockzon`, `Pallid`, `Host`, `Ändradint`, `Plockpallsnr.`, `Bolag`, `Batch`, `Lagerplats`, `Rowid`, `Pris`, `Valuta`, `Vikt gross`, `Volym`, `Struktur`, `Vagn`, `Kund ref`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Ordernr` | Orderkoppling i arkivloggen. |
| `Artikel` | Plockad artikel pa radniva. |
| `Plockzon` | Zon dar plocket skedde. |
| `Pallid` | Pallsparet i arkiverat plockflode. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Analyseramottagnarader\dblog_pick_log-20260416120733.csv`
- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\Analyseramottagnarader\dblog_pick_log-20260416120915.csv`

### Kallor

- Wiki: ask-csv-to-views
- Wiki: analyze-receiving-lines-data
- Observerad sample-fil

## prognos_excel

**Namn:** Kundprognos Excel

**Kort svar:** RELEX-lik prognosfil som normaliseras till artikel, beskrivning, antal styck, rader och butiker.

**Kalla:** Kund-Excel

**Vy/tabell:** Kundprognos

**ASK-omrade:** Kundinput

**Datatyp:** Planerings- och behovsprognos

**Flodessteg:** Prognos och planering

**Anvands i CLI v1:** Nej

### Observerat i fil

`Artikelnummer`, `Beskrivning`, `Antal styck`, `Antal rader`, `Antal butiker`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Product code / Artikelnummer` | Artikel-id i prognosen. |
| `Product name / Beskrivning` | Artikelnamn eller beskrivning. |
| `Antal styck` | Behov i styck efter normalisering. |
| `Antal rader` | Antal order- eller prognosrader. |
| `Antal butiker` | Hur manga butiker som omfattas. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\Prognos idag_1227934_3062311682516.xlsx`

### Kallor

- Originalkod: read_prognos_xlsx
- Observerad sample-fil

## campaign_excel

**Namn:** Kampanjvolym Excel

**Kort svar:** Kampanjvolymfil som normaliseras till artikelnummer och antal styck efter rad- och kolumnrensning.

**Kalla:** Kund-Excel

**Vy/tabell:** Kampanjvolymer

**ASK-omrade:** Kundinput

**Datatyp:** Kampanj- och volymplanering

**Flodessteg:** Prognos och planering

**Anvands i CLI v1:** Nej

### Observerat i fil

`Artikelnummer`, `Antal styck`

### Viktiga kolumner

| Kolumn | Tolkning |
| --- | --- |
| `Artikelnummer` | Kampanjartikel efter normalisering. |
| `Antal styck` | Volym per artikel efter normalisering. |

### Kodifierat i logik

Inga kanda kodalias i nuvarande CLI.

### Sample-filer

- `C:\Users\emikad\OneDrive - Dole Nordic AB\Skrivbordet\projects\allokering\testdata\Granngården prognos kampanjplock +6v_1208059_2993592205456.xlsx`

### Kallor

- Originalkod: read_campaign_xlsx
- Observerad sample-fil

### Osakert eller infererat

- Rafilen ar inte en vanlig tabell utan ett kampanjgrid som maste rensas fore tolkning.
