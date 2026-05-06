# Testning och CLI

Projektet har nu en inbyggd CLI i `allokering12.1.py` sa att centrala arbetsfloden kan koras utan GUI.

## Installera testberoenden

```powershell
pip install -r requirements-dev.txt
```

## Koer testerna

```powershell
python -m pytest -q
```

## Visa tillgangliga CLI-kommandon

```powershell
python allokering12.1.py --help
```

## Exempel

Allokering:

```powershell
python allokering12.1.py allocate `
  --orders .\orders.csv `
  --buffer .\buffer.csv `
  --result-out .\out\allocated.csv `
  --near-miss-out .\out\near_miss.csv `
  --json
```

OrderSaldo:

```powershell
python allokering12.1.py ordersaldo `
  --orders .\ordersaldo.csv `
  --saldo .\saldo.csv `
  --complete-orders-out .\out\complete.txt `
  --shortage-out .\out\shortage.csv `
  --json
```

LYX:

```powershell
python allokering12.1.py lyx `
  --saldo .\saldo.csv `
  --max-csv .\lowfreqdata\buffertpall\artikel_max.csv `
  --output .\out\lyx.txt `
  --json
```

Pafyllnadsprio:

```powershell
python allokering12.1.py pafyllnadsprio `
  --orders .\ordersaldo.csv `
  --saldo .\saldo.csv `
  --report-out .\out\pafyllnadsprio.xlsx `
  --json
```

HIB-koppling:

```powershell
python allokering12.1.py hib-koppling `
  --details .\bestallningslinjer.csv `
  --overview .\orderoversikt.csv `
  --changes-out .\out\hib_andringar.csv `
  --missed-out .\out\hib_missade.csv `
  --json
```

Orderkontroll:

```powershell
python allokering12.1.py overview-check `
  --overview .\orderoversikt.csv `
  --details .\bestallningslinjer.csv `
  --report-out .\out\orderkontroll.xlsx `
  --json
```

Dispatchkontroll:

```powershell
python allokering12.1.py dispatch-check `
  --overview .\orderoversikt.csv `
  --dispatch .\dispatchpallar.csv `
  --details .\bestallningslinjer.csv `
  --report-out .\out\dispatchkontroll.csv `
  --json
```

Vecka 27:

```powershell
python allokering12.1.py vecka27-check `
  --orders .\bestallningslinjer.csv `
  --report-out .\out\vecka27.txt `
  --json
```

Eftersok:

```powershell
python allokering12.1.py eftersok `
  --purchase 123456 `
  --article 2003511 `
  --wms-receive .\v_ask_receive_log.csv `
  --wms-booking .\v_ask_booking_putaway.csv `
  --wms-buffert .\v_ask_article_buffertpallet.csv `
  --wms-trans .\v_ask_trans_log.csv `
  --wms-pick .\v_ask_pick_log_full.csv `
  --wms-correct .\v_ask_correct_log.csv `
  --report-out .\out\eftersok.txt `
  --json
```
