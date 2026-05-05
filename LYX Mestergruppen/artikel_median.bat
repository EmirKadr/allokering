@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

set "FOUND="
for /f "delims=" %%F in ('dir /b /a:-d /o:-d "v_ask_article_buffertpallet*" 2^>nul') do (
    if not defined FOUND set "FOUND=%%F"
)

if not defined FOUND (
    echo Hittade ingen fil som borjar med v_ask_article_buffertpallet i %~dp0
    pause
    exit /b 1
)

echo Anvander indatafil: %FOUND%

python -c "import sys, pandas as pd; from pathlib import Path; p=Path(sys.argv[1]); df=pd.read_csv(p, sep='\t', dtype=str, encoding='utf-8-sig'); df['Antal']=pd.to_numeric(df['Antal'], errors='coerce'); df=df.dropna(subset=['Artikel','Antal']); df['Artikel']=df['Artikel'].astype(str).str.strip(); out=p.parent/'artikel_median.csv'; ex=(pd.read_csv(out, dtype=str, encoding='utf-8-sig') if (out.exists() and out.stat().st_size > 0) else None) if True else None; existing=set(ex['artikelnummer'].astype(str).str.strip()) if (ex is not None and 'artikelnummer' in ex.columns) else set(); new_df=df[~df['Artikel'].isin(existing)]; med=new_df.groupby('Artikel')['Antal'].median().reset_index().rename(columns={'Artikel':'artikelnummer','Antal':'median'}); (med.to_csv(out, mode='a', header=False, index=False, encoding='utf-8') if (existing and not med.empty) else None) if existing else med.to_csv(out, index=False, encoding='utf-8-sig'); print(f'La till {len(med)} nya artiklar i {out} (befintliga {len(existing)} oforandrade).' if existing else f'Skapade {out} med {len(med)} unika artiklar.')" "%FOUND%"


if errorlevel 1 (
    echo Python-skriptet misslyckades.
    pause
    exit /b 1
)

pause
