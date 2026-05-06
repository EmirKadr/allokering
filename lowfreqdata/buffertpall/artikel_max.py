"""Uppdatera observations.csv.gz fran v_ask_article_buffertpallet och racka om artikel_max.csv.

Pipeline:
1. Las buffer-CSV (tab-separerad)
2. Filtrera status 30 (ENDAST status 30 sparas i observations)
3. Lagg till nya pallid i observations.csv.gz (dedup pa pallid)
4. Racka om artikel_max.csv fran ALLA observationer med outlier-filter
"""

import sys
import pandas as pd
import numpy as np
from pathlib import Path

OBSERVATIONS_COLS = ['artikelnummer', 'pallid', 'antal']


def max_utan_outlier(grupp: pd.DataFrame) -> tuple[float, str]:
    """Returnera (max, pallid) efter Tukey IQR-outlier-filter.

    grupp ska ha kolumner 'antal' och 'pallid' med unikt index.
    Filtret aktiveras bara nar gruppen har >2 pallar.
    """
    if len(grupp) > 2:
        q1, q3 = np.percentile(grupp['antal'], [25, 75])
        ovre = q3 + 1.5 * (q3 - q1)
        sub = grupp[grupp['antal'] <= ovre]
        if not sub.empty:
            row = sub.loc[sub['antal'].idxmax()]
            return float(row['antal']), str(row['pallid'])
    row = grupp.loc[grupp['antal'].idxmax()]
    return float(row['antal']), str(row['pallid'])


def lasin_observations(path: Path) -> pd.DataFrame:
    if path.exists() and path.stat().st_size > 0:
        df = pd.read_csv(path, dtype=str)
        for col in OBSERVATIONS_COLS:
            if col not in df.columns:
                df[col] = ''
        return df[OBSERVATIONS_COLS]
    return pd.DataFrame(columns=OBSERVATIONS_COLS)


def racka_om_artikel_max(observations: pd.DataFrame, ut_path: Path) -> int:
    """Racka om artikel_max.csv fran observations. Returnerar antal artiklar."""
    if observations.empty:
        pd.DataFrame(columns=['artikelnummer', 'max', 'pallid']).to_csv(
            ut_path, index=False, encoding='utf-8-sig'
        )
        return 0

    df = observations.copy()
    df['antal'] = pd.to_numeric(df['antal'], errors='coerce')
    df = df.dropna(subset=['artikelnummer', 'antal', 'pallid'])
    df['artikelnummer'] = df['artikelnummer'].astype(str).str.strip()
    df['pallid'] = df['pallid'].astype(str).str.strip()
    df = df.drop_duplicates(subset='pallid').reset_index(drop=True)

    rader = []
    for art, grupp in df.groupby('artikelnummer'):
        max_val, pall_id = max_utan_outlier(grupp)
        rader.append({'artikelnummer': art, 'max': max_val, 'pallid': pall_id})

    pd.DataFrame(rader, columns=['artikelnummer', 'max', 'pallid']).to_csv(
        ut_path, index=False, encoding='utf-8-sig'
    )
    return len(rader)


def main(buffer_path: Path) -> None:
    df = pd.read_csv(buffer_path, sep='\t', dtype=str, encoding='utf-8-sig')
    df['Antal'] = pd.to_numeric(df['Antal'], errors='coerce')
    df['Status'] = pd.to_numeric(df['Status'], errors='coerce')
    df = df.dropna(subset=['Artikel', 'Antal', 'Status', 'Pallid'])
    df['Artikel'] = df['Artikel'].astype(str).str.strip()
    df['Pallid'] = df['Pallid'].astype(str).str.strip()

    df = df[df['Status'] == 30]

    nya = pd.DataFrame({
        'artikelnummer': df['Artikel'].values,
        'pallid': df['Pallid'].values,
        'antal': df['Antal'].astype(int).astype(str).values,
    }).drop_duplicates(subset='pallid')

    obs_path = buffer_path.parent / 'observations.csv.gz'
    befintliga = lasin_observations(obs_path)
    befintliga_pallid = set(befintliga['pallid'].astype(str))

    nya = nya[~nya['pallid'].isin(befintliga_pallid)]
    n_nya = len(nya)

    if n_nya:
        kombinerat = pd.concat([befintliga, nya], ignore_index=True)
    else:
        kombinerat = befintliga

    kombinerat.to_csv(obs_path, index=False, compression='gzip')

    max_path = buffer_path.parent / 'artikel_max.csv'
    n_artiklar = racka_om_artikel_max(kombinerat, max_path)

    print(f'Nya pallid tillagda: {n_nya} (totalt {len(kombinerat)} observationer).')
    print(f'artikel_max.csv uppdaterad: {n_artiklar} artiklar.')


if __name__ == '__main__':
    main(Path(sys.argv[1]))
