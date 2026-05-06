import pandas as pd
from pathlib import Path
from datetime import datetime

obs = Path('lowfreqdata/buffertpall/observations.csv.gz')
mxp = Path('lowfreqdata/buffertpall/artikel_max.csv')

print('=== observations.csv.gz ===')
if obs.exists():
    st = obs.stat()
    print(f'  storlek: {st.st_size} bytes')
    print(f'  senast andrad: {datetime.fromtimestamp(st.st_mtime)}')
    df = pd.read_csv(obs, dtype=str)
    print(f'  rader: {len(df)}')
    print(f'  unika pallid: {df["pallid"].nunique()}')
    print(f'  TESTART finns: {(df["artikelnummer"]=="TESTART").any()}')
else:
    print('  SAKNAS')

print()
print('=== artikel_max.csv ===')
if mxp.exists():
    st = mxp.stat()
    print(f'  storlek: {st.st_size} bytes')
    print(f'  senast andrad: {datetime.fromtimestamp(st.st_mtime)}')
    mxd = pd.read_csv(mxp, dtype=str)
    print(f'  artiklar: {len(mxd)}')
    print(f'  TESTART finns: {(mxd["artikelnummer"]=="TESTART").any()}')
else:
    print('  SAKNAS')
