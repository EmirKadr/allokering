# order_saldo_gui.py (v3.4)
# – Komplett version baserad på din uppladdade fil, med uppdaterad analyze():
#   * Endast "Plock ≥ Beställt" avgör tillräckligt saldo
#   * "Kompletta ordrar" kräver att alla rader uppfyller det
#   * 1×N-flikarna visar bara ordrar där saldot räcker
#   * "Plockat" läses in men används inte i beräkningarna

from __future__ import annotations

import os
import re
import base64
import gzip
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# DnD är valfritt (pip install tkinterdnd2)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _DND_AVAILABLE = True
except Exception:
    _DND_AVAILABLE = False

APP_TITLE = "Order/Saldo-analys – Fristående GUI"

# ---------- Hjälpfunktioner ----------

# ---------- Konfiguration för förbjudna artiklar ----------
# Artiklar med vissa mått är förbjudna (om längd, bredd eller höjd är exakt 150 cm, eller om längd + omkrets > 300 cm).
# Dessutom är artiklar som väger mer än 20 kg förbjudna att förekomma i 1×N-listorna.
# Orderlistorna 1×N filtrerar bort ordrar som innehåller någon av dessa artiklar.


_RESTRICTED_ARTICLES_B64: str = (
    "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"
    "/MS5xoz4cu9ldv1HfI34GvE14ss5KBA+4su5yJz4co9qnJPMia8TX85N1vf3iQbCN+Ab8DGfOvOpD+z1Ijz4jXNPdBMe/MY5KJB+"
    "4TfORYHw4TfHb47fHL85fuMcFQgffnP85viNc1YgfO0359zlrCuBE1zgBg8ooIIGwjfgG/AN+DrPnHONs94464yzn3b2pz5fPQG7"
    "PdYbZ71x7vWcez3nXs+513Pu9Zx7Pecez7m3c+7pnHs55x7OWX+cdcdZb5x1xllfnHXFWU+cdcRZP5x1w1kvnHXCWR+cdcFZD5x1"
    "wJn/nXnfme+ded6Z35153ZnPnXncmb+deduZr5152pmfnXnZmY+dediZf51515lvnXnWmV9diYvOW24/c2/k2uuzK37jHilwgwcU"
    "ED78yb2KM396/04WiN+4n/C+n4hZvvycqKAX7mUg5YpT4gAnuBqLN1aBZ4Eb5LnyXPu5Vv740+fixAPKv/zgIb8QKIz8brxlLRwj"
    "T/Qp5Qm1sTVWjITCmEEaJ7jADbZ9jpDGrs8b4EJb0ZnM+7ekv1L9SDw+pcjQLp2x9FOyW8oPD/No/Ke0bknyN+iXk9JbJ586+dTp"
    "p04/dfmz9Kc0P6X1Ke1P6cvyvlHsRNfcn9L5lORXWuAGParenlYp7xne0vjUjU/d/tTtt27Fjt3fuir9rfvTXpU+deNTNz51+1P3"
    "tqcxcGItf0vyKemnlIv4ZaEkf0s/Tf9o+kfT/2jGbumPJiX5W/ppbvDW1ujIL3Ceyvf8jLZRyiI/7wA3eEABtdEaV8Q6rLf8KdF6"
    "l+anbn7q1qdu/a2Tj5187ORjJ3/s+l3yV8vGAU6w3y1nq0ZrPNQfnp8NHlDBq087+Cy/222ER+CR5nG87rs+kjGjtD6l/f+SvCWL"
    "qWT+StVKflfU2K3M0W89x/KOVUqbZ5t3Sq+ueaXtLZ28J0Za40rySnprybKQ1rjS8atnPXfOyKXbmuCTku4zcm/mJ9HPlfat9UOv"
    "NL+nR3J6oGe80rWNkcD7Wn7fUJKTpzNvNBtvGU7vVsidwHHxsSvpRprgAu9zm7DUt28tsQKlpHAP7ILgSvktbksK0uplXGPSaqQ1"
    "jAsb3m3udfH2aK/3mW650nmfnfNKcBxakfzNGYk41g8YSLdOL55XIv55nYKdxmInzE9d0k/JfiW9bdm6XrbN+66t3u8V/Vj+MAKy"
    "pO94iJL8qTvT8M4hBw89za+rG3OhvnHvO+3+WLIZcg/VH0XGoLnrxi3NT4kMvKW/ddOoMwG7F0aGGtExYuNEMLb6RDA20b3bWaOz"
    "Zs0Xd+OmvCn3m668DU2MqTj8xLxwS/v8KdHLFcnWGOlHPW1tNHbvo1ZMBrElwNvrxMHgPL/SyFngLc1P6Xw05VPST8l/JYnx2j2O"
    "I4V2X1LylmLL0fkjOuzRls5oKdS156uSrKX1GLUrf0WauSjW77FXUiTrDMm/VuiZMAxW57/lD8aOJOjFDnJR63PfZ94xLKlrNfev"
    "LeXf5iCRFyX9nu1X6v7pjvFTkqn2HjmlXodKWq+0X6ltY8vS60JK9M/fd0tptiT565nWHmvEbFvSkiud/CYgpfyzrNNSTD+vdAZS"
    "/uaGJM3n9btHS3ljV5JFKqYU027+SojkAym/dUY691n+6nSlYhn5KfBEyrNSSiv/cqWkvfI8VtLJv8VKKZJk9LMzcy1DsoWU38S3"
    "lHdKJa38HrelM2DZ+a39lcYrrVfarwRfpHfbSv4FHFL+pQuSCdLShZTf617pPvNbO/KOuKSZX5VeyZDW41eiL/K+b+zIjb7UF1lI"
    "t1fz+Hil9UqXJVL6Svt5pWth/pOw3Xl6vdK5ktLT7Yfas+6zdMIr8UzzW52WzvpJ8NlDFPqve670ezZfab3SfqXzSv1uuWzplTqL"
    "84PfWzvzC+eS6kuSlszus7ytu1L3II6O+pPmK61X2q90Xkleib7sO2byByxaO9dX+SPUfTbJ8ZRgPutcKb/GQiKLU1qv9KuVV7rM"
    "+/b5CKMxpauntweyrw9q13cleqUPmZ0/2NB7zTuKkizS/Up+ny3yNKTrK1vXV7auh+zmUEjMAim9LPbW+mtxPWn79q/ugZH0eaW3"
    "9vrAV+dkrNnaHp+xZ++xtZ76CrKkk/fPJensmWbllPlKU15JX8mu1Fm8+nv/K7226zLnX322RFRT8veZX9vTIyUmg1ubUtfGqtsZ"
    "llLPu2vWnYqyd+ierp1fVJckzColyZXoX0rdv2XMByXt/wCwUsoWdzwAAA=="
)

# Den här koden introducerar en ny base64-sträng som ersätter den ursprungliga listan.
# Denna sträng innehåller alla artikelnummer som uppfyller något av måttkraven ovan eller har en vikt över 20 kg.
_NEW_RESTRICTED_ARTICLES_B64: str = (
    "H4sIAALgJmkC/2WcW7LtKAhAJ9TVFUVe859Ya5Rlzu6692MlKhhANJp92vM8I59/n3/apLAocoXsUKpA955val0HpJBBVU9LR39iQKdeN+ubRNqlAe1669+lPBRj96+1lltH0zZOqWYf0GlrZlqUu38tTE5p5NHbZgODTr3UyKJjv5amA6pSdKTH7l+Xnm2Tdx+QQ3Go"
    "uPbc1sO+/rzSey8AI9wAzlUEH5Bn9j5Z/XiOdH/zzOrZE/tdczmJrP7qjX90TJp3z7+a0M2brsa6T0/+uT/75/CU9c516tnXtfU0"
    "zm2BbN9Etv5A8uClfcpgzRDUmuMv/hx9UNx96a5Yd6PCwdEDi7Y2nlTg/68mc+Y4sV4eQUveC6qgr+eIz+epuQv9+3uQJh84iK9K"
    "wX+z9u2SZYQWzutF8wvb8NCv7X8KYXelyYuUiMSj1EUxhGJoyQg+z9Y2Q0vLMpZW4xdPUytNi5r2eIYfvdYE+S1vjGdrQJLryT7T"
    "p+qFd4gHjkMV6LMRTsHEF0QKW408c4pJQvPjJ/k5h+/j587m/Eu0M3O0cv59/xE6ir4fiQ4HAcIg0UA09XjNBxLx13d7y5HStYxY"
    "nXaZ6PSxwm2sTVD9UHB6ZVrkjb6zY9Opm4bLEkofPwrYoG3fna2p39GjJ1l/9WVzv+30iXgi8hQ+2KCYawJskL2YT3Qt8WIefhD9I"
    "zovHfwmfi38e/Jj/T+FV/prxDzv+bsKovHbYXPHGHk1FEHwo+gA6wNSxg3O9LWVrSO/dD0cfJpnOIj6iX9oIeqX9I+mP6RydmRT8"
    "IHtvMd4z6J/Jb2oe6Gv9X/aFSfvFX5hfOf4V/Fi8pvWrdcKMWbeKV55OHysT289D3wT7o8/Cv0Tn4RDnkF3lFO0a3p4JG7pJkdr0"
    "wER39Ec9mPb37DUhfNf07/Ab+SvddNEPTyJL748G2t0/k+fQ/8Zvxf+feDfviJnscjbdUt4p78n3/Jmqu+SfgUC9wAW/Mt+yb31/"
    "vnqWwbwYfFmq3zi0/bPBy/V9fbi0YDnePhtceo/9AH6hHfGz+hP+oT/1PtjvutrbGV8dxO4N8C9+vODnqI+8SLl5kDBV3CeyZ0jj"
    "J/BH/Mb8R3+hd9yX/iWOy89goK/1N/rtp318gCyvOEfOt4uP9NrmERW4Kw39yHfG7kiLuQv4N8If0mfjb9X/Fq7NHvQTH/Scu1T7"
    "D6n7Jr/1Z+iFf/EeyN/wR/Qi+yGfyN/wR/Qh7+Z751M8E1yU/ws/shf8If0pn4P/S1wH/B3+h/9nftB/wSfFT8or/2v0L/0nvwJv"
    "9X/0D8Nf8Lv4XfC74HfB74LfBb8Lfhf8Lvhd8LPgZ8HPgp8FPwt+Fvws+Fnws+Bnwc+CnwU/C34W/Cz4WfCz4WfCz4WfCz4WfCz4"
    "WfCz4WfCz4WXxon3P8qn3X20A/ke5r9VF+E5+lf0vb5JfzCcwj9YH3Ys/CuRFwoWIv+ljpuREr33H2CIr5/fgUV7Hf0LGdyH7q8m"
    "zL25oZ8Z20YQVNjDrm3fgik9JlpOJP2kz7+aPnJoEfFn5u+E/zWMr5b+p+if7/lY+uf8X8rP0Dt2FvX3fc39ufq/2Q4hKg5D/b/6"
    "P/0HyM/yW9+afU/8bV8iX8MB/4d/sdL/vPh4Z/Z41/cV/U/jO/i3wp/pb+Bn/en/oz/MTO4rJ6Lh4HwM8HPyfsbkOn47/mP8+T/T"
    "/+Cl/k14EF9Ef8z3Htf7Pv9E+UzwLVwvf835/dz/1m/Ev/yf8MA/j//4b8AEAAA=="
)

# Dekomprimera bas64-strängen en gång vid import för att skapa mängden förbjudna artiklar.
try:
    # Använd den nya base64-strängen (_NEW_RESTRICTED_ARTICLES_B64) istället för den ursprungliga för att generera mängden
    # förbjudna artiklar. Detta inkluderar både mått- och viktkrav.
    _RESTRICTED_ARTICLES_DATA = gzip.decompress(base64.b64decode(_NEW_RESTRICTED_ARTICLES_B64)).decode("utf-8")
    _RESTRICTED_ARTICLES = set(filter(None, _RESTRICTED_ARTICLES_DATA.split(",")))
except Exception:
    _RESTRICTED_ARTICLES = set()

def load_restricted_articles() -> set[str]:
    """
    Returnera en kopia av mängden förbjudna artikelnummer.
    Denna funktion gör att listan kan användas utan att riskera mutation av det interna setet.
    """
    return set(_RESTRICTED_ARTICLES)

def _lazy_pd():
    """Lazy-import av pandas."""
    import pandas as pd  # type: ignore
    return pd

def read_csv_flex(path: str):
    """Läs CSV/TXT med försök på tab/;/, som svenska exporter ofta använder."""
    pd = _lazy_pd()
    seps = ["\t", ";", ","]
    last_err = None
    for sep in seps:
        try:
            df = pd.read_csv(path, sep=sep, engine="python")
            # Om fel sep gav en enda kolumn, prova nästa
            if df.shape[1] > 1:
                return df
        except Exception as e:
            last_err = e
    if last_err:
        raise last_err
    raise ValueError("Kunde inte läsa filen med vanliga separatorer (tab/;/,).")

def to_numeric_safe(series):
    pd = _lazy_pd()
    return pd.to_numeric(series, errors="coerce").fillna(0)

def df_to_treeview(tree: ttk.Treeview, df):
    tree.delete(*tree.get_children())
    cols = list(df.columns)
    tree["columns"] = cols
    tree["show"] = "headings"
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=max(80, min(260, int(len(c) * 8))), anchor="w")
    for _, row in df.iterrows():
        tree.insert("", "end", values=[row[c] for c in cols])

def open_text_window(parent, title: str, lines: list[str]):
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("600x480")
    txt = tk.Text(win, wrap="none")
    txt.pack(fill="both", expand=True)
    txt.insert("1.0", "\n".join(lines))
    txt.configure(state="normal")

def open_df_window(parent, title: str, df):
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("900x520")
    frm = ttk.Frame(win)
    frm.pack(fill="both", expand=True)
    tree = ttk.Treeview(frm)
    ysb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
    xsb = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
    tree.configure(yscroll=ysb.set, xscroll=xsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    ysb.grid(row=0, column=1, sticky="ns")
    xsb.grid(row=1, column=0, sticky="ew")
    frm.columnconfigure(0, weight=1)
    frm.rowconfigure(0, weight=1)
    df_to_treeview(tree, df.reset_index(drop=True))

def save_df_dialog(parent, df, default_name: str):
    path = filedialog.asksaveasfilename(
        parent=parent,
        defaultextension=".csv",
        initialfile=default_name,
        filetypes=[("CSV", "*.csv"), ("Text", "*.txt"), ("Alla filer", "*.*")]
    )
    if not path:
        return
    _, ext = os.path.splitext(path)
    if ext.lower() == ".txt" and df.shape[1] == 1:
        with open(path, "w", encoding="utf-8") as f:
            for v in df.iloc[:, 0].astype(str):
                f.write(str(v) + "\n")
    else:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    messagebox.showinfo(APP_TITLE, f"Sparat: {path}")

def copy_to_clipboard(root: tk.Tk, text: str):
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()

# ---------- Autokartläggning av kolumner ----------

def _normalize(s: str) -> str:
    s = s.lower()
    s = s.replace("å", "a").replace("ä", "a").replace("ö", "o")
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s

CANDIDATES = {
    "order": [
        "ordernr", "ordernummer", "ordernumber", "orderno", "orderid", "order",
        "ordernrnr"
    ],
    "article": [
        "artikel", "artikelnr", "artikelnummer", "artnr",
        "item", "itemno", "sku", "productcode", "artikel1"
    ],
    "name": [
        "benamning", "artikelnamn", "namn", "produktnamn", "description", "productname", "artikel1"
    ],
    "pick": [
        "plock", "saldo", "lager", "stock", "available", "availableqty", "qtyavailable"
    ],
    "demand": [
        "bestallt", "bestalltantal", "bestalld", "antal", "quantity", "qty", "ordered", "orderqty"
    ],
    "pickedqty": [
        "plockat", "plockad", "picked", "pickedqty", "qtypicked"
    ],
    # NYTT: Plockplats för SK-filter
    "pickloc": [
        "plockplats", "lagerplats", "bin", "bincode", "location", "loc", "picklocation", "plockplatsnr"
    ],
}

REQUIRED_KEYS = ("order", "article", "pick", "demand")

def auto_map_columns(df) -> dict[str, str | None]:
    cols = list(df.columns)
    norm_map = {_normalize(c): c for c in cols}

    mapping: dict[str, str | None] = {
        "order": None, "article": None, "name": None,
        "pick": None, "demand": None, "pickedqty": None, "pickloc": None
    }

    # Exakta träffar
    for key, cand_list in CANDIDATES.items():
        for cand in cand_list:
            if cand in norm_map:
                mapping[key] = norm_map[cand]
                break

    # Fuzzy fallback (innehåll av kandidat)
    for key in mapping:
        if mapping[key]:
            continue
        for norm, original in norm_map.items():
            if any(norm.find(cand) >= 0 for cand in CANDIDATES.get(key, [])):
                mapping[key] = original
                break

    missing = [k for k in REQUIRED_KEYS if not mapping[k]]
    if missing:
        miss_names = ", ".join(missing)
        raise ValueError(
            "Kunde inte hitta obligatoriska kolumner: "
            f"{miss_names}. Byt gärna kolumnnamn i din fil till något vanligt (t.ex. "
            "'Order nr', 'Artikel', 'Plock', 'Beställt')."
        )
    return mapping

# ---------- Huvud-appen ----------

_BaseTk = TkinterDnD.Tk if _DND_AVAILABLE else tk.Tk

class App(_BaseTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1020x760")

        self.df = None
        self.mapping: dict[str, str | None] = {}

        # Överdel: fil
        top = ttk.LabelFrame(self, text="Inläsning")
        top.pack(fill="x", padx=10, pady=10)

        self.path_var = tk.StringVar()
        ttk.Button(top, text="Öppna CSV…", command=self.load_csv_dialog).grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(top, textvariable=self.path_var, width=80, state="readonly").grid(row=0, column=1, padx=6, pady=6, sticky="w")
        hint = " (dra & släpp fil i fönstret)" if _DND_AVAILABLE else ""
        ttk.Label(top, text=f"Tips: Öppna en fil{hint}").grid(row=0, column=2, padx=6, pady=6, sticky="w")

        # Info-rad för upptäckta kolumner
        self.map_label = ttk.Label(top, text="Kolumner: –")
        self.map_label.grid(row=1, column=0, columnspan=3, padx=6, pady=(0, 6), sticky="w")

        ttk.Button(top, text="Analysera igen", command=self.analyze).grid(row=2, column=0, padx=6, pady=6, sticky="w")

        # Notebook med 2 flikar + nya
        self.frame_lists = ttk.Notebook(self)
        self.frame_lists.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab1 = ttk.Frame(self.frame_lists)
        self.tab2 = ttk.Frame(self.frame_lists)
        self.frame_lists.add(self.tab1, text="Lista 1 – Kompletta ordrar")
        self.frame_lists.add(self.tab2, text="Lista 2 – Artiklar att beställa")

        # TAB 1
        self.lbl1 = ttk.Label(self.tab1, text="Ingen data än.")
        self.lbl1.pack(anchor="w", padx=8, pady=8)
        btns1 = ttk.Frame(self.tab1); btns1.pack(anchor="w", padx=8, pady=4)
        self.btn1_copy = ttk.Button(btns1, text="Kopiera ordernummer", command=self.copy_tab1, state="disabled")
        self.btn1_copy.pack(side="left", padx=4)
        self.btn1_details = ttk.Button(btns1, text="Läs mer", command=self.details_tab1, state="disabled")
        self.btn1_details.pack(side="left", padx=4)
        self.btn1_export = ttk.Button(btns1, text="Exportera (txt)", command=self.export_tab1, state="disabled")
        self.btn1_export.pack(side="left", padx=4)

        # TAB 2 (helheten)
        self.lbl2 = ttk.Label(self.tab2, text="Ingen data än.")
        self.lbl2.pack(anchor="w", padx=8, pady=8)
        btns2 = ttk.Frame(self.tab2); btns2.pack(anchor="w", padx=8, pady=4)
        self.btn2_copy_art = ttk.Button(btns2, text="Kopiera artikelnr", command=self.copy_tab2_art, state="disabled")
        self.btn2_copy_art.pack(side="left", padx=4)
        self.btn2_copy_combo = ttk.Button(btns2, text="Kopiera artikelnr + namn", command=self.copy_tab2_combo, state="disabled")
        self.btn2_copy_combo.pack(side="left", padx=4)
        self.btn2_details = ttk.Button(btns2, text="Läs mer", command=self.details_tab2, state="disabled")
        self.btn2_details.pack(side="left", padx=4)
        self.btn2_export = ttk.Button(btns2, text="Exportera (csv)", command=self.export_tab2, state="disabled")
        self.btn2_export.pack(side="left", padx=4)

        # Data containers
        self.complete_orders: list[str] = []
        self.holistic_short = None  # pandas DataFrame när den finns

        # ---- NYTT: flikar för "1 rad & N beställt" ----
        self.order_tabs = []  # list med metadata för generiska flikar

        def add_order_tab(title: str, attr: str, default_filename: str):
            frame = ttk.Frame(self.frame_lists)
            self.frame_lists.add(frame, text=title)
            lbl = ttk.Label(frame, text="Ingen data än.")
            lbl.pack(anchor="w", padx=8, pady=8)
            btns = ttk.Frame(frame); btns.pack(anchor="w", padx=8, pady=4)
            btn_copy = ttk.Button(btns, text="Kopiera ordernummer",
                                  command=lambda a=attr: self._copy_order_list(a), state="disabled")
            btn_copy.pack(side="left", padx=4)
            btn_details = ttk.Button(btns, text="Läs mer",
                                     command=lambda a=attr: self._details_order_list(a), state="disabled")
            btn_details.pack(side="left", padx=4)
            btn_export = ttk.Button(btns, text="Exportera (txt)",
                                    command=lambda a=attr, fn=default_filename: self._export_order_list(a, fn),
                                    state="disabled")
            btn_export.pack(side="left", padx=4)
            self.order_tabs.append({
                "attr": attr,
                "label": lbl,
                "btn_copy": btn_copy,
                "btn_details": btn_details,
                "btn_export": btn_export,
                "filename": default_filename,
            })

        # Listattribut för de fyra nya listorna (1×1, 1×2, 1×3, 1×4).
        self.orders_1x1: list[str] = []
        self.orders_1x2: list[str] = []
        self.orders_1x3: list[str] = []
        self.orders_1x4: list[str] = []

        add_order_tab("Ordrar: 1 rad & 1 beställt", "orders_1x1", "ordrar_1rad_1st.txt")
        add_order_tab("Ordrar: 1 rad & 2 beställt", "orders_1x2", "ordrar_1rad_2st.txt")
        add_order_tab("Ordrar: 1 rad & 3 beställt", "orders_1x3", "ordrar_1rad_3st.txt")
        add_order_tab("Ordrar: 1 rad & 4 beställt", "orders_1x4", "ordrar_1rad_4st.txt")

        # Drag & drop
        if _DND_AVAILABLE:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._on_drop_files)
            except Exception:
                pass

        # ----- Ny knapp för att rensa all data -----
        # Lägg till en knapp i överdelen som återställer allt till startskick.
        # Den placeras bredvid "Analysera igen" så att användaren enkelt kan börja om.
        ttk.Button(top, text="Rensa", command=self.reset_state).grid(row=2, column=1, padx=6, pady=6, sticky="w")

    # ---------- DnD ----------
    def _on_drop_files(self, event):
        try:
            paths = self.tk.splitlist(event.data)
            if not paths:
                return
            path = paths[0]
            if os.path.isdir(path):
                return
            self.load_csv_path(path)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Fel vid drag & släpp:\n{e}")

    # ---------- Filinläsning ----------
    def load_csv_dialog(self):
        path = filedialog.askopenfilename(
            title="Välj CSV/TXT",
            filetypes=[("CSV/text", "*.csv *.txt"), ("Alla filer", "*.*")]
        )
        if not path:
            return
        self.load_csv_path(path)

    def load_csv_path(self, path: str):
        try:
            df = read_csv_flex(path)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Kunde inte läsa filen:\n{e}")
            return

        self.df = df
        self.path_var.set(path)

        # Autokartlägg kolumner
        try:
            self.mapping = auto_map_columns(df)
        except Exception as e:
            self.map_label.config(text="Kolumner: FEL – " + str(e))
            messagebox.showerror(APP_TITLE, str(e))
            return

        # Visa sammanfattning och kör analys direkt
        m = self.mapping
        self.map_label.config(
            text=("Kolumner: "
                  f"Order={m['order']} | Artikel={m['article']} | Plock={m['pick']} | "
                  f"Beställt={m['demand']} | Plockat={m.get('pickedqty') or '–'} | "
                  f"Namn={m.get('name') or '–'} | Plockplats={m.get('pickloc') or '–'}")
        )

        messagebox.showinfo(APP_TITLE, f"Inläst {os.path.basename(path)} med {len(df)} rader.")
        self.analyze()

    # ---------- Analys (UPPDATERAD) ----------
    def analyze(self):
        if self.df is None or not self.mapping:
            messagebox.showwarning(APP_TITLE, "Öppna en CSV först.")
            return

        pd = _lazy_pd()
        df = self.df.copy()
        m = self.mapping

        order_col   = m["order"]
        article_col = m["article"]
        pick_col    = m["pick"]           # kvantitet tillgängligt (Plock)
        demand_col  = m["demand"]         # Beställt
        name_col    = m.get("name")
        picked_col  = m.get("pickedqty")  # LÄSES IN MEN ANVÄNDS INTE
        pickloc_col = m.get("pickloc")    # för SK-filter

        # Typer
        df[order_col]   = df[order_col].astype(str)
        df[article_col] = df[article_col].astype(str)
        if name_col:
            df[name_col] = df[name_col].astype(str)
        if pickloc_col and pickloc_col in df.columns:
            df[pickloc_col] = df[pickloc_col].astype(str)

        df[pick_col]   = to_numeric_safe(df[pick_col])
        df[demand_col] = to_numeric_safe(df[demand_col])
        if picked_col and picked_col in df.columns:
            # Vi läser in den men använder den inte i analyslogiken
            df[picked_col] = to_numeric_safe(df[picked_col])

        # --- Viktigt: "rad OK" = BARA saldo räcker (Plock ≥ Beställt) ---
        df["_enough_row"] = df[pick_col] >= df[demand_col]

        # Lista 1 – kompletta ordrar (alla rader i ordern har tillräckligt saldo)
        complete_mask = df.groupby(order_col)["_enough_row"].all()
        self.complete_orders = sorted(complete_mask[complete_mask].index.astype(str).tolist())

        # Helheten per artikel → Lista 2 (oförändrad mot din struktur)
        demand_by_art = df.groupby(article_col)[demand_col].sum(min_count=1)
        stock_by_art  = df.groupby(article_col)[pick_col].max()

        holistic = pd.DataFrame({
            "Total beställt": demand_by_art,
            "Tillgängligt saldo (Plock)": stock_by_art
        }).fillna(0)
        holistic["Underskott"] = (holistic["Total beställt"] - holistic["Tillgängligt saldo (Plock)"]).clip(lower=0)
        holistic["Finns saldo men inte för allt"] = holistic["Underskott"] > 0

        if name_col and name_col in df.columns:
            first_names = df.drop_duplicates(subset=[article_col]).set_index(article_col)
            holistic["Benämning"] = first_names.get(name_col, "")
        else:
            holistic["Benämning"] = ""

        orders_per_article = df.groupby(article_col)[order_col].apply(lambda s: sorted(pd.unique(s.astype(str)))).rename("Påverkade ordrar")
        holistic = holistic.join(orders_per_article, how="left")

        holistic_short = holistic[holistic["Finns saldo men inte för allt"]].copy()
        holistic_short.index.name = "Artikel"
        self.holistic_short = holistic_short

        # --- NY ANALYS: Ordrar med exakt 1 rad & N beställt, OCH saldo räcker ---
        # Konsolidera (Order, Artikel): summera Beställt, och ta max Plock (saldo) för raden.
        oap = df.groupby([order_col, article_col]).agg(
            demand_sum=(demand_col, "sum"),
            pick_max=(pick_col, "max")
        )
        # En rad räcker om pick_max ≥ demand_sum
        oap["enough"] = oap["pick_max"] >= oap["demand_sum"]

        # Per order:
        lines_per_order   = oap.groupby(level=0).size()                     # antal (Order,Artikel)-rader
        qty_per_order     = oap.groupby(level=0)["demand_sum"].sum()        # total beställt i ordern
        enough_per_order  = oap.groupby(level=0)["enough"].all()            # alla rader har tillräckligt

        # Räkna ordrar med exakt 1 rad och specifikt beställt antal. Inget SK-filter längre.
        def orders_1line_qty(n: int) -> list[str]:
            cond = (lines_per_order == 1) & (qty_per_order == n) & (enough_per_order)
            return sorted(cond[cond].index.astype(str).tolist())

        self.orders_1x1 = orders_1line_qty(1)
        self.orders_1x2 = orders_1line_qty(2)
        self.orders_1x3 = orders_1line_qty(3)
        self.orders_1x4 = orders_1line_qty(4)

        # Filtrera bort ordrar som innehåller en artikel i den förbjudna listan
        # eller där artiklarna inte har tillräckligt lager för att täcka hela beställningsbehoven.
        restricted_articles = load_restricted_articles()
        # Artiklar som saknar täckning för hela behovet (globalt underskott)
        shortage_articles: set[str] = set()
        if self.holistic_short is not None and not self.holistic_short.empty:
            shortage_articles = set(self.holistic_short.index.astype(str))
        if restricted_articles or shortage_articles:
            # gruppera artiklar per order för att kunna filtrera
            articles_by_order = df.groupby(order_col)[article_col].apply(lambda s: set(s.astype(str)))
            for attr in ("orders_1x1", "orders_1x2", "orders_1x3", "orders_1x4"):
                lst: list[str] = getattr(self, attr, [])
                filtered: list[str] = []
                for o in lst:
                    arts = articles_by_order.get(o, set())
                    # hoppa över ordern om någon artikel är förbjuden
                    if restricted_articles and not arts.isdisjoint(restricted_articles):
                        continue
                    # hoppa över ordern om någon artikel har globalt underskott
                    if shortage_articles and not arts.isdisjoint(shortage_articles):
                        continue
                    filtered.append(o)
                setattr(self, attr, filtered)

        # --- Uppdatera UI ---
        # Tab 1
        self.lbl1.config(text=f"Kompletta ordrar: {len(self.complete_orders)} st")
        enable1 = "normal" if self.complete_orders else "disabled"
        self.btn1_copy.config(state=enable1)
        self.btn1_details.config(state=enable1)
        self.btn1_export.config(state=enable1)

        # Tab 2
        self.lbl2.config(text=f"Artiklar att beställa (helheten): {len(self.holistic_short)} st")
        enable2 = "normal" if len(self.holistic_short) else "disabled"
        self.btn2_copy_art.config(state=enable2)
        self.btn2_copy_combo.config(state=enable2)
        self.btn2_details.config(state=enable2)
        self.btn2_export.config(state=enable2)

        # Nya flikar (generiska)
        for meta in self.order_tabs:
            attr = meta["attr"]
            lst = getattr(self, attr, [])
            meta["label"].config(text=f"Ordrar: {len(lst)} st")
            state = "normal" if lst else "disabled"
            meta["btn_copy"].config(state=state)
            meta["btn_details"].config(state=state)
            meta["btn_export"].config(state=state)

        # Ingen SK-flik längre, så inget speciellt meddelande för saknad plockplats.

        messagebox.showinfo(APP_TITLE, "Analys klar.")

    # ---------- Tab 1 ----------
    def copy_tab1(self):
        copy_to_clipboard(self, "\n".join(self.complete_orders))
        messagebox.showinfo(APP_TITLE, "Ordernummer kopierade.")

    def details_tab1(self):
        if not self.complete_orders:
            return
        open_text_window(self, "Kompletta ordrar", self.complete_orders)

    def export_tab1(self):
        if not self.complete_orders:
            return
        pd = _lazy_pd()
        df = pd.DataFrame({"Order nr": self.complete_orders})
        save_df_dialog(self, df[["Order nr"]], "kompletta_ordrar.txt")

    # ---------- Tab 2 (helhet) ----------
    def copy_tab2_art(self):
        if self.holistic_short is None or self.holistic_short.empty:
            return
        arts = list(self.holistic_short.index.astype(str))
        copy_to_clipboard(self, "\n".join(arts))
        messagebox.showinfo(APP_TITLE, "Artikelnr kopierade.")

    def copy_tab2_combo(self):
        if self.holistic_short is None or self.holistic_short.empty:
            return
        lines = [f"{idx} – {name}".strip(" –") for idx, name in
                 zip(self.holistic_short.index.astype(str), self.holistic_short["Benämning"].astype(str))]
        copy_to_clipboard(self, "\n".join(lines))
        messagebox.showinfo(APP_TITLE, "Artikelnr + namn kopierade.")

    def details_tab2(self):
        if self.holistic_short is None or self.holistic_short.empty:
            return
        open_df_window(self, "Sammanfattning per artikel (helheten)", self.holistic_short.reset_index())

    def export_tab2(self):
        if self.holistic_short is None or self.holistic_short.empty:
            return
        save_df_dialog(self, self.holistic_short.reset_index(), "artiklar_att_bestalla.csv")

    # ---------- Generiska handlers för nya orderflikar ----------
    def _copy_order_list(self, attr: str):
        lst = getattr(self, attr, [])
        if not lst:
            return
        copy_to_clipboard(self, "\n".join(lst))
        messagebox.showinfo(APP_TITLE, "Ordernummer kopierade.")

    def _details_order_list(self, attr: str):
        lst = getattr(self, attr, [])
        if not lst:
            return
        title_map = {
            "orders_1x1": "Ordrar – 1 rad & 1 beställt",
            "orders_1x2": "Ordrar – 1 rad & 2 beställt",
            "orders_1x3": "Ordrar – 1 rad & 3 beställt",
            "orders_1x4": "Ordrar – 1 rad & 4 beställt",
        }
        open_text_window(self, title_map.get(attr, "Ordrar"), lst)

    def _export_order_list(self, attr: str, default_name: str):
        lst = getattr(self, attr, [])
        if not lst:
            return
        pd = _lazy_pd()
        df = pd.DataFrame({"Order nr": lst})
        save_df_dialog(self, df[["Order nr"]], default_name)

    # ---------- Återställning ----------
    def reset_state(self):
        """
        Återställ programmets tillstånd helt. Alla inlästa data och beräknade listor nollställs,
        sökvägen töms och alla knappar och etiketter återgår till ursprungsläget. Detta
        möjliggör att användaren kan börja om genom att ladda upp en ny fil utan att starta
        om applikationen.
        """
        # Töm data
        self.df = None
        self.mapping = {}

        # Töm filväg och kolumninfo
        self.path_var.set("")
        self.map_label.config(text="Kolumner: –")

        # Nollställ interna listor och DataFrames
        self.complete_orders = []
        self.holistic_short = None
        self.orders_1x1 = []
        self.orders_1x2 = []
        self.orders_1x3 = []
        self.orders_1x4 = []

        # Uppdatera flik 1
        self.lbl1.config(text="Ingen data än.")
        for btn in (self.btn1_copy, self.btn1_details, self.btn1_export):
            btn.config(state="disabled")

        # Uppdatera flik 2
        self.lbl2.config(text="Ingen data än.")
        for btn in (self.btn2_copy_art, self.btn2_copy_combo, self.btn2_details, self.btn2_export):
            btn.config(state="disabled")

        # Uppdatera de generiska orderflikarna
        for meta in self.order_tabs:
            meta["label"].config(text="Ingen data än.")
            meta["btn_copy"].config(state="disabled")
            meta["btn_details"].config(state="disabled")
            meta["btn_export"].config(state="disabled")

        # Meddela användaren
        messagebox.showinfo(APP_TITLE, "All data har rensats. Du kan nu ladda upp en ny fil.")


if __name__ == "__main__":
    app = App()
    try:
        # Liten UI-skalning för läsbarhet
        app.tk.call("tk", "scaling", 1.25)
    except Exception:
        pass
    app.mainloop()
