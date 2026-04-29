"""
Combined WMS Analyzer and GUI
=============================

This module combines the analysis engine and the graphical user interface
for the Warehouse Management System (WMS) flow analysis into a single
Python file.  It allows users to drag and drop all relevant CSV log
files, specify the purchase order number and article number, and view
a detailed report of received pallets, shipments to customers, saldo
remaining per pallet and buffer updates.

The analysis logic is provided by the ``WMSAnalyzerUpdated`` class,
which classifies each pallet's status (Ej inlagrad, Plockplats,
Inlagrad, Skickad) and detects buffer updates and new pallet chains.
The GUI logic is provided by the ``WMSSearchApp`` class, which
presents a drop zone, status indicators for each log file, input
fields for the purchase and article numbers, and a results display.

To run the application simply execute this script with Python 3.  If
the optional ``tkinterdnd2`` package is installed, drag and drop
functionality will be enabled; otherwise you can click the drop area
to select files manually.
"""

from __future__ import annotations

import argparse
import os
import shutil
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional, Tuple

import pandas as pd

# Optional drag‑and‑drop support via tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None  # type: ignore


class WMSAnalyzerUpdated:
    """Enhanced analyzer with pallet location classification and buffer tracking."""

    def __init__(self, data_path: Optional[str] = None) -> None:
        # Determine directory of CSV files
        if data_path is None:
            data_path = os.path.dirname(os.path.abspath(__file__))
        self.data_path = data_path

        # Find CSV filenames dynamically.  If a file isn't present it will be
        # set to None so the analyzer can continue with missing data.
        try:
            self.receive_file = self._find_file_prefix('v_ask_receive_log')
        except FileNotFoundError:
            self.receive_file = None
        try:
            self.booking_file = self._find_file_prefix('v_ask_booking_putaway')
        except FileNotFoundError:
            self.booking_file = None
        try:
            self.buffert_file = self._find_file_prefix('v_ask_article_buffertpallet')
        except FileNotFoundError:
            self.buffert_file = None
        try:
            self.trans_file = self._find_file_prefix('v_ask_trans_log')
        except FileNotFoundError:
            self.trans_file = None
        try:
            self.pick_file = self._find_file_prefix('v_ask_pick_log_full')
        except FileNotFoundError:
            self.pick_file = None
        # Correction (saldojustering) log is optional
        try:
            self.correct_file = self._find_file_prefix('v_ask_correct_log')
        except FileNotFoundError:
            self.correct_file = None

        # Load dataframes. Missing files will result in empty DataFrames.
        self.receive_df = self._load_csv(self.receive_file)
        self.booking_df = self._load_csv(self.booking_file)
        self.buffert_df = self._load_csv(self.buffert_file)
        self.trans_df = self._load_csv(self.trans_file)
        self.pick_df = self._load_csv(self.pick_file)
        # Load correction log if available
        self.correct_df = self._load_csv(self.correct_file)

        # Pre-clean some columns for performance
        self._prepare_dataframes()

    def _prepare_dataframes(self) -> None:
        """Preprocess key columns for faster lookups."""
        # Strip whitespace in key columns
        for df, cols in [
            (self.receive_df, ['Inköpsnr', 'Artikel', 'Pallid', 'Mottaget']),
            (self.booking_df, ['Pall nr', 'Inköpsnr']),
            (self.buffert_df, ['Pallid', 'Lagerplats']),
            (self.trans_df, ['Pallid', 'Till', 'Timestamp', 'Från']),
            (self.pick_df, ['Pallid', 'Artikelnr', 'Plockat', 'Ordernr']),
            (self.correct_df, ['Pallid', 'Antal', 'Anledning', 'Artikel']),
        ]:
            for col in cols:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()

        # Convert numeric fields to numeric where appropriate
        if 'Mottaget' in self.receive_df.columns:
            self.receive_df['Mottaget_num'] = pd.to_numeric(
                self.receive_df['Mottaget'].str.replace(',', '.'), errors='coerce'
            )
        if 'Plockat' in self.pick_df.columns:
            self.pick_df['Plockat_num'] = pd.to_numeric(
                self.pick_df['Plockat'].str.replace(',', '.'), errors='coerce'
            )
        # Numeric conversion for correction amounts
        if not self.correct_df.empty and 'Antal' in self.correct_df.columns:
            self.correct_df['Antal_num'] = pd.to_numeric(
                self.correct_df['Antal'].str.replace(',', '.'), errors='coerce'
            )

    def _find_file_prefix(self, prefix: str) -> str:
        """Find the first CSV file beginning with prefix in data_path."""
        for fname in os.listdir(self.data_path):
            if fname.startswith(prefix) and fname.endswith('.csv'):
                return os.path.join(self.data_path, fname)
        raise FileNotFoundError(f"No file starting with {prefix!r} in {self.data_path}")

    def _load_csv(self, filename: Optional[str]) -> pd.DataFrame:
        """Load a CSV file into a DataFrame.

        If the filename is ``None`` or does not exist, return an empty
        DataFrame.  Otherwise attempt to load the file using tab as the
        separator.  If loading fails for any reason, return an empty
        DataFrame instead of raising an exception.

        Parameters
        ----------
        filename:
            Path to the CSV file or ``None``.

        Returns
        -------
        pd.DataFrame
            Loaded DataFrame or empty DataFrame.
        """
        if not filename or not os.path.exists(filename):
            return pd.DataFrame()
        try:
            return pd.read_csv(filename, sep='\t', encoding='utf-8', dtype=str)
        except Exception:
            # Fall back to auto-separator detection
            try:
                return pd.read_csv(filename, dtype=str, sep=None, engine="python")
            except Exception:
                return pd.DataFrame()

    # Classification helper methods
    def _is_putaway(self, pallid: str) -> bool:
        """Check if the pallet is listed in the putaway booking table.

        This returns True only if the booking DataFrame is non‑empty
        and contains a column named 'Pall nr' that matches the given
        pall ID.  If the booking DataFrame is empty or does not have
        the expected column, the pallet is treated as not being in
        the putaway list.
        """
        if self.booking_df.empty or 'Pall nr' not in self.booking_df.columns:
            return False
        try:
            return not self.booking_df[self.booking_df['Pall nr'] == pallid].empty
        except KeyError:
            return False

    def _get_buffert_location(self, pallid: str) -> Optional[str]:
        """Return the location of the pallet in buffertpallet, or None if not found."""
        rows = self.buffert_df[self.buffert_df['Pallid'] == pallid]
        if not rows.empty:
            # Return first location string
            return rows.iloc[0]['Lagerplats']
        return None

    def _get_latest_trans_destination(self, pallid: str) -> Optional[str]:
        """
        Get the 'Till' field of the latest transaction for this pallet.

        If no transactions exist, return None.  Latest is determined
        by lexicographically comparing the Timestamp strings (assuming
        ISO-like format), or by row order if timestamps are equal/missing.
        """
        rows = self.trans_df[self.trans_df['Pallid'] == pallid]
        if rows.empty:
            return None
        # Sort by Timestamp descending if available
        if 'Timestamp' in rows.columns:
            # Coerce to datetime, ignoring errors; fallback to original string sort
            try:
                tmp = pd.to_datetime(rows['Timestamp'], errors='coerce')
                rows = rows.assign(_ts=tmp)
                rows = rows.sort_values(by='_ts', ascending=False, na_position='last')
            except Exception:
                pass
        return rows.iloc[0]['Till'] if 'Till' in rows.columns else None

    def _find_new_pallids_from_buffert_update(self, plock_loc: str, exclude: List[str]) -> List[str]:
        """
        Given a pick location (plockplats), find new pallet IDs created by buffer updates.

        A buffer update is detected as a transaction where the pallet moves
        from the given plock location to 'CRANE'.  We look up rows in
        the transaction log where `Från` equals `plock_loc`, `Till`
        equals 'CRANE' (case insensitive), and the pallet ID is not in
        the `exclude` list.  We return a list of new pallet IDs sorted
        by the earliest timestamp.
        """
        if 'Från' not in self.trans_df.columns or 'Till' not in self.trans_df.columns:
            return []
        loc_upper = plock_loc.upper() if plock_loc else ''
        rows = self.trans_df[
            (self.trans_df['Från'].str.upper() == loc_upper) &
            (self.trans_df['Till'].str.upper() == 'CRANE') &
            (~self.trans_df['Pallid'].isin(exclude))
        ].copy()
        if rows.empty:
            return []
        # Sort by timestamp ascending if available
        if 'Timestamp' in rows.columns:
            try:
                ts = pd.to_datetime(rows['Timestamp'], errors='coerce')
                rows = rows.assign(_ts=ts)
                rows = rows.sort_values('_ts', ascending=True, na_position='last')
            except Exception:
                pass
        new_ids = rows['Pallid'].dropna().unique().tolist()
        return new_ids

    def _classify_location(self, pallid: str) -> Tuple[str, Optional[str]]:
        """
        Determine the pallet's status and its relevant location string.

        Returns a tuple of (status, location) where status is one of
        "Ej inlagrad", "Plockplats", "Inlagrad", "Skickad" or
        "Okänt" and location is the raw location/destination string.
        """
        # Check putaway list first
        if self._is_putaway(pallid):
            # Could also pick up location from UTE later
            return ("Ej inlagrad", None)

        # Check buffer table for a location
        loc = self._get_buffert_location(pallid)
        if loc:
            loc_upper = loc.upper()
            # UTE implies not stored yet
            if loc_upper.startswith('UTE'):
                return ("Ej inlagrad", loc)
            # Check pick locations
            if (loc_upper.startswith('B') or loc_upper.startswith('L') or loc_upper.startswith('AA')):
                # exclude AA75/AA76 from being considered pick
                if not (loc_upper.startswith('AA75') or loc_upper.startswith('AA76')):
                    return ("Plockplats", loc)
            # Otherwise, it's a normal buffer location
            return ("Inlagrad", loc)

        # No location in buffer; check latest transaction
        latest_dest = self._get_latest_trans_destination(pallid)
        if latest_dest:
            dest_upper = latest_dest.upper()
            # Customer shipments
            if dest_upper.startswith('TO') or dest_upper.startswith('PR'):
                return ("Skickad", latest_dest)
            # UTE is still considered not stored
            if dest_upper.startswith('UTE'):
                return ("Ej inlagrad", latest_dest)
            # Pick location
            if (dest_upper.startswith('B') or dest_upper.startswith('L') or dest_upper.startswith('AA')):
                if not (dest_upper.startswith('AA75') or dest_upper.startswith('AA76')):
                    return ("Plockplats", latest_dest)
            # If nothing matches, treat as inlagrad/buffer (e.g. LI/HH etc.)
            return ("Inlagrad", latest_dest)
        # If completely missing information
        return ("Okänt", None)

    def _get_event_timestamps(self, pallid: str) -> Dict[str, str]:
        """
        Retrieve timestamps for various events related to a given pallet:

        - 'mottag'          : first receiving time from the receive log (Ändrad)
        - 'putaway'         : first putaway time from booking list (Ändrad)
        - 'buffert'         : first buffer timestamp (Datum/tid)
        - 'pick'            : first pick time from pick log (Datum)
        - 'saldojustering'  : first saldo adjustment time (Ändrad)
        - 'trans'           : first transaction timestamp (Timestamp)

        Returns a dictionary with keys mapping to string timestamps or empty
        strings if not found.
        """
        ts: Dict[str, str] = {}
        # receiving log timestamp
        if not self.receive_df.empty and 'Ändrad' in self.receive_df.columns:
            rcv = self.receive_df[self.receive_df['Pallid'] == pallid]
            if not rcv.empty:
                val = rcv['Ändrad'].iloc[0]
                ts['mottag'] = val
        # putaway list timestamp
        if not self.booking_df.empty and 'Ändrad' in self.booking_df.columns:
            put = self.booking_df[self.booking_df['Pall nr'] == pallid]
            if not put.empty:
                val = put['Ändrad'].iloc[0]
                ts['putaway'] = val
        # buffer pallet timestamp
        if not self.buffert_df.empty and 'Datum/tid' in self.buffert_df.columns:
            buf = self.buffert_df[self.buffert_df['Pallid'] == pallid]
            if not buf.empty:
                val = buf['Datum/tid'].iloc[0]
                ts['buffert'] = val
        # pick log timestamp
        if not self.pick_df.empty and 'Datum' in self.pick_df.columns:
            pk = self.pick_df[self.pick_df['Pallid'] == pallid]
            if not pk.empty:
                val = pk['Datum'].iloc[0]
                ts['pick'] = val
        # saldojustering timestamp
        if not self.correct_df.empty and 'Ändrad' in self.correct_df.columns:
            co = self.correct_df[self.correct_df['Pallid'] == pallid]
            if not co.empty:
                val = co['Ändrad'].iloc[0]
                ts['saldojustering'] = val
        # transaction log timestamp
        if not self.trans_df.empty and 'Timestamp' in self.trans_df.columns:
            tr = self.trans_df[self.trans_df['Pallid'] == pallid]
            if not tr.empty:
                val = tr['Timestamp'].iloc[0]
                ts['trans'] = val
        return ts

    def analyze(self, purchase_number: str, article_number: str) -> str:
        """
        Analyze the specified purchase and article, producing a detailed report
        with location classifications and buffer update handling.
        """
        purchase_number = str(purchase_number).strip()
        article_number = str(article_number).strip()

        recv_matches = self.receive_df[
            (self.receive_df['Inköpsnr'] == purchase_number) &
            (self.receive_df['Artikel'] == article_number)
        ]
        if recv_matches.empty:
            return (f"Inga mottagna pallar hittades för inköpsnr '{purchase_number}' "
                    f"och artikelnr '{article_number}'.")

        # Collect initial pallids and total quantity received on these palls
        initial_pall_ids = recv_matches['Pallid'].dropna().unique().tolist()
        # Build the full list of pallids by following potential buffer update chains
        pall_ids = list(initial_pall_ids)
        checked: set[str] = set(initial_pall_ids)
        i = 0
        while i < len(pall_ids):
            pid = pall_ids[i]
            status, loc = self._classify_location(pid)
            if status == "Plockplats" and loc:
                # Find new pallet IDs created by buffer update from this location
                new_ids = self._find_new_pallids_from_buffert_update(loc, list(checked))
                for nid in new_ids:
                    if nid not in checked:
                        pall_ids.append(nid)
                        checked.add(nid)
            i += 1
        # Sum total received only on the initial pallids (new palls may not be in receive log)
        total_received = recv_matches['Mottaget_num'].sum(skipna=True)
        # Aggregate pick quantities by order and by pallid for all pallids in chain
        pick_matches = self.pick_df[
            (self.pick_df['Pallid'].isin(pall_ids)) &
            (self.pick_df['Artikelnr'] == article_number)
        ]
        shipment_totals = pick_matches.groupby('Ordernr')['Plockat_num'].sum(min_count=1)
        total_shipped = shipment_totals.sum()

        # Prepare correction adjustments
        pall_corrections: Dict[str, float] = {}
        if not self.correct_df.empty:
            # Pallid-based adjustments for all palls in the chain
            corr_pall = self.correct_df[
                (~self.correct_df['Pallid'].str.lower().isin(['nan', '', 'none'])) &
                (self.correct_df['Pallid'].isin(pall_ids))
            ]
            if not corr_pall.empty:
                pall_corrections = corr_pall.groupby('Pallid')['Antal_num'].sum(min_count=1).to_dict()
            # General adjustments based on purchase number in Anledning
            corr_general = self.correct_df[self.correct_df['Anledning'] == purchase_number]
            general_adjustment = corr_general['Antal_num'].sum(skipna=True) if not corr_general.empty else 0.0
        else:
            general_adjustment = 0.0

        # Compute remaining quantities and classification per pall
        saldo_lines: List[str] = []
        for pallid in pall_ids:
            # Sum received, picked and correction quantities for this pallet
            recv_qty = recv_matches[recv_matches['Pallid'] == pallid]['Mottaget_num'].sum(skipna=True)
            picked_qty = pick_matches[pick_matches['Pallid'] == pallid]['Plockat_num'].sum(skipna=True)
            corr_qty = pall_corrections.get(pallid, 0.0)
            # Compute remaining (may be NaN or negative)
            remaining = None
            if pd.notnull(recv_qty) and pd.notnull(picked_qty) and pd.notnull(corr_qty):
                remaining = recv_qty - picked_qty + corr_qty
            status, loc = self._classify_location(pallid)
            loc_str = loc if loc is not None else ""
            # Include correction in output if exists
            corr_info = f", justerat: {corr_qty}" if corr_qty != 0 else ""
            origin_label = "inköp" if pallid in initial_pall_ids else "buffert"
            # For new pallet IDs (origin_label == 'buffert') with negative remaining, split into two lines:
            if origin_label == 'buffert' and remaining is not None and remaining < 0:
                # Determine quantity moved (positive value) and buffer update location
                ship_qty = picked_qty - recv_qty + corr_qty if pd.notnull(picked_qty) and pd.notnull(recv_qty) and pd.notnull(corr_qty) else None
                # Fallback: if ship_qty is still negative, take absolute value
                if ship_qty is not None:
                    try:
                        ship_qty_val = float(ship_qty)
                    except Exception:
                        ship_qty_val = None
                else:
                    ship_qty_val = None
                # Find buffer update location: look for first trans row where this pallet moves from a pick location to CRANE
                buffer_loc = None
                try:
                    # rows where this pallid has destination CRANE
                    buf_rows = self.trans_df[(self.trans_df['Pallid'] == pallid) & (self.trans_df['Till'].str.upper() == 'CRANE')]
                    if not buf_rows.empty and 'Från' in buf_rows.columns:
                        # pick the first row's 'Från' as the pick location
                        buffer_loc = buf_rows.iloc[0]['Från']
                except Exception:
                    buffer_loc = None
                buffer_loc_str = buffer_loc if buffer_loc is not None else ""
                # First line: buffer update event
                if ship_qty_val is not None:
                    line1 = (
                        f"   - PallID {pallid} [{origin_label}] : Buffertuppdatering (plats: {buffer_loc_str})"
                        f" - antal: {ship_qty_val}{corr_info}"
                    )
                else:
                    line1 = (
                        f"   - PallID {pallid} [{origin_label}] : Buffertuppdatering (plats: {buffer_loc_str})"
                        f" - antal: {picked_qty}{corr_info}"
                    )
                # Second line: shipped event with remaining set to 0.0
                line2 = (
                    f"   - PallID {pallid} [{origin_label}] : Skickad (plats: {loc_str})"
                    f" - kvar: 0.0{corr_info}"
                )
                # Append timestamps to second line
                ev = self._get_event_timestamps(pallid)
                if ev:
                    ev_parts = [f"{k}: {v}" for k, v in ev.items()]
                    line2 += "\n      • " + " | ".join(ev_parts)
                saldo_lines.append(line1)
                saldo_lines.append(line2)
            else:
                # Normal output line
                line = (
                    f"   - PallID {pallid} [{origin_label}] : {status} (plats: {loc_str})"
                    f" - kvar: {remaining}{corr_info}"
                )
                ev = self._get_event_timestamps(pallid)
                if ev:
                    ev_parts = [f"{k}: {v}" for k, v in ev.items()]
                    line += "\n      • " + " | ".join(ev_parts)
                saldo_lines.append(line)

        # Build report
        report_lines: List[str] = []
        report_lines.append(f"Analys för Inköpsnr: {purchase_number}, Artikelnr: {article_number}")
        report_lines.append("""
Flödesöversikt:
---------------
1. **Mottagna pallar:**
   - Antal mottagningsrader: {num_recv}
   - Unika pallID: {num_palls}
   - Summa mottaget antal: {total_recv}
2. **Levererat till kund:**
   - Summa plockat antal: {total_ship}
   - Orderfördelning:
""".format(num_recv=len(recv_matches), num_palls=len(pall_ids), total_recv=total_received, total_ship=total_shipped))

        if shipment_totals.empty:
            report_lines.append("   (Inga leveranser hittades i plockloggen för de matchande pallarna.)")
        else:
            for order, qty in shipment_totals.sort_index().items():
                report_lines.append(f"   - Order {order}: {qty}")

        # Remaining total
        # Compute total corrections (per pallid + general)
        total_pall_correction = sum(pall_corrections.values()) if pall_corrections else 0.0
        total_correction = (total_pall_correction + general_adjustment) if (pall_corrections or general_adjustment) else 0.0
        remaining_total = None
        if pd.notna(total_received) and pd.notna(total_shipped):
            try:
                remaining_total = total_received - total_shipped + total_correction
            except Exception:
                remaining_total = None
        if remaining_total is not None:
            report_lines.append(f"\n3. **Saldo kvar totalt:**\n   - Kvar att leverera: {remaining_total}")
        # Add detailed saldo breakdown
        report_lines.append("\n4. **Saldo per pallID:**")
        report_lines.extend(saldo_lines)
        # Add correction summary if applicable
        if total_correction != 0.0 or not pd.isna(general_adjustment):
            report_lines.append("\n5. **Saldojusteringar:**")
            # General adjustments
            if general_adjustment:
                report_lines.append(f"   - Justeringar på inköpsnr {purchase_number}: {general_adjustment}")
            # Pall specific adjustments
            for pid, adj in pall_corrections.items():
                report_lines.append(f"   - Justeringar på pallID {pid}: {adj}")

        return "\n".join(report_lines)

    # Optional CLI run method (unused by GUI but kept for completeness)
    def run(self, use_cli: bool) -> None:
        if use_cli:
            self._run_cli()
        else:
            self._run_gui()

    def _run_cli(self) -> None:
        print("WMS Analyzer (uppdaterad) – Kommandoradsgränssnitt\n")
        while True:
            try:
                purchase = input("Ange inköpsnummer (tom rad för att avsluta): ").strip()
                if not purchase:
                    print("Avslutar.")
                    break
                article = input("Ange artikelnr: ").strip()
                print("\nBearbetar...\n")
                result = self.analyze(purchase, article)
                print(result)
                print("\n" + "-" * 60 + "\n")
            except (EOFError, KeyboardInterrupt):
                print("\nAvslutar.")
                break

    def _run_gui(self) -> None:
        # Minimal GUI for CLI use; this method is unused in the combined file
        if tk is None:
            print("Tkinter är inte tillgängligt. Startar CLI istället.")
            self._run_cli()
            return
        root = tk.Tk()
        root.title("WMS Analyzer (uppdaterad)")
        root.geometry("700x500")
        ttk.Label(root, text="Inköpsnr:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        purchase_var = tk.StringVar()
        purchase_entry = ttk.Entry(root, textvariable=purchase_var, width=40)
        purchase_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        ttk.Label(root, text="Artikelnr:").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        article_var = tk.StringVar()
        article_entry = ttk.Entry(root, textvariable=article_var, width=40)
        article_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=10)
        result_text = tk.Text(root, wrap="word", height=20)
        result_text.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        root.grid_rowconfigure(3, weight=1)
        root.grid_columnconfigure(1, weight=1)
        def on_analyze() -> None:
            purchase = purchase_var.get().strip()
            article = article_var.get().strip()
            if not purchase or not article:
                messagebox.showwarning("Fel", "Både inköpsnr och artikelnr måste fyllas i.")
                return
            result = self.analyze(purchase, article)
            result_text.delete('1.0', tk.END)
            result_text.insert(tk.END, result)
        analyze_button = ttk.Button(root, text="Analysera", command=on_analyze)
        analyze_button.grid(row=2, column=0, columnspan=2, pady=5)
        root.mainloop()


def _parse_dnd_paths(event_data: str) -> List[str]:
    """Parse the raw ``event.data`` from a DND drop into a list of paths.

    The ``<<Drop>>`` event provides a string that may contain one or
    more file paths separated by whitespace.  Paths with spaces are
    wrapped in braces or quotes.  This helper splits the string and
    removes any surrounding braces or quotes.

    Parameters
    ----------
    event_data:
        Raw data string from the DND event.

    Returns
    -------
    List[str]
        A list of absolute file paths.
    """
    if not event_data:
        return []
    raw = str(event_data).strip()
    paths: List[str] = []
    current: List[str] = []
    # Split on whitespace but keep grouped braces together
    for token in raw.split():
        token_stripped = token.strip()
        # Start of a brace/quoted path
        if token_stripped.startswith('{') and not token_stripped.endswith('}'):
            current.append(token_stripped.lstrip('{'))
            continue
        if token_stripped.endswith('}') and current:
            current.append(token_stripped.rstrip('}'))
            paths.append(" ".join(current))
            current = []
            continue
        if current:
            # Middle of a multi‑word path
            current.append(token_stripped)
            continue
        # Single token path; strip surrounding braces/quotes
        if token_stripped.startswith('{') and token_stripped.endswith('}'):
            paths.append(token_stripped[1:-1])
        elif token_stripped.startswith('"') and token_stripped.endswith('"'):
            paths.append(token_stripped[1:-1])
        else:
            paths.append(token_stripped)
    return paths


class WMSSearchApp:
    """Main application class for the WMS search GUI."""

    def __init__(self, master: tk.Misc) -> None:
        # Store the root window
        self.master = master
        self.master.title("WMS Search")
        self.master.geometry("800x600")
        # Track loaded file paths; keys correspond to prefixes used by
        # WMSAnalyzerUpdated
        self.file_map: Dict[str, Optional[str]] = {
            'receive': None,
            'booking': None,
            'buffert': None,
            'trans': None,
            'pick': None,
            'correct': None,
        }
        # Human readable names for status display
        self.file_labels: Dict[str, str] = {
            'receive': 'Mottagningslogg',
            'booking': 'Ej inlagrade',
            'buffert': 'Buffertpall',
            'trans': 'Translogg',
            'pick': 'Plocklogg',
            'correct': 'Saldojustering',
        }
        # Create UI
        self._create_widgets()

    def _create_widgets(self) -> None:
        # Frame for drop area and file statuses
        container = ttk.Frame(self.master)
        container.pack(fill='both', expand=True, padx=10, pady=10)

        # Drop zone
        self.drop_label = ttk.Label(
            container,
            text="Drag och släpp alla loggfiler här",
            relief="groove",
            padding=20,
            anchor='center'
        )
        self.drop_label.pack(fill='x', pady=(0, 10))

        # Bind drop events if TkinterDnD is available
        if TkinterDnD and DND_FILES:
            try:
                # The master must be a TkinterDnD.Tk if DND is supported
                self.master.drop_target_register(DND_FILES)
                self.drop_label.drop_target_register(DND_FILES)
                self.drop_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass
        else:
            # If drag & drop is unavailable, allow file picking via dialog
            self.drop_label.bind("<Button-1>", self._pick_files_via_dialog)

        # Status panel
        status_frame = ttk.Frame(container)
        status_frame.pack(fill='x', pady=(0, 10))
        self.status_labels: Dict[str, ttk.Label] = {}
        row = 0
        for key in ['receive', 'booking', 'buffert', 'trans', 'pick', 'correct']:
            label_text = f"{self.file_labels[key]}"
            lbl = ttk.Label(status_frame, text=f"{label_text} (X)", foreground='red')
            lbl.grid(row=row, column=0, sticky='w', padx=5, pady=2)
            self.status_labels[key] = lbl
            row += 1

        # Input fields for purchase and article
        input_frame = ttk.Frame(container)
        input_frame.pack(fill='x', pady=(0, 10))
        ttk.Label(input_frame, text="Inköpsnr:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.purchase_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.purchase_var, width=30).grid(row=0, column=1, sticky='w', padx=5, pady=5)
        ttk.Label(input_frame, text="Artikelnr:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.article_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.article_var, width=30).grid(row=1, column=1, sticky='w', padx=5, pady=5)

        # Analyse button and clear button container
        button_frame = ttk.Frame(container)
        button_frame.pack(pady=(0, 10))
        # Analyse button – disabled until at least the receive log is loaded
        self.analyze_button = ttk.Button(button_frame, text="Analysera", command=self._run_analysis, state='disabled')
        self.analyze_button.pack(side='left', padx=5)
        # Clear button – resets all loaded files and form fields
        self.clear_button = ttk.Button(button_frame, text="Rensa", command=self._clear_all)
        self.clear_button.pack(side='left', padx=5)

        # Result text
        self.result_text = tk.Text(container, wrap='word')
        self.result_text.pack(fill='both', expand=True)

    def _pick_files_via_dialog(self, _event: Optional[tk.Event] = None) -> None:
        """Fallback when drag‑and‑drop is unavailable: use file dialog to pick files."""
        file_paths = filedialog.askopenfilenames(title="Välj loggfiler (CSV)", filetypes=[("CSV", "*.csv"), ("Alla filer", "*.*")])
        if file_paths:
            self._process_dropped_files(list(file_paths))

    def _on_drop(self, event: tk.Event) -> None:
        """Handle file drop event: parse and process the dropped files."""
        try:
            paths = _parse_dnd_paths(event.data)
            if paths:
                self._process_dropped_files(paths)
        except Exception as e:
            messagebox.showerror("Fel", f"Misslyckades med att läsa släppta filer: {e}")

    def _process_dropped_files(self, paths: List[str]) -> None:
        """Process a list of dropped file paths, attempting to classify each one.

        When all required files have been identified, enable the analyse
        button.
        """
        for path in paths:
            if not os.path.isfile(path):
                continue
            ftype = self._guess_file_type(path)
            if not ftype:
                # Unknown file; attempt to guess by header
                ftype = self._guess_file_type_by_header(path)
            if ftype:
                # Keep only the first file of each type; later drops override
                self.file_map[ftype] = path
                # Update status indicator
                lbl = self.status_labels[ftype]
                lbl.config(text=f"{self.file_labels[ftype]} (✓)", foreground='green')
        # Enable analyse button if at least the receiving log is present
        if self.file_map['receive']:
            self.analyze_button.config(state='normal')
        else:
            self.analyze_button.config(state='disabled')

    def _guess_file_type(self, path: str) -> Optional[str]:
        """Guess the file type based on its filename.

        Returns the key used in ``file_map`` if the filename suggests it
        corresponds to a known log.  Otherwise returns ``None``.
        """
        name = os.path.basename(path).lower()
        if 'receive' in name or 'mottagn' in name:
            return 'receive'
        if 'booking' in name or 'putaway' in name or 'ejinlagrad' in name or 'ej_inlagrad' in name:
            return 'booking'
        if 'buffert' in name and 'pallet' in name or 'buffertpallet' in name:
            return 'buffert'
        if 'trans' in name and 'log' in name:
            return 'trans'
        if 'pick' in name and 'log' in name:
            return 'pick'
        if 'correct' in name or 'saldo' in name or 'just' in name:
            return 'correct'
        return None

    def _guess_file_type_by_header(self, path: str) -> Optional[str]:
        """Guess the file type by inspecting column headers in the CSV.

        This function reads a small sample of the file to infer which
        log type it is.  Because the analyzer expects tab‑separated files,
        but some may be semicolon‑separated, we attempt to let Pandas
        auto‑detect the separator.
        """
        try:
            df = pd.read_csv(path, dtype=str, sep=None, engine="python", nrows=5)
        except Exception:
            try:
                df = pd.read_csv(path, dtype=str, sep='\t', engine="python", nrows=5)
            except Exception:
                return None
        cols = [c.lower() for c in df.columns]
        def has_all(names: List[str]) -> bool:
            return all(any(name in col for col in cols) for name in names)
        # Receive log: Inköpsnr, Artikel, Pallid
        if has_all(['inköpsnr', 'artikel', 'pallid']):
            return 'receive'
        # Booking / putaway: Pall nr, Inköpsnr
        if has_all(['pall nr', 'pallnr', 'pallid']) and any('inköpsnr' in col for col in cols):
            return 'booking'
        # Buffertpallet: Lagerplats, Pallid
        if has_all(['lagerplats', 'pallid']):
            return 'buffert'
        # Translogg: Till, Från, Pallid
        if has_all(['till', 'från', 'pallid']):
            return 'trans'
        # Pick log: Plockat, Ordernr, Pallid
        if has_all(['plockat', 'ordernr', 'pallid']):
            return 'pick'
        # Correct log: Anledning, Antal
        if has_all(['anledning', 'antal']):
            return 'correct'
        return None

    def _run_analysis(self) -> None:
        """Copy the loaded files to a temporary directory and run the analysis."""
        purchase = self.purchase_var.get().strip()
        article = self.article_var.get().strip()
        if not purchase or not article:
            messagebox.showwarning("Fel", "Både inköpsnr och artikelnr måste fyllas i.")
            return
        # Ensure at least the receive log exists; other logs are optional
        if not self.file_map['receive']:
            messagebox.showwarning("Fel", "Du måste åtminstone ladda mottagningsloggen (receive) för att kunna analysera.")
            return
        # Create a temporary directory to stage the files
        with tempfile.TemporaryDirectory() as tmpdir:
            # Define target names matching prefixes expected by WMSAnalyzerUpdated
            name_map = {
                'receive': 'v_ask_receive_log.csv',
                'booking': 'v_ask_booking_putaway.csv',
                'buffert': 'v_ask_article_buffertpallet.csv',
                'trans': 'v_ask_trans_log.csv',
                'pick': 'v_ask_pick_log_full.csv',
                'correct': 'v_ask_correct_log.csv',
            }
            # Copy each loaded file into the temp directory with the expected name
            for key, src in self.file_map.items():
                if src:
                    dst = os.path.join(tmpdir, name_map[key])
                    try:
                        shutil.copy(src, dst)
                    except Exception as e:
                        messagebox.showerror("Fel", f"Kunde inte kopiera filen '{src}': {e}")
                        return
            # Initialise the analyzer using the temporary directory
            try:
                analyzer = WMSAnalyzerUpdated(data_path=tmpdir)
                result = analyzer.analyze(purchase, article)
            except Exception as e:
                messagebox.showerror("Fel", f"Ett fel uppstod under analysen: {e}")
                return
        # Display the result
        self.result_text.delete('1.0', tk.END)
        self.result_text.insert(tk.END, result)

    def _clear_all(self) -> None:
        """Reset the UI and internal state to allow fresh uploads and analysis.

        This method clears the file mapping, resets status indicators to
        red "X", disables the analyse button, clears the purchase and
        article number fields, and empties the results text area.  It
        allows users to load new files without restarting the program.
        """
        # Reset file map
        for key in self.file_map:
            self.file_map[key] = None
        # Reset status labels
        for key, lbl in self.status_labels.items():
            lbl.config(text=f"{self.file_labels[key]} (X)", foreground='red')
        # Disable analyse button
        self.analyze_button.config(state='disabled')
        # Clear entry fields
        self.purchase_var.set("")
        self.article_var.set("")
        # Clear results
        self.result_text.delete('1.0', tk.END)


def main() -> None:
    """Entry point for the combined WMS search application."""
    if TkinterDnD:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = WMSSearchApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()