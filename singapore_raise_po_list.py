import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------------
MONTH_ORDER = {
    "january": 1, "jan": 1, "february": 2, "feb": 2, "march": 3, "mar": 3,
    "april": 4, "apr": 4, "may": 5, "june": 6, "jun": 6, "july": 7, "jul": 7,
    "august": 8, "aug": 8, "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10, "november": 11, "nov": 11, "december": 12, "dec": 12,
}

# Start column (1-based) of each category block in the source monthly sheets
CATEGORIES = [
    ("Monthly", 1),
    ("Prepaid", 4),
    ("Blanket", 7),
    ("Direct Purchase", 10),
    ("Unsorted", 13),
]

FIRST_DATA_ROW = 6
ACCOUNT_RE = re.compile(r"^BBAP-\d+[A-Za-z]?$", re.IGNORECASE)

WAITING_PATTERNS = [
    r"wait\w*\s*(for\s*)?(raise|rasise|issue)?\s*po",
    r"await\w*\s*(for\s*)?(raise|rasise|issue)?\s*po",
    r"(raise|rasise|rise)\s*(new\s*)?po",
    r"issue\s*po",
    r"for\s*po\b",
    r"generate\s*po",
    r"awaiting\s*customer",
]

# Output layout: 8 columns per category + 1 spacer column
OUT_HEADERS = [
    "Account ID", "Latest Waiting Month", "Waiting Status", "PIC",
    "Current Status ({latest})", "Latest Billed", "Chain Summary", "History Trace (newest -> oldest)",
]
OUT_WIDTHS = [12, 18, 34, 8, 26, 26, 46, 80]
GROUP_COLS = len(OUT_HEADERS)
GROUP_STRIDE = GROUP_COLS + 1


# ----------------------------------------------------------------------------
# Status helpers
# ----------------------------------------------------------------------------
def parse_month_order(sheet_name: str) -> int:
    """'Oct 2026' / 'February  2026' -> 202610 / 202602. Non-month sheets -> year*100 + 0."""
    tokens = re.findall(r"[a-zA-Z]+|\d+", sheet_name.lower())
    month_num, year_num = 0, 2026
    for t in tokens:
        if t in MONTH_ORDER:
            month_num = MONTH_ORDER[t]
        elif t.isdigit() and len(t) == 4:
            year_num = int(t)
    return year_num * 100 + month_num


def is_done_status(status: str) -> bool:
    sl = status.lower().strip()
    if not sl:
        return False
    return sl.startswith("done") or "free done" in sl or "no further actions" in sl


def is_waiting_po_status(status: str) -> bool:
    sl = status.lower().strip()
    if not sl or is_done_status(sl):
        return False
    return any(re.search(p, sl) for p in WAITING_PATTERNS)


def is_billed_status(status: str) -> bool:
    """'Billed.', 'Billed and sent', 'Billed.Waiting raise PO', 'Liyana Billed.' ... (Billed is NOT Done)."""
    sl = status.lower()
    return bool(re.search(r"\bbilled\b", sl)) and not is_done_status(sl)


def account_sort_key(acc: str):
    m = re.match(r"^BBAP-(\d+)(.*)$", acc, re.IGNORECASE)
    return (int(m.group(1)), m.group(2)) if m else (10**9, acc)


def put(ws, row, col, value, **style):
    """Write a cell; text starting with '=' is stored as plain text, never as a formula."""
    cell = ws.cell(row, col)
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"
    for k, v in style.items():
        setattr(cell, k, v)
    return cell


# ----------------------------------------------------------------------------
# Core logic
# ----------------------------------------------------------------------------
def collect_history(wb, month_sheets, log):
    """
    Returns {account: {month_order: record}}.
    One record per account per month. If an account is listed more than once in the
    same month, a Waiting-PO row wins, then a row with a non-empty status.
    """
    history, skipped = {}, set()
    for sheet_name in month_sheets:
        ws = wb[sheet_name]
        m_order = parse_month_order(sheet_name)
        for category, start_col in CATEGORIES:
            for row in range(FIRST_DATA_ROW, ws.max_row + 1):
                acc_val = ws.cell(row, start_col).value
                if acc_val is None or str(acc_val).strip() == "":
                    continue
                acc_id = str(acc_val).strip().upper()
                if not ACCOUNT_RE.match(acc_id):
                    skipped.add(f"{acc_id} ({sheet_name})")
                    continue
                status_val = ws.cell(row, start_col + 1).value
                pic_val = ws.cell(row, start_col + 2).value if (start_col + 2) <= ws.max_column else None
                rec = {
                    "month": sheet_name,
                    "mo": m_order,
                    "category": category,
                    "status": str(status_val).strip() if status_val is not None else "",
                    "pic": str(pic_val).strip() if pic_val is not None else "",
                }
                rank = 2 if is_waiting_po_status(rec["status"]) else (1 if rec["status"] else 0)
                by_month = history.setdefault(acc_id, {})
                old = by_month.get(m_order)
                if old is None or rank > old["rank"]:
                    rec["rank"] = rank
                    by_month[m_order] = rec
    if skipped:
        log(f"Skipped {len(skipped)} non-standard account ID(s): " + ", ".join(sorted(skipped)[:15])
            + (" ..." if len(skipped) > 15 else ""))
    return history


def analyse_account(acc_id, by_month, month_sheets, latest_mo):
    """
    Rules:
      1. Account must still be listed in the newest sheet (otherwise it is treated as closed).
      2. Take its most recent Waiting-PO month.
      3. If a Done appears in any LATER month, the cycle is finished -> skip.
         (Billed does NOT finish a cycle.)
      4. Walk back from that waiting month, count Waiting-PO months, stop at the first Done.
         No Done found -> reported as 'no Done found going back'.
    Returns a result dict, or None if the account is not pending.
    """
    if latest_mo not in by_month:
        return None

    recs = [by_month[m] for m in sorted(by_month, reverse=True)]  # newest -> oldest
    w_idx = next((i for i, r in enumerate(recs) if is_waiting_po_status(r["status"])), None)
    if w_idx is None:
        return None
    if any(is_done_status(r["status"]) for r in recs[:w_idx]):
        return None

    waiting = recs[w_idx]
    wait_count, last_done = 0, None
    for r in recs[w_idx:]:
        if is_done_status(r["status"]):
            last_done = r
            break
        if is_waiting_po_status(r["status"]):
            wait_count += 1

    billed = next((r for r in recs if is_billed_status(r["status"])), None)
    latest_billed = f"{billed['month']} ({billed['status']})" if billed else "No Billed record"

    stop_mo = last_done["mo"] if last_done else recs[-1]["mo"]
    steps = []
    for sheet in month_sheets:  # newest -> oldest
        mo = parse_month_order(sheet)
        if mo < stop_mo:
            break
        r = by_month.get(mo)
        if r is None:
            steps.append(f"{sheet}: (not listed)")
        else:
            pic = f" [{r['pic']}]" if r["pic"] else ""
            steps.append(f"{sheet}: {r['status'] or 'Empty'}{pic}")

    plural = "time" if wait_count == 1 else "times"
    if last_done:
        summary = (f"Waiting PO appeared {wait_count} {plural} since last Done "
                   f"({last_done['month']}: {last_done['status']})")
    else:
        summary = f"Waiting PO appeared {wait_count} {plural}; no Done found going back"

    return {
        "account": acc_id,
        "category": by_month[latest_mo]["category"],  # current category (newest sheet)
        "wait_month": waiting["month"],
        "wait_status": waiting["status"],
        "pic": waiting["pic"],
        "current_status": by_month[latest_mo]["status"],
        "latest_billed": latest_billed,
        "summary": summary,
        "trace": " <- ".join(steps),
        "no_done": last_done is None,
    }


def process_workbook(filepath, output_path, output_sheet_name="Pending_PO_Summary", log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)

    if os.path.normcase(os.path.abspath(filepath)) == os.path.normcase(os.path.abspath(output_path)):
        raise ValueError("Output file must be different from the source file (the source is never modified).")

    log("Loading workbook...")
    wb = openpyxl.load_workbook(filepath)

    month_sheets = [s for s in wb.sheetnames
                    if s != output_sheet_name and parse_month_order(s) % 100 != 0]
    if not month_sheets:
        raise ValueError("No monthly sheets (e.g. 'Oct 2026') found in this workbook.")
    month_sheets.sort(key=parse_month_order, reverse=True)
    latest_sheet = month_sheets[0]
    latest_mo = parse_month_order(latest_sheet)
    log(f"Found {len(month_sheets)} monthly sheet(s): {', '.join(month_sheets)}")
    log(f"Newest sheet = [{latest_sheet}]. Accounts not listed there are ignored.")

    history = collect_history(wb, month_sheets, log)
    in_latest = sum(1 for m in history.values() if latest_mo in m)
    log(f"Scanned {len(history)} account(s); {in_latest} listed in [{latest_sheet}].")

    results = {cat: [] for cat, _ in CATEGORIES}
    for acc_id in sorted(history, key=account_sort_key):
        res = analyse_account(acc_id, history[acc_id], month_sheets, latest_mo)
        if res:
            results[res["category"]].append(res)

    # ---- build output sheet -------------------------------------------------
    if output_sheet_name in wb.sheetnames:
        del wb[output_sheet_name]
    out = wb.create_sheet(title=output_sheet_name, index=0)

    title_fill = PatternFill(fill_type="solid", start_color="FFC000", end_color="FFC000")
    cat_fill = PatternFill(fill_type="solid", start_color="D9E1F2", end_color="D9E1F2")
    sub_fill = PatternFill(fill_type="solid", start_color="EDEDED", end_color="EDEDED")
    nodone_fill = PatternFill(fill_type="solid", start_color="FCE4D6", end_color="FCE4D6")
    bold = Font(bold=True)
    wrap_top = Alignment(wrap_text=True, vertical="top")
    top = Alignment(vertical="top")

    last_col = len(CATEGORIES) * GROUP_STRIDE - 1
    put(out, 1, 1, f"Sequencing Accounts - Pending Waiting PO Summary (as of {latest_sheet})",
        font=Font(bold=True, size=13), fill=title_fill)
    out.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    put(out, 2, 1, "Orange rows = Waiting PO but no Done found going back. "
                   "Accounts not listed in the newest sheet are excluded.", font=Font(italic=True))

    total, no_done_total = 0, 0
    for idx, (cat_name, _) in enumerate(CATEGORIES):
        base = idx * GROUP_STRIDE + 1
        items = results[cat_name]
        total += len(items)
        no_done_total += sum(1 for i in items if i["no_done"])
        log(f"  - {cat_name}: {len(items)} account(s)")

        put(out, 3, base, f"{cat_name} ({len(items)})", font=bold, fill=cat_fill,
            alignment=Alignment(horizontal="center"))
        out.merge_cells(start_row=3, start_column=base, end_row=3, end_column=base + GROUP_COLS - 1)

        for off, h in enumerate(OUT_HEADERS):
            put(out, 4, base + off, h.format(latest=latest_sheet), font=bold, fill=sub_fill)

        for r_off, it in enumerate(items):
            r = 5 + r_off
            values = [it["account"], it["wait_month"], it["wait_status"], it["pic"] or None,
                      it["current_status"] or None, it["latest_billed"], it["summary"], it["trace"]]
            for off, v in enumerate(values):
                c = put(out, r, base + off, v, alignment=wrap_top if off in (2, 4, 5, 6, 7) else top)
                if it["no_done"]:
                    c.fill = nodone_fill

        for off, w in enumerate(OUT_WIDTHS):
            out.column_dimensions[get_column_letter(base + off)].width = w
        if idx < len(CATEGORIES) - 1:
            out.column_dimensions[get_column_letter(base + GROUP_COLS)].width = 3

    out.freeze_panes = "A5"

    log(f"Saving new file: {output_path}")
    wb.save(output_path)  # source file is only read, never written
    log("Done!")
    return total, no_done_total, {k: len(v) for k, v in results.items()}


# ----------------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------------
class SequencingFilterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Sequencing Account PO Status Filter")
        self.root.geometry("700x520")
        self.root.resizable(False, False)

        self.filepath_var = tk.StringVar()
        self.sheetname_var = tk.StringVar(value="Pending_PO_Summary")
        self.output_var = tk.StringVar()
        self._build_ui()

    def _build_ui(self):
        main = tk.Frame(self.root, padx=20, pady=15)
        main.pack(fill=tk.BOTH, expand=True)

        tk.Label(main, text="Sequencing Account Summary - Waiting PO Filter",
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 12))

        file_frame = tk.LabelFrame(main, text="1. Select Excel File (.xlsx)", padx=10, pady=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Entry(file_frame, textvariable=self.filepath_var, width=64,
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 8), ipady=3)
        self.browse_btn = tk.Button(file_frame, text="Browse...", width=12,
                                    command=self.browse_file, font=("Segoe UI", 9))
        self.browse_btn.pack(side=tk.LEFT)

        out_frame = tk.LabelFrame(main, text="2. Output File (a new file; the source is never modified)",
                                  padx=10, pady=10)
        out_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Entry(out_frame, textvariable=self.output_var, width=64,
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 8), ipady=3)
        self.save_btn = tk.Button(out_frame, text="Save As...", width=12,
                                  command=self.choose_output, font=("Segoe UI", 9))
        self.save_btn.pack(side=tk.LEFT)

        action = tk.LabelFrame(main, text="3. Settings & Execution", padx=10, pady=10)
        action.pack(fill=tk.X, pady=(0, 10))
        tk.Label(action, text="Output Sheet Name:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        tk.Entry(action, textvariable=self.sheetname_var, width=24,
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(6, 12), ipady=3)
        self.run_btn = tk.Button(action, text="Run Filter", width=16, bg="#0078D7", fg="white",
                                 font=("Segoe UI", 9, "bold"), command=self.start_processing)
        self.run_btn.pack(side=tk.RIGHT)

        log_frame = tk.LabelFrame(main, text="Execution Log", padx=10, pady=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_box = scrolledtext.ScrolledText(log_frame, height=11, font=("Consolas", 9),
                                                 state="disabled", bg="#F9F9F9")
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def append_log(self, message: str):
        def _update():
            self.log_box.configure(state="normal")
            self.log_box.insert(tk.END, message + "\n")
            self.log_box.see(tk.END)
            self.log_box.configure(state="disabled")
        self.root.after(0, _update)

    def browse_file(self):
        path = filedialog.askopenfilename(title="Select Sequencing Account Summary Excel File",
                                          filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.filepath_var.set(path)
            base, ext = os.path.splitext(path)
            self.output_var.set(f"{base}_PO_Summary{ext}")  # default suggestion, editable
            self.append_log(f"Selected file: {path}")

    def choose_output(self):
        src = self.filepath_var.get().strip()
        path = filedialog.asksaveasfilename(
            title="Save Output Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialdir=os.path.dirname(src) if src else None,
            initialfile=os.path.basename(self.output_var.get()) or "PO_Summary.xlsx",
        )
        if path:
            self.output_var.set(path)

    def start_processing(self):
        filepath = self.filepath_var.get().strip()
        output_sheet = self.sheetname_var.get().strip() or "Pending_PO_Summary"
        output_path = self.output_var.get().strip()
        if not filepath:
            messagebox.showwarning("No File Selected", "Please select an Excel (.xlsx) file first.")
            return
        if not output_path:
            messagebox.showwarning("No Output File", "Please choose where to save the output file.")
            return
        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"
            self.output_var.set(output_path)
        if os.path.normcase(os.path.abspath(filepath)) == os.path.normcase(os.path.abspath(output_path)):
            messagebox.showwarning("Same File", "The output file must be different from the source file.")
            return
        if os.path.exists(output_path) and not messagebox.askyesno(
                "Overwrite?", f"{os.path.basename(output_path)} already exists.\nOverwrite it?"):
            return

        self.run_btn.configure(state="disabled", text="Processing...")
        self.browse_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled")
        self.append_log("-" * 55)
        threading.Thread(target=self._run_worker, args=(filepath, output_path, output_sheet),
                         daemon=True).start()

    def _run_worker(self, filepath, output_path, output_sheet):
        try:
            total, no_done, breakdown = process_workbook(
                filepath, output_path, output_sheet_name=output_sheet,
                log_callback=self.append_log)
            detail = "\n".join(f"• {k}: {v} account(s)" for k, v in breakdown.items())
            msg = (f"Saved new file:\n{output_path}\n\nSheet [{output_sheet}] generated.\n\n"
                   f"Accounts with a pending Waiting PO: {total}\n"
                   f"  (of which no Done found going back: {no_done})\n{detail}")
            self.root.after(0, lambda: messagebox.showinfo("Process Complete", msg))
        except PermissionError:
            self.append_log("ERROR: Permission denied. Close the output file if it is open in Excel, or pick another location.")
            self.root.after(0, lambda: messagebox.showerror(
                "Save Failed", "Cannot write the output file. It may be open in Excel, or the folder is "
                               "read-only. Close it or choose another location and try again."))
        except Exception as e:
            err = str(e)  # 'e' is deleted when the except block ends, so copy it for the lambda
            self.append_log(f"ERROR: {err}")
            self.root.after(0, lambda: messagebox.showerror(
                "Error", f"An unexpected error occurred:\n{err}"))
        finally:
            self.root.after(0, self._reset_buttons)

    def _reset_buttons(self):
        self.run_btn.configure(state="normal", text="Run Filter")
        self.browse_btn.configure(state="normal")
        self.save_btn.configure(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    SequencingFilterApp(root)
    root.mainloop()