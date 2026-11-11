from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import os
import queue
import re
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from playwright.sync_api import sync_playwright
import pandas as pd


# ---------------------------------------------------------------------------
# Parsing constants
# ---------------------------------------------------------------------------
MONEY_RE = re.compile(r"^\$?[\d,]*\.\d{2}$")
NUMBER_RE = re.compile(r"^[\d,]*\.?\d+$")
CSZ_RE = re.compile(
    r"^(?P<city>.+?)\s*,?\s+(?P<state>[A-Za-z]{2})\s*,?\s+(?P<zip>\d{5}(?:-\d{4})?)$"
)
ADDR_START_RE = re.compile(
    r"^(\d|P\.?\s*O\.?\s*BOX|POB\b|C/O|RR\b|HC\b|RURAL|PMB|ATTN)", re.IGNORECASE
)
STR_RE = re.compile(r"^\s*([^/\s]+)\s*/\s*([^/\s]+)\s*/\s*([^/\s]+)\s*$")
DETAIL_RE = re.compile(r"^([A-Za-z][A-Za-z /]*?):\s*(.*)$")
UNIT_RE = re.compile(
    r"^([A-Z0-9_ /&]+?)\s+-\s+(.+?)\s+-\s+(\d+(?:\.\d+)?)\s*%$"
)
EXTRA_CODE_RE = re.compile(r"^[A-Z]{1,5}\d{2,7}$")
ILLEGAL_XLSX_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

SESSION_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "scraper_session.json"
)

HEAD_COLUMNS = [
    "URL", "Status", "Owner", "Owner Address", "Owner City", "Owner State",
    "Owner Zip", "Appraised Value", "Assessed Value", "Tax Year",
    "Parcel Number", "PPIN", "Account Number", "Physical Address",
    "Deed Book/Page", "Utilities - Gas", "Utilities - Water",
    "Utilities - Sewer", "Amenities", "District", "Exempt Type", "Acres",
    "Tax Due", "Deed Date", "Road Access", "Topography", "Nuisances",
    "Historical", "Section/Township/Range", "Section", "Township", "Range",
    "Brief Property Description", "Improvement Count",
]
LAND_FIELDS = ["Current Use", "Type", "Acres", "Appraised", "Tax Class"]
LAND_COLUMNS = [f"Land {f}" for f in LAND_FIELDS]
TAX_COLUMNS = [
    "Recent Tax Year", "Recent Tax Paid", "Recent Tax Appraised",
    "Recent Tax Assessed", "Recent Tax Amount", "Recent Tax Fees",
    "Recent Tax Total",
]
JSON_COLUMNS = ["Payment History JSON", "Owner History JSON"]
IMP_FIELDS = [
    "Description", "Value", "Year Built", "Class", "Base Rate",
    "Adjusted Rate", "Total Depreciation", "Base Area", "Upper Floor Area",
    "Living Area", "Adjusted Area", "Tax Class",
]
IMP_COL_RE = re.compile(r"^Improvement (\d+)_(.+)$")
LAND_COL_RE = re.compile(r"^Land (\d+) (.+)$")


def tokenize(lines):
    """Split lines on newlines AND tabs; drop empties."""
    out = []
    for line in lines:
        for part in line.split("\t"):
            part = part.replace("\u00a0", " ").strip()
            if part:
                out.append(part)
    return out


class BlountPropertyScraperApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Blount County Property Batch Scraper - Fully Fixed")
        self.root.geometry("1150x750")
        self.root.minsize(900, 600)

        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar(
            value="blount_properties_output.xlsx"
        )
        self.max_workers_var = tk.IntVar(value=3)

        self.is_running = False
        self.stop_requested = False
        self.log_queue = queue.Queue()
        self.scraped_data_list = []
        self.executor = None

        self.create_widgets()
        self.restore_session()
        self.root.after(100, self.process_queues)

    def create_widgets(self):
        # --- Top Configuration Panel ---
        control_frame = ttk.LabelFrame(
            self.root, text=" Configuration & Controls ", padding=12
        )
        control_frame.pack(fill="x", padx=15, pady=10)

        ttk.Label(control_frame, text="Input CSV File:").grid(
            row=0, column=0, sticky="w", padx=5, pady=5
        )
        ttk.Entry(
            control_frame, textvariable=self.input_file_path, width=45
        ).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(
            control_frame, text="Browse...", command=self.browse_input_file
        ).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(control_frame, text="Parallel Workers:").grid(
            row=1, column=0, sticky="w", padx=5, pady=5
        )
        ttk.Spinbox(
            control_frame,
            from_=1,
            to=10,
            textvariable=self.max_workers_var,
            width=5,
        ).grid(row=1, column=1, sticky="w", padx=5, pady=5)

        # Action Buttons
        btn_action_frame = ttk.Frame(control_frame)
        btn_action_frame.grid(row=1, column=2, padx=5, pady=5, sticky="ew")

        self.start_btn = ttk.Button(
            btn_action_frame,
            text="🚀 Start",
            command=self.start_scraping_thread,
        )
        self.start_btn.pack(side="left", padx=2)

        self.resume_btn = ttk.Button(
            btn_action_frame,
            text="▶ Resume",
            command=lambda: self.start_scraping_thread(resume=True),
        )
        self.resume_btn.pack(side="left", padx=2)

        self.stop_btn = ttk.Button(
            btn_action_frame,
            text="🛑 Stop",
            command=self.stop_scraping,
            state="disabled",
        )
        self.stop_btn.pack(side="left", padx=2)

        self.export_btn = ttk.Button(
            btn_action_frame,
            text="📥 Export",
            command=self.manual_export,
            state="disabled",
        )
        self.export_btn.pack(side="left", padx=2)

        # --- Progress & Status ---
        progress_frame = ttk.Frame(self.root, padding=10)
        progress_frame.pack(fill="x", padx=15)

        self.progress_bar = ttk.Progressbar(
            progress_frame, orient="horizontal", mode="determinate"
        )
        self.progress_bar.pack(fill="x", pady=3)

        self.status_label = ttk.Label(
            progress_frame, text="Status: Ready to load CSV..."
        )
        self.status_label.pack(anchor="w")

        # --- Notebook for Table View & Logs ---
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=15, pady=10)

        # Tab 1: Live Table
        table_tab = ttk.Frame(notebook)
        notebook.add(table_tab, text="📊 Live Results Table")

        self.tree_scroll_y = ttk.Scrollbar(table_tab, orient="vertical")
        self.tree_scroll_x = ttk.Scrollbar(table_tab, orient="horizontal")

        self.tree = ttk.Treeview(
            table_tab,
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set,
        )
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.tree_scroll_y.pack(side="right", fill="y")
        self.tree_scroll_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        # Tab 2: Execution Log
        log_tab = ttk.Frame(notebook)
        notebook.add(log_tab, text="🛠️ Execution Log")

        self.log_text = tk.Text(
            log_tab,
            wrap="word",
            bg="#1e1e1e",
            fg="#d4d4d4",
            font=("Consolas", 10),
        )
        log_scroll = ttk.Scrollbar(
            log_tab, orient="vertical", command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="Select Input CSV",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
        )
        if filename:
            self.input_file_path.set(filename)
            self.refresh_resume_hint()

    def log(self, message):
        self.log_queue.put(message)

    def process_queues(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
        self.root.after(100, self.process_queues)

    # ------------------------------------------------------------------ #
    # Record template / scraping
    # ------------------------------------------------------------------ #
    @staticmethod
    def new_record(url):
        data = {c: None for c in HEAD_COLUMNS + LAND_COLUMNS + TAX_COLUMNS + JSON_COLUMNS}
        data["URL"] = url
        data["Status"] = "Failed"
        return data

    def scrape_property_page(self, url):
        data = self.new_record(url)
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                try:
                    page = browser.new_page()
                    page.goto(url, timeout=30000)
                    page.wait_for_selector("text=Property Details", timeout=10000)
                    body_text = page.inner_text("body")
                finally:
                    browser.close()
            self.parse_property_text(body_text, data)
            data["Status"] = "Success"
        except Exception as e:
            data["Status"] = f"Error: {str(e)}"
        return data

    # ------------------------------------------------------------------ #
    # Parsing (pure functions of the page text -> easy to test offline)
    # ------------------------------------------------------------------ #
    @staticmethod
    def field(text, label, value=r"[^\n\t]*"):
        """Value on the SAME line as the label only (never spills to the
        next line, which is what mixed Historical / Section-Township-Range)."""
        m = re.search(label + r"[ \t\u00a0]*(" + value + ")", text, re.IGNORECASE)
        if not m:
            return None
        return m.group(1).strip() or None

    def parse_property_text(self, body_text, data):
        f = self.field
        self.parse_owner(body_text, data)

        data["Appraised Value"] = f(body_text, r"Appraised:", r"[\d,\.]+")
        data["Assessed Value"] = f(body_text, r"Assessed:", r"[\d,\.]+")
        data["Tax Year"] = f(body_text, r"Tax Year:", r"\d{4}")
        data["Parcel Number"] = f(body_text, r"Parcel Number:", r"[\d\-\.]+")
        data["PPIN"] = f(body_text, r"PPIN:", r"\d+")
        data["Account Number"] = f(body_text, r"Account Number:", r"\d+")
        data["Physical Address"] = f(body_text, r"Physical Address:")
        data["Deed Book/Page"] = f(body_text, r"Deed Book/Page:")
        data["Utilities - Gas"] = f(body_text, r"Utilities\s*-\s*Gas:")
        data["Utilities - Water"] = f(body_text, r"Utilities\s*-\s*Water:")
        data["Utilities - Sewer"] = f(body_text, r"Utilities\s*-\s*Sewer:")
        data["Amenities"] = f(body_text, r"Amenities:")
        data["District"] = f(body_text, r"District:")
        data["Exempt Type"] = f(body_text, r"Exempt Type:")
        data["Acres"] = f(body_text, r"Acres:", r"[\d\.]+")
        data["Tax Due"] = f(body_text, r"Tax\s*Due:")
        data["Deed Date"] = f(body_text, r"Deed\s*Date:", r"[\d/]+")
        data["Road Access"] = f(body_text, r"Road Access:")
        data["Topography"] = f(body_text, r"Topography:")
        data["Nuisances"] = f(body_text, r"Nuisances:")
        data["Historical"] = f(body_text, r"Historical:")
        data["Brief Property Description"] = f(
            body_text, r"Brief Property Description For Tax Purposes Only:"
        )

        # Section / Township / Range (kept as TEXT + split into 3 columns)
        str_val = f(body_text, r"Section/Township/Range:")
        data["Section/Township/Range"] = str_val
        if str_val:
            m = STR_RE.match(str_val)
            if m:
                data["Section"], data["Township"], data["Range"] = m.groups()

        self.parse_improvements(body_text, data)
        self.parse_land(body_text, data)
        self.parse_recent_taxes(body_text, data)
        self.parse_histories(body_text, data)

    # ---- Owner ---------------------------------------------------------
    def parse_owner(self, body_text, data):
        lines = [l.strip() for l in body_text.split("\n")]
        start = next((i for i, l in enumerate(lines) if l == "OWNER"), None)
        if start is None:
            return
        block = []
        for l in lines[start + 1:]:
            if not l:
                continue
            if l in ("VALUE", "INFO") or re.match(r"^(Appraised|Assessed|Tax Year):", l):
                break
            block.append(l)
        if not block:
            return

        csz = CSZ_RE.match(block[-1]) if len(block) >= 2 else None
        if csz:
            data["Owner City"] = csz.group("city").strip()
            data["Owner State"] = csz.group("state").upper()
            data["Owner Zip"] = csz.group("zip")
            rest = block[:-1]
        else:
            rest = block

        # names come first; the address starts at the first street-looking line
        split_at = next(
            (i for i in range(1, len(rest)) if ADDR_START_RE.match(rest[i])), 1
        )
        data["Owner"] = " / ".join(rest[:split_at])
        addr = rest[split_at:]
        data["Owner Address"] = ", ".join(addr) if addr else None

    # ---- Improvements (any number, every field in its own column) -------
    @staticmethod
    def parse_improvement_cards(body_text):
        lines = [l.strip() for l in body_text.split("\n")]
        start = next((i for i, l in enumerate(lines) if l.upper() == "IMPROVEMENTS"), None)
        if start is None:
            return []
        end = next(
            (i for i in range(start + 1, len(lines)) if lines[i].upper() == "LAND"),
            len(lines),
        )
        tokens = tokenize(lines[start + 1:end])

        cards, cur, desc, mode, pending_code = [], None, [], "desc", None
        for t in tokens:
            up = t.upper()
            if up in ("DESCRIPTION", "FINAL VALUE"):
                continue
            if up == "BUILDING DETAILS":
                mode = "details"
                continue
            if up == "CONSTRUCTION UNITS":
                mode = "units"
                continue
            if up == "EXTRA FEATURES":
                mode, pending_code = "extras", None
                continue

            if mode == "desc":
                if MONEY_RE.match(t):
                    cur = {"Description": " ".join(desc), "Value": t, "details": {}, "units": {}, "extras": []}
                    cards.append(cur)
                    desc, mode = [], "details"
                else:
                    desc.append(t)
                continue

            if mode == "extras":
                if pending_code is not None:
                    # token right after a code = that feature's description
                    if EXTRA_CODE_RE.match(t):
                        cur["extras"].append((pending_code, ""))
                        pending_code = t
                    else:
                        cur["extras"].append((pending_code, t))
                        pending_code = None
                    continue
                if EXTRA_CODE_RE.match(t):
                    pending_code = t
                    continue
                # not a code -> first line of the NEXT improvement's description
                desc, mode = [t], "desc"
                continue

            # mode is "details" or "units": belongs to the current card
            um = UNIT_RE.match(t)
            if um:
                key, text, pct = um.group(1).strip(), um.group(2).strip(), um.group(3)
                val = f"{text} - {pct}%"
                prev = cur["units"].get(key)
                cur["units"][key] = f"{prev} | {val}" if prev else val
                mode = "units"
                continue
            dm = DETAIL_RE.match(t) if mode == "details" else None
            if dm:
                cur["details"][dm.group(1).strip()] = dm.group(2).strip()
                continue
            # anything else = first line of the NEXT improvement's description
            desc, mode = [t], "desc"
        if mode == "extras" and pending_code is not None:
            cur["extras"].append((pending_code, ""))
        return cards

    def parse_improvements(self, body_text, data):
        cards = self.parse_improvement_cards(body_text)
        data["Improvement Count"] = len(cards)
        for n, card in enumerate(cards, start=1):
            p = f"Improvement {n}_"
            data[p + "Description"] = card["Description"]
            data[p + "Value"] = card["Value"]
            for k, v in card["details"].items():
                data[p + k] = v
            for k, v in card["units"].items():
                data[p + k] = v
            for e, (code, text) in enumerate(card["extras"], start=1):
                data[f"{p}Extra {e}_Code"] = code
                data[f"{p}Extra {e}_Description"] = text

    # ---- Land (supports multiple land rows; Current Use may be blank) ----
    def parse_land(self, body_text, data):
        lines = [l.strip() for l in body_text.split("\n")]
        start = next((i for i, l in enumerate(lines) if l.upper() == "LAND"), None)
        if start is None:
            return
        end = next(
            (i for i in range(start + 1, len(lines))
             if lines[i].upper() in ("MOST RECENT TAXES", "PAYMENT HISTORY")),
            len(lines),
        )
        headers = {"CURRENT USE", "CURENT USE", "TYPE", "ACRES", "APPRAISED", "TAX CLASS"}
        tokens = [t for t in tokenize(lines[start + 1:end]) if t.upper() not in headers]

        rows, last, j = [], 0, 0
        while j + 2 < len(tokens):
            if (
                j - last >= 1
                and NUMBER_RE.match(tokens[j])
                and MONEY_RE.match(tokens[j + 1])
                and re.fullmatch(r"\d+", tokens[j + 2])
            ):
                rows.append(tokens[last:j + 3])
                last, j = j + 3, j + 3
            else:
                j += 1

        for n, row in enumerate(rows, start=1):
            cur_use = " ".join(row[:-4]) or None  # blank when the cell is empty
            vals = [cur_use, row[-4], row[-3], row[-2], row[-1]]
            for fname, v in zip(LAND_FIELDS, vals):
                key = f"Land {fname}" if n == 1 else f"Land {n} {fname}"
                data[key] = v

    # ---- Most recent taxes (unchanged logic) ----------------------------
    def parse_recent_taxes(self, body_text, data):
        tax_match = re.search(
            r"MOST RECENT TAXES\s*(.*?)(?=Payment History|$)", body_text, re.DOTALL
        )
        if not tax_match:
            return
        tax_lines = tokenize(tax_match.group(1).split("\n"))
        clean_tax = [
            l for l in tax_lines
            if l not in ["YEAR", "PAID", "APPRAISED", "ASSESSED", "TAXES", "FEES", "TOTAL"]
            and not l.startswith("View")
            and not l.startswith("Request")
            and not l.startswith("Assessed To:")
        ]
        if len(clean_tax) >= 7:
            for col, v in zip(TAX_COLUMNS, clean_tax[:7]):
                data[col] = v

    # ---- Payment / owner history (unchanged logic) -----------------------
    def parse_histories(self, body_text, data):
        pay = re.search(r"Payment History.*?(?=Owner History)", body_text, re.DOTALL)
        if pay:
            lines = [l.strip() for l in pay.group(0).split("\n") if l.strip()][4:]
            payments, i = [], 0
            while i < len(lines) - 4:
                if lines[i].isdigit() and len(lines[i]) == 4:
                    payments.append({
                        "Tax Year": lines[i], "Assessed Value": lines[i + 1],
                        "Tax Due": lines[i + 2], "Owner": lines[i + 3],
                        "Paid By": lines[i + 4],
                    })
                    i += 5
                else:
                    i += 1
            data["Payment History JSON"] = json.dumps(payments)

        oh = re.search(r"Owner History.*", body_text, re.DOTALL)
        if oh:
            owner_lines = [l.strip() for l in oh.group(0).split("\n") if l.strip()][5:]
            owners, i = [], 0
            while i < len(owner_lines) - 4:
                if owner_lines[i].isdigit() and len(owner_lines[i]) == 4:
                    owners.append({
                        "Tax Year": owner_lines[i], "Date": owner_lines[i + 1],
                        "Grantee": owner_lines[i + 2], "Instrument": owner_lines[i + 3],
                        "Bk/Pg": owner_lines[i + 4],
                    })
                    i += 5
                else:
                    i += 1
            data["Owner History JSON"] = json.dumps(owners)

    # ------------------------------------------------------------------ #
    # DataFrame / export (dynamic improvement columns, stable ordering)
    # ------------------------------------------------------------------ #
    def build_dataframe(self):
        rows = list(self.scraped_data_list)
        if not rows:
            return pd.DataFrame()
        present, seen = [], set()
        for r in rows:
            for k in r:
                if k not in seen:
                    seen.add(k)
                    present.append(k)
        first_seen = {k: i for i, k in enumerate(present)}

        def imp_key(k):
            n, fld = IMP_COL_RE.match(k).groups()
            fixed = IMP_FIELDS.index(fld) if fld in IMP_FIELDS else len(IMP_FIELDS)
            em = re.match(r"^Extra (\d+)_(Code|Description)$", fld)
            if em:
                return (int(n), fixed, 1, int(em.group(1)), em.group(2) != "Code", 0)
            return (int(n), fixed, 0, 0, 0, first_seen[k])

        def land_key(k):
            n, fld = LAND_COL_RE.match(k).groups()
            return (int(n), LAND_FIELDS.index(fld) if fld in LAND_FIELDS else 99)

        imp_cols = sorted((k for k in present if IMP_COL_RE.match(k)), key=imp_key)
        land_extra = sorted((k for k in present if LAND_COL_RE.match(k)), key=land_key)

        ordered = [c for c in HEAD_COLUMNS if c in seen] + imp_cols
        ordered += [c for c in LAND_COLUMNS if c in seen] + land_extra
        ordered += [c for c in TAX_COLUMNS + JSON_COLUMNS if c in seen]
        ordered += [c for c in present if c not in set(ordered)]
        return pd.DataFrame(rows).reindex(columns=ordered)

    def save_dataframe(self, df, path):
        """.xlsx keeps Section/Township/Range, dates, parcel numbers etc. as text
        (Excel auto-converts '8 / 14 / 1' to a date when opening a CSV)."""
        ext = os.path.splitext(path)[1].lower()
        if ext in (".xlsx", ".xlsm"):
            if len(df) > 1_048_575:
                path = os.path.splitext(path)[0] + ".csv"
                self.log("⚠️ Too many rows for one Excel sheet - saving as CSV instead.")
            else:
                try:
                    clean = df.copy()
                    for col in clean.columns:
                        clean[col] = clean[col].map(
                            lambda v: ILLEGAL_XLSX_RE.sub("", v) if isinstance(v, str) else v
                        )
                    clean.to_excel(path, index=False, engine="openpyxl")
                    return path
                except ImportError:
                    path = os.path.splitext(path)[0] + ".csv"
                    self.log("⚠️ openpyxl not installed (pip install openpyxl) - saving as CSV instead.")
        df.to_csv(path, index=False, encoding="utf-8-sig")
        return path

    # ------------------------------------------------------------------ #
    # Checkpoint / session helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def checkpoint_path(input_file):
        return os.path.splitext(input_file)[0] + "_checkpoint.jsonl"

    def count_checkpoint(self, input_file):
        path = self.checkpoint_path(input_file)
        if not input_file or not os.path.exists(path):
            return 0
        urls = set()
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                try:
                    rec = json.loads(line)
                except ValueError:
                    continue
                if rec.get("Status") == "Success":
                    urls.add(rec.get("URL"))
        return len(urls)

    def save_session(self):
        try:
            with open(SESSION_FILE, "w", encoding="utf-8") as fh:
                json.dump(
                    {
                        "input": self.input_file_path.get(),
                        "output": self.output_file_path.get(),
                        "workers": self.max_workers_var.get(),
                    },
                    fh,
                )
        except OSError:
            pass

    def restore_session(self):
        """After a crash/shutdown/close, remember the last input file."""
        try:
            with open(SESSION_FILE, encoding="utf-8") as fh:
                sess = json.load(fh)
        except (OSError, ValueError):
            return
        if sess.get("input") and os.path.exists(sess["input"]):
            self.input_file_path.set(sess["input"])
        if sess.get("output"):
            self.output_file_path.set(sess["output"])
        if sess.get("workers"):
            self.max_workers_var.set(sess["workers"])
        self.refresh_resume_hint()

    def refresh_resume_hint(self):
        n = self.count_checkpoint(self.input_file_path.get())
        if n:
            self.status_label.config(
                text=f"Status: Saved progress found - {n} properties already scraped. Click ▶ Resume to continue."
            )

    def start_scraping_thread(self, resume=False):
        input_file = self.input_file_path.get()
        if not input_file or not os.path.exists(input_file):
            messagebox.showerror(
                "Error", "Please select a valid input CSV file first."
            )
            return

        if self.is_running:
            return

        existing = self.count_checkpoint(input_file)
        ckpt = self.checkpoint_path(input_file)
        if resume and existing == 0:
            messagebox.showinfo(
                "Nothing to resume",
                "No saved progress was found for this input file. Use Start instead.",
            )
            return
        if not resume and existing > 0:
            if not messagebox.askyesno(
                "Saved progress found",
                f"{existing} properties were already scraped for this input file.\n\n"
                "Yes = start over from the beginning (old progress is kept as a "
                "timestamped backup file).\nNo = cancel (then click Resume to continue).",
            ):
                return
            os.replace(
                ckpt,
                os.path.splitext(ckpt)[0] + time.strftime("_old_%Y%m%d_%H%M%S.jsonl"),
            )

        self.save_session()
        self.is_running = True
        self.stop_requested = False
        self.start_btn.config(state="disabled")
        self.resume_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.export_btn.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.scraped_data_list = []

        threading.Thread(
            target=self.run_scraper, args=(input_file,), daemon=True
        ).start()

    def stop_scraping(self):
        if self.is_running:
            self.stop_requested = True
            self.log(
                "🛑 Stop requested! Waiting for active threads to wrap up..."
            )
            self.status_label.config(
                text="Status: Stopping... Please wait."
            )

    def manual_export(self):
        if not self.scraped_data_list:
            messagebox.showwarning(
                "Warning", "No data available to export yet."
            )
            return
        output_file = self.save_dataframe(
            self.build_dataframe(), self.output_file_path.get()
        )
        self.log(
            f"📥 Manually exported {len(self.scraped_data_list)} records to '{output_file}'."
        )
        messagebox.showinfo("Exported", f"Successfully saved to {output_file}")

    def update_table_view(self):
        self.tree.delete(*self.tree.get_children())
        if not self.scraped_data_list:
            return

        df = self.build_dataframe().fillna("")
        columns = list(df.columns)

        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="w")

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def run_scraper(self, input_file):
        try:
            df_input = pd.read_csv(input_file)
            url_col = next(
                (
                    col
                    for col in df_input.columns
                    if col.lower() in ["url", "link", "property_url"]
                ),
                None,
            )

            if not url_col:
                self.log(
                    "❌ Error: CSV file must contain a column named 'url' or 'link'."
                )
                self.reset_ui()
                return

            urls = df_input[url_col].dropna().tolist()
            total_urls = len(urls)

            # ---- crash-safe checkpoint: every result is written to disk
            # the moment it finishes, so a shutdown loses at most the pages
            # that were in flight. Re-running with the same input resumes.
            checkpoint = self.checkpoint_path(input_file)
            wanted = set(urls)
            done = {}
            if os.path.exists(checkpoint):
                with open(checkpoint, encoding="utf-8") as fh:
                    for line in fh:
                        try:
                            rec = json.loads(line)
                        except ValueError:
                            continue  # half-written last line after a power cut
                        if rec.get("Status") == "Success" and rec.get("URL") in wanted:
                            done[rec["URL"]] = rec
            self.scraped_data_list = list(done.values())
            urls_todo = [u for u in urls if u not in done]
            if done:
                self.log(f"♻️ Resuming: {len(done)} already scraped (from {checkpoint}), {len(urls_todo)} left.")
                self.root.after(0, self.update_table_view)
            ckpt_file = open(checkpoint, "a", encoding="utf-8")
            max_workers = self.max_workers_var.get()

            self.log(
                f"🚀 Loaded {total_urls} links. Starting scraper with {max_workers} threads..."
            )

            self.progress_bar["maximum"] = total_urls
            self.progress_bar["value"] = 0

            completed = len(done)
            self.progress_bar["value"] = completed

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                self.executor = executor
                future_to_url = {
                    executor.submit(self.scrape_property_page, url): url
                    for url in urls_todo
                }

                for future in as_completed(future_to_url):
                    if self.stop_requested:
                        self.log(
                            "⚠️ Scraping halted early by user request."
                        )
                        break

                    completed += 1
                    result = future.result()
                    self.scraped_data_list.append(result)
                    ckpt_file.write(json.dumps(result, default=str) + "\n")
                    ckpt_file.flush()
                    os.fsync(ckpt_file.fileno())

                    self.root.after(0, self.update_table_view)

                    self.progress_bar["value"] = completed
                    self.status_label.config(
                        text=f"Progress: {completed}/{total_urls} processed"
                    )

                    status_icon = (
                        "✅" if result["Status"] == "Success" else "⚠️"
                    )
                    self.log(
                        f"[{completed}/{total_urls}] {status_icon} Scraped: {result['URL']} | Owner: {result['Owner']}"
                    )

            ckpt_file.close()

            output_file = self.output_file_path.get()
            if self.scraped_data_list:
                output_file = self.save_dataframe(
                    self.build_dataframe(), output_file
                )
                self.log(
                    f"\n🎉 Scraping finished! Successfully saved to '{output_file}'."
                )

        except Exception as e:
            self.log(f"❌ Critical Error: {str(e)}")
        finally:
            self.reset_ui()

    def reset_ui(self):
        self.is_running = False
        self.stop_requested = False
        self.start_btn.config(state="normal")
        self.resume_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.export_btn.config(state="normal")
        self.status_label.config(text="Status: Idle / Ready")


if __name__ == "__main__":
    root = tk.Tk()
    app = BlountPropertyScraperApp(root)
    root.mainloop()