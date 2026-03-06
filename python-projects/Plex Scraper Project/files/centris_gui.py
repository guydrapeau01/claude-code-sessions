#!/usr/bin/env python3
"""
Centris Scraper — Desktop App
==============================
Double-click to run. No command line needed.
Requires: pip install playwright openpyxl
          playwright install chromium
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import subprocess
import sys
import os
import queue
import re
from datetime import datetime

# ── Attempt to import scraper logic ──────────────────────────────────────────
SCRAPER_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRAPER_DIR)

try:
    from centris_app import scrape_all, analyze, generate_excel, CITY_SLUGS
    SCRAPER_AVAILABLE = True
except ImportError:
    SCRAPER_AVAILABLE = False

# ── Colour palette ────────────────────────────────────────────────────────────
BG          = "#1a1a2e"
PANEL       = "#16213e"
ACCENT      = "#0f3460"
GREEN       = "#00d4aa"
RED         = "#e94560"
YELLOW      = "#f5a623"
TEXT        = "#e8e8e8"
TEXT_DIM    = "#8888aa"
ENTRY_BG    = "#0d1b2a"
ENTRY_FG    = "#ffffff"
BTN_BG      = "#00d4aa"
BTN_FG      = "#0a0a1a"
BTN_HOV     = "#00ffcc"

FONT_TITLE  = ("Segoe UI", 22, "bold")
FONT_HEAD   = ("Segoe UI", 11, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_MONO   = ("Consolas", 9)

# ── Main App ──────────────────────────────────────────────────────────────────
class CentrisApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Centris Scraper")
        self.geometry("780x680")
        self.minsize(680, 580)
        self.configure(bg=BG)
        self.resizable(True, True)

        # Set window icon (Windows .ico / Mac/Linux fallback)
        try:
            self.iconbitmap(default='')
        except Exception:
            pass

        self._log_queue = queue.Queue()
        self._running   = False
        self._result    = None  # path to output Excel

        self._build_ui()
        self._check_deps()
        self._poll_log()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ──
        hdr = tk.Frame(self, bg=ACCENT, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🏢  Centris Multifamilial Scraper",
                 font=FONT_TITLE, bg=ACCENT, fg=GREEN).pack()
        tk.Label(hdr, text="Automated TGA analysis for income properties",
                 font=FONT_SMALL, bg=ACCENT, fg=TEXT_DIM).pack()

        # ── Body split: left=filters, right=log ──
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=12)
        body.columnconfigure(0, weight=0, minsize=290)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # ─── LEFT: Filters panel ───
        left = tk.Frame(body, bg=PANEL, relief="flat", bd=0)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        left.pack_propagate(False)
        left.configure(width=290)

        tk.Label(left, text="SEARCH FILTERS", font=FONT_HEAD,
                 bg=PANEL, fg=GREEN).pack(anchor="w", padx=16, pady=(14,6))
        ttk.Separator(left, orient="horizontal").pack(fill="x", padx=12, pady=2)

        # Cities
        self._section(left, "Cities")
        self.city_vars = {}
        cities = [('montreal','Montréal'), ('laval','Laval'), ('longueuil','Longueuil')]
        cf = tk.Frame(left, bg=PANEL)
        cf.pack(fill="x", padx=16, pady=(2,8))
        for slug, label in cities:
            v = tk.BooleanVar(value=(slug == 'montreal'))
            self.city_vars[slug] = v
            cb = tk.Checkbutton(cf, text=label, variable=v,
                                bg=PANEL, fg=TEXT, selectcolor=ACCENT,
                                activebackground=PANEL, activeforeground=GREEN,
                                font=FONT_BODY, anchor="w", cursor="hand2")
            cb.pack(anchor="w")

        # Max Price
        self._section(left, "Max Price ($)")
        self.price_var = tk.StringVar(value="2,000,000")
        self._entry(left, self.price_var)

        # Min Units
        self._section(left, "Min Number of Units")
        self.units_var = tk.StringVar(value="5")
        self._entry(left, self.units_var)

        # Days
        self._section(left, "Listed in last N days  (0 = any)")
        self.days_var = tk.StringVar(value="7")
        self._entry(left, self.days_var)

        # Max props
        self._section(left, "Max properties per city")
        self.max_var = tk.StringVar(value="50")
        self._entry(left, self.max_var)

        # Output folder
        self._section(left, "Save Excel to folder")
        frow = tk.Frame(left, bg=PANEL)
        frow.pack(fill="x", padx=16, pady=(2,10))
        self.folder_var = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        fe = tk.Entry(frow, textvariable=self.folder_var, bg=ENTRY_BG, fg=ENTRY_FG,
                      font=FONT_SMALL, relief="flat", bd=4, insertbackground=ENTRY_FG)
        fe.pack(side="left", fill="x", expand=True)
        tk.Button(frow, text="…", bg=ACCENT, fg=TEXT, relief="flat",
                  font=FONT_BODY, cursor="hand2", padx=6,
                  command=self._browse_folder).pack(side="left", padx=(4,0))

        # Spacer + Run button
        tk.Frame(left, bg=PANEL, height=8).pack()
        self.run_btn = tk.Button(
            left, text="▶  Run Scraper",
            bg=BTN_BG, fg=BTN_FG, font=("Segoe UI", 12, "bold"),
            relief="flat", cursor="hand2", pady=10,
            command=self._on_run
        )
        self.run_btn.pack(fill="x", padx=16, pady=(4,16))
        self.run_btn.bind("<Enter>", lambda e: self.run_btn.config(bg=BTN_HOV))
        self.run_btn.bind("<Leave>", lambda e: self.run_btn.config(bg=BTN_BG))

        # ─── RIGHT: Log panel ───
        right = tk.Frame(body, bg=PANEL)
        right.grid(row=0, column=1, sticky="nsew")

        loghead = tk.Frame(right, bg=PANEL)
        loghead.pack(fill="x", padx=12, pady=(12,4))
        tk.Label(loghead, text="LIVE LOG", font=FONT_HEAD,
                 bg=PANEL, fg=GREEN).pack(side="left")
        tk.Button(loghead, text="Clear", bg=ACCENT, fg=TEXT_DIM,
                  font=FONT_SMALL, relief="flat", cursor="hand2", padx=6,
                  command=self._clear_log).pack(side="right")

        ttk.Separator(right, orient="horizontal").pack(fill="x", padx=8, pady=2)

        # Log text area
        lf = tk.Frame(right, bg=PANEL)
        lf.pack(fill="both", expand=True, padx=8, pady=(4,4))
        self.log_text = tk.Text(lf, bg="#0a0e1a", fg=TEXT, font=FONT_MONO,
                                relief="flat", state="disabled",
                                wrap="word", bd=0,
                                insertbackground=TEXT, selectbackground=ACCENT)
        sb = ttk.Scrollbar(lf, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)

        # Tag colours for log
        self.log_text.tag_config("green",  foreground=GREEN)
        self.log_text.tag_config("yellow", foreground=YELLOW)
        self.log_text.tag_config("red",    foreground=RED)
        self.log_text.tag_config("dim",    foreground=TEXT_DIM)
        self.log_text.tag_config("bold",   font=("Consolas", 9, "bold"))

        # ── Status bar ──
        status_bar = tk.Frame(self, bg=ACCENT, height=30)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)
        self.status_var = tk.StringVar(value="Ready — configure filters and click Run")
        tk.Label(status_bar, textvariable=self.status_var,
                 bg=ACCENT, fg=TEXT_DIM, font=FONT_SMALL,
                 anchor="w", padx=14).pack(side="left", fill="y")
        self.open_btn = tk.Button(
            status_bar, text="📂 Open Excel", bg=GREEN, fg=BTN_FG,
            font=FONT_SMALL, relief="flat", cursor="hand2", padx=10,
            command=self._open_result, state="disabled"
        )
        self.open_btn.pack(side="right", padx=8, pady=3)

    def _section(self, parent, text):
        tk.Label(parent, text=text, font=FONT_SMALL, bg=PANEL,
                 fg=TEXT_DIM).pack(anchor="w", padx=16, pady=(8,2))

    def _entry(self, parent, var):
        e = tk.Entry(parent, textvariable=var, bg=ENTRY_BG, fg=ENTRY_FG,
                     font=FONT_BODY, relief="flat", bd=6,
                     insertbackground=ENTRY_FG)
        e.pack(fill="x", padx=16, pady=(0,2))

    def _browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get())
        if folder:
            self.folder_var.set(folder)

    # ── Dependency check ──────────────────────────────────────────────────────
    def _check_deps(self):
        missing = []
        try: import playwright
        except ImportError: missing.append("playwright")
        try: import openpyxl
        except ImportError: missing.append("openpyxl")

        if missing or not SCRAPER_AVAILABLE:
            msg = []
            if not SCRAPER_AVAILABLE:
                msg.append("⚠  centris_app.py not found in the same folder as this file.")
            if missing:
                msg.append(f"⚠  Missing packages: {', '.join(missing)}")
                msg.append(f"   Run: pip install {' '.join(missing)}")
                msg.append("   Then: playwright install chromium")
            self._log("\n".join(msg), "red")
            self.run_btn.config(state="disabled")
        else:
            self._log("✓ All dependencies found.", "green")
            self._log("Configure filters and click ▶ Run Scraper\n", "dim")

    # ── Log helpers ───────────────────────────────────────────────────────────
    def _log(self, text, tag=None):
        """Thread-safe log write."""
        self._log_queue.put((text, tag))

    def _flush_log(self, text, tag):
        self.log_text.configure(state="normal")
        # Colour-code common patterns automatically
        if tag is None:
            if any(x in text for x in ["✓","✔","🟢","🏆"]):  tag = "green"
            elif any(x in text for x in ["⚠","🔴","✗","ERROR"]): tag = "red"
            elif any(x in text for x in ["🟡","…","Loading"]): tag = "yellow"
            elif text.startswith("═") or text.startswith("─"): tag = "dim"
        self.log_text.insert("end", text + "\n", tag or "")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _poll_log(self):
        while True:
            try:
                text, tag = self._log_queue.get_nowait()
                self._flush_log(text, tag)
            except queue.Empty:
                break
        self.after(100, self._poll_log)

    # ── Run ───────────────────────────────────────────────────────────────────
    def _parse_int(self, var, default):
        try:
            return int(re.sub(r'[^\d]', '', var.get()) or default)
        except Exception:
            return default

    def _on_run(self):
        if self._running:
            return

        # Collect filters
        cities = [(label, slug) for slug, (label, _s) in
                  [('montreal', ('Montréal','montreal')),
                   ('laval',    ('Laval','laval')),
                   ('longueuil',('Longueuil','longueuil'))]
                  if self.city_vars[slug].get()]

        # Map to the tuple format scraper expects: (label, slug)
        city_tuples = []
        slug_map = {'montreal': ('Montréal','montreal'),
                    'laval':    ('Laval','laval'),
                    'longueuil':('Longueuil','longueuil')}
        for slug, v in self.city_vars.items():
            if v.get():
                city_tuples.append(slug_map[slug])

        if not city_tuples:
            messagebox.showwarning("No city selected", "Please select at least one city.")
            return

        filters = {
            'cities'    : city_tuples,
            'max_price' : self._parse_int(self.price_var, 2_000_000),
            'min_units' : self._parse_int(self.units_var, 5),
            'days'      : self._parse_int(self.days_var, 7),
            'max_props' : self._parse_int(self.max_var, 50),
            'output_dir': self.folder_var.get(),
        }

        self._running = True
        self._result  = None
        self.run_btn.config(state="disabled", text="⏳  Running…", bg="#555")
        self.open_btn.config(state="disabled")
        self.status_var.set("Scraping in progress…")
        self._clear_log()

        threading.Thread(target=self._run_scraper, args=(filters,),
                         daemon=True).start()

    def _run_scraper(self, filters):
        import io, contextlib

        class LogCapture(io.TextIOWrapper):
            """Redirect stdout prints to our log queue."""
            def __init__(self, log_fn):
                self._log_fn = log_fn
                self._buf = ""
            def write(self, s):
                self._buf += s
                while "\n" in self._buf:
                    line, self._buf = self._buf.split("\n", 1)
                    self._log_fn(line)
            def flush(self): pass

        cap = LogCapture(self._log)
        old_stdout = sys.stdout
        sys.stdout = cap

        try:
            self._log("═"*50, "dim")
            self._log(f"  Starting scrape — {datetime.now().strftime('%H:%M:%S')}", "bold")
            self._log("═"*50, "dim")

            properties = scrape_all(filters)

            self._log(f"\n  Scraped {len(properties)} properties", "green")

            if not properties:
                self._log("\n⚠ No properties found. Try adjusting filters.", "yellow")
                return

            # Analysis
            self._log("\n" + "═"*50, "dim")
            self._log("  RUNNING TGA ANALYSIS…", "bold")
            self._log("═"*50, "dim")

            analyses, no_income = [], []
            for prop in properties:
                if prop.get('gross_income'):
                    result = analyze(prop)
                    if result:
                        analyses.append(result)
                        r    = result['best_ratio']
                        icon = '🟢' if r >= 1.15 else '🟡' if r >= 1.0 else '🔴'
                        self._log(f"  {icon} {prop.get('address','?')[:46]}")
                        self._log(f"     Ratio: {r:.1%}  CF: ${result['best_cf']:,.0f}/mo  ROI: {result['best_roi']:.1%}", "dim")
                    else:
                        no_income.append(prop)
                else:
                    no_income.append(prop)
                    self._log(f"  ⚪ {prop.get('address','?')[:46]} — no income data", "dim")

            for p in no_income:
                analyses.append({**p, 'best_ratio':0, 'noi':None, 'total_exp':None,
                                  'exp_ratio':None, 'best_scenario':None,
                                  'best_econ':None, 'best_down':None,
                                  'best_pmt':None, 'best_cf':None, 'best_roi':None,
                                  'all_sc':{}, 'expenses':{}})

            # Excel
            self._log("\n" + "═"*50, "dim")
            self._log("  GENERATING EXCEL…", "bold")

            # Override output dir
            original_dir = os.getcwd()
            out_dir = filters.get('output_dir', os.path.expanduser("~/Desktop"))
            os.chdir(out_dir)
            try:
                fname = generate_excel(analyses, filters)
                full_path = os.path.join(out_dir, fname)
            finally:
                os.chdir(original_dir)

            self._result = full_path
            self._log(f"\n  ✓ Excel saved:", "green")
            self._log(f"  {full_path}", "green")

            scored = [a for a in analyses if a.get('noi')]
            if scored:
                best = max(scored, key=lambda x: x['best_ratio'])
                self._log("\n  🏆 BEST DEAL:", "bold")
                self._log(f"  {best.get('address','?')}", "green")
                self._log(f"  Price: ${best['price']:,.0f}  |  Ratio: {best['best_ratio']:.1%}  |  CF: ${best['best_cf']:,.0f}/mo", "green")

            self._log("\n" + "═"*50, "dim")
            self._log("  DONE ✓", "green")
            self._log("═"*50, "dim")

        except Exception as ex:
            self._log(f"\n❌ Error: {ex}", "red")
            import traceback
            self._log(traceback.format_exc(), "red")
        finally:
            sys.stdout = old_stdout
            self._running = False
            self.after(0, self._on_done)

    def _on_done(self):
        self.run_btn.config(state="normal", text="▶  Run Scraper", bg=BTN_BG)
        if self._result:
            self.status_var.set(f"Done ✓  →  {os.path.basename(self._result)}")
            self.open_btn.config(state="normal")
        else:
            self.status_var.set("Finished — check log for details")

    def _open_result(self):
        if self._result and os.path.exists(self._result):
            if sys.platform == "win32":
                os.startfile(self._result)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self._result])
            else:
                subprocess.Popen(["xdg-open", self._result])

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = CentrisApp()
    app.mainloop()
