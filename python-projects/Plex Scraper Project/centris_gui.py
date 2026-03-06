#!/usr/bin/env python3
"""
CENTRIS INTELLIGENCE — Premium Multifamilial Analysis Terminal
Dark luxury terminal aesthetic. Bloomberg meets real estate.
"""

import tkinter as tk
from tkinter import filedialog
import threading, subprocess, sys, os, queue, re, math, time
from datetime import datetime

SCRAPER_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRAPER_DIR)
try:
    from centris_app import scrape_all, analyze, generate_excel
    SCRAPER_AVAILABLE = True
except ImportError:
    SCRAPER_AVAILABLE = False

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  DESIGN TOKENS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BG       = "#080A0F"
SURFACE  = "#0D1018"
PANEL    = "#111520"
CARD     = "#161B26"
BORDER   = "#1E2535"
BORDER_B = "#2A3348"

AMBER    = "#D4A843"
AMBER_L  = "#F0C865"
AMBER_D  = "#8A6B22"
AMBER_DIM= "#3D2E0A"

CYAN     = "#1BE6CC"
CYAN_D   = "#0B4A42"

GREEN    = "#3DD68C"
RED      = "#E05555"
YELLOW   = "#F0BE55"
BLUE     = "#4A9EFF"

TEXT     = "#D8DCE8"
TEXT2    = "#7A8499"
TEXT3    = "#3A4155"
WHITE    = "#F0F2F8"

# Fonts
MONO    = "Courier New"
SERIF   = "Georgia"
SANS    = "Segoe UI"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  CUSTOM WIDGETS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class DataField(tk.Frame):
    """Labeled input with amber focus ring animation."""
    def __init__(self, parent, label, value='', **kw):
        super().__init__(parent, bg=PANEL, **kw)
        self.var = tk.StringVar(value=value)
        self._focused = False

        lbl_row = tk.Frame(self, bg=PANEL)
        lbl_row.pack(fill='x', padx=14, pady=(10,3))
        tk.Label(lbl_row, text='▸ ', font=(MONO,12,'bold'),
                 bg=PANEL, fg=AMBER_D).pack(side='left')
        tk.Label(lbl_row, text=label, font=(MONO,12,'bold'),
                 bg=PANEL, fg=TEXT2).pack(side='left')

        self._ring = tk.Frame(self, bg=BORDER_B)
        self._ring.pack(fill='x', padx=14, pady=(0,10))

        inner = tk.Frame(self._ring, bg=CARD)
        inner.pack(fill='x', padx=1, pady=1)

        self.entry = tk.Entry(inner, textvariable=self.var,
                              bg=CARD, fg=WHITE,
                              font=(SANS, 12), relief='flat',
                              bd=10, insertbackground=AMBER,
                              selectbackground=AMBER_D,
                              selectforeground=WHITE)
        self.entry.pack(fill='x')
        self.entry.bind('<FocusIn>',  self._on_focus)
        self.entry.bind('<FocusOut>', self._on_blur)

    def _on_focus(self, e):
        self._ring.config(bg=AMBER)
    def _on_blur(self, e):
        self._ring.config(bg=BORDER_B)
    def get(self):
        return self.var.get()


class MarketToggle(tk.Label):
    """Market toggle chip - reliable tk.Label based."""
    def __init__(self, parent, label, active=False, **kw):
        self.var = tk.BooleanVar(value=active)
        super().__init__(parent, text=label,
                         font=(MONO, 8, 'bold'),
                         padx=14, pady=9,
                         cursor='hand2', **kw)
        self._refresh()
        self.bind('<Button-1>', self._click)
        self.bind('<Enter>',    self._hover_on)
        self.bind('<Leave>',    self._hover_off)

    def _refresh(self):
        if self.var.get():
            self.config(bg=AMBER,   fg=BG,    relief='flat')
        else:
            self.config(bg=CARD,    fg=TEXT3, relief='flat')

    def _click(self, e=None):
        self.var.set(not self.var.get())
        self._refresh()

    def _hover_on(self, e=None):
        if not self.var.get(): self.config(bg=BORDER_B, fg=AMBER)

    def _hover_off(self, e=None):
        self._refresh()

    def get(self):
        return self.var.get()


class MetricTile(tk.Frame):
    """Big number metric tile with label."""
    def __init__(self, parent, label, color=TEXT2, **kw):
        super().__init__(parent, bg=CARD, **kw)
        self._color = color
        self.configure(padx=20, pady=16)
        tk.Label(self, text=label, font=(MONO,13,'bold'),
                 bg=CARD, fg=TEXT3).pack(anchor='w')
        self.num = tk.Label(self, text='—',
                            font=(SERIF,24,'bold'),
                            bg=CARD, fg=color)
        self.num.pack(anchor='w', pady=(2,0))

    def set(self, val, color=None):
        self.num.config(text=str(val),
                        fg=color or self._color)


class RunButton(tk.Frame):
    """Reliable styled run button using tk.Button underneath."""
    def __init__(self, parent, command, **kw):
        super().__init__(parent, bg=PANEL, **kw)
        self._cmd     = command
        self._enabled = True
        self._label   = '▶   EXECUTE SCAN'

        self._btn = tk.Button(
            self, text=self._label,
            bg=AMBER, fg=BG,
            font=(MONO, 12, 'bold'),
            relief='flat', bd=0,
            cursor='hand2',
            padx=0, pady=18,
            activebackground=AMBER_L,
            activeforeground=BG,
            command=self._click
        )
        self._btn.pack(fill='both', expand=True)

    def _click(self):
        if self._enabled and self._cmd:
            self._cmd()

    def set_state(self, label, enabled=True):
        self._label   = label
        self._enabled = enabled
        if enabled:
            self._btn.config(
                text=label, state='normal',
                bg=AMBER, fg=BG,
                activebackground=AMBER_L
            )
        else:
            self._btn.config(
                text=label, state='disabled',
                bg=CARD, fg=TEXT3
            )

    def tick(self):
        pass  # no animation needed - tk.Button handles hover natively


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  MAIN APPLICATION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('CENTRIS INTELLIGENCE')
        self.configure(bg=BG)
        self.resizable(True, True)
        try:    self.state('zoomed')
        except: self.geometry('1200x820')

        self._q        = queue.Queue()
        self._running  = False
        self._result   = None
        self._frame    = 0   # animation frame counter

        self._build_ui()
        self._check_deps()
        self._loop()

    # ─── UI CONSTRUCTION ─────────────────────────────────────────────────────

    def _build_ui(self):
        self._topbar()
        tk.Frame(self, bg=AMBER, height=1).pack(fill='x')    # amber hairline
        self._body()
        self._statusbar()

    def _topbar(self):
        bar = tk.Frame(self, bg=SURFACE, height=62)
        bar.pack(fill='x')
        bar.pack_propagate(False)

        # Left: wordmark
        lf = tk.Frame(bar, bg=SURFACE)
        lf.pack(side='left', padx=20, fill='y')
        tk.Label(lf, text='◈', font=(SERIF,16),
                 bg=SURFACE, fg=AMBER).pack(side='left', pady=12)
        tk.Label(lf, text='  CENTRIS', font=(SERIF,15,'bold'),
                 bg=SURFACE, fg=WHITE).pack(side='left')
        tk.Label(lf, text=' INTELLIGENCE', font=(MONO,11),
                 bg=SURFACE, fg=AMBER_D).pack(side='left', pady=(18,0))

        # Right: system info
        rf = tk.Frame(bar, bg=SURFACE)
        rf.pack(side='right', padx=20, fill='y')
        self._clock_lbl = tk.Label(rf, text='', font=(MONO,12),
                                   bg=SURFACE, fg=TEXT3)
        self._clock_lbl.pack(side='right', pady=18)
        tk.Label(rf, text='MULTIFAMILIAL ENGINE  •  ', font=(MONO,12),
                 bg=SURFACE, fg=TEXT3).pack(side='right', pady=18)

    def _body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill='both', expand=True)
        body.columnconfigure(0, weight=0, minsize=295)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        self._sidebar(body)
        self._main_panel(body)

    def _sidebar(self, parent):
        sb = tk.Frame(parent, bg=SURFACE, width=295)
        sb.grid(row=0, column=0, sticky='nsew')
        sb.pack_propagate(False)

        # Section: Parameters
        self._divider(sb, 'PARAMETERS')

        # Markets
        tk.Label(sb, text='MARKETS', font=(MONO,13,'bold'),
                 bg=SURFACE, fg=AMBER_D).pack(anchor='w', padx=18, pady=(14,6))

        mrow = tk.Frame(sb, bg=SURFACE)
        mrow.pack(fill='x', padx=18, pady=(0,6))
        self._cities = {}
        for slug, label, on in [('montreal','MONTRÉAL',True),
                                  ('laval','LAVAL',False),
                                  ('longueuil','LONGUEUIL',False)]:
            w = MarketToggle(mrow, label, active=on)
            w.pack(side='left', padx=(0,5))
            self._cities[slug] = w

        tk.Frame(sb, bg=BORDER, height=1).pack(fill='x', padx=18, pady=6)

        # Filter fields
        for attr, lbl, val in [
            ('e_price', 'MAX PRICE  ($)',            '2,000,000'),
            ('e_units', 'MIN UNITS',                 '5'),
            ('e_days',  'LISTED WITHIN  (days)',     '7'),
            ('e_max',   'MAX PROPERTIES  (per city)','50'),
        ]:
            w = DataField(sb, lbl, val)
            w.pack(fill='x', padx=4)
            setattr(self, attr, w)

        # Output folder
        self._divider(sb, 'OUTPUT')
        of = tk.Frame(sb, bg=SURFACE)
        of.pack(fill='x', padx=18, pady=(8,6))
        self.folder_var = tk.StringVar(value=os.path.expanduser('~/Desktop'))
        ring = tk.Frame(of, bg=BORDER_B, pady=1)
        ring.pack(side='left', fill='x', expand=True)
        inn  = tk.Frame(ring, bg=CARD)
        inn.pack(fill='x', padx=1, pady=1)
        fe = tk.Entry(inn, textvariable=self.folder_var, bg=CARD, fg=TEXT2,
                      font=(SANS,12), relief='flat', bd=8, insertbackground=AMBER)
        fe.pack(fill='x')
        fe.bind('<FocusIn>',  lambda e: ring.config(bg=AMBER))
        fe.bind('<FocusOut>', lambda e: ring.config(bg=BORDER_B))
        tk.Button(of, text='⋯', bg=CARD, fg=TEXT2, font=(SANS,13),
                  relief='flat', cursor='hand2', padx=10, pady=2,
                  command=self._browse,
                  activebackground=BORDER_B,
                  activeforeground=AMBER).pack(side='left', padx=(5,0))

        # Spacer
        tk.Frame(sb, bg=SURFACE).pack(fill='both', expand=True)

        # Run button
        self._run_btn = RunButton(sb, command=self._on_run)
        self._run_btn.pack(fill='x', padx=18, pady=20)

    def _main_panel(self, parent):
        mp = tk.Frame(parent, bg=BG)
        mp.grid(row=0, column=1, sticky='nsew')
        mp.rowconfigure(1, weight=1)
        mp.columnconfigure(0, weight=1)

        # ── Metrics row ──
        metrics = tk.Frame(mp, bg=BG)
        metrics.grid(row=0, column=0, sticky='ew', padx=16, pady=(14,8))
        for i in range(4): metrics.columnconfigure(i, weight=1)

        defs = [
            ('PROPERTIES FOUND', CYAN),
            ('ANALYZED',         AMBER),
            ('RATIO ≥ 1.15×',    GREEN),
            ('BEST RATIO',       AMBER_L),
        ]
        self._tiles = []
        for i, (lbl, col) in enumerate(defs):
            outer = tk.Frame(metrics, bg=BORDER)
            outer.grid(row=0, column=i, sticky='ew',
                       padx=(0, 10 if i<3 else 0))
            tile = MetricTile(outer, lbl, color=col)
            tile.pack(fill='x', padx=1, pady=1)
            self._tiles.append(tile)

        # ── Terminal panel ──
        tp = tk.Frame(mp, bg=SURFACE)
        tp.grid(row=1, column=0, sticky='nsew', padx=16, pady=(0,16))
        tp.rowconfigure(1, weight=1)
        tp.columnconfigure(0, weight=1)

        # Terminal header bar
        th = tk.Frame(tp, bg=PANEL, height=36)
        th.grid(row=0, column=0, sticky='ew')
        th.pack_propagate(False)

        # Traffic-light dots (decorative)
        dots = tk.Frame(th, bg=PANEL)
        dots.pack(side='left', padx=14, pady=10)
        for col in [RED, YELLOW, GREEN]:
            tk.Label(dots, text='●', font=(MONO,12),
                     bg=PANEL, fg=col).pack(side='left', padx=2)

        tk.Label(th, text='LIVE ANALYSIS FEED',
                 font=(MONO,12,'bold'), bg=PANEL, fg=TEXT2).pack(side='left', padx=6)

        self._status_lbl = tk.Label(th, text='○  STANDBY',
                                    font=(MONO,12,'bold'), bg=PANEL, fg=TEXT3)
        self._status_lbl.pack(side='left', padx=12)

        # Scan line indicator
        self._scan_canvas = tk.Canvas(th, width=60, height=16,
                                      bg=PANEL, highlightthickness=0)
        self._scan_canvas.pack(side='left', padx=4, pady=10)

        # Right buttons
        btn_f = tk.Frame(th, bg=PANEL)
        btn_f.pack(side='right', padx=10, pady=6)

        self._open_btn = tk.Label(btn_f, text='↗ OPEN EXCEL',
                                  font=(MONO,13,'bold'), bg=CYAN_D, fg=CYAN,
                                  padx=10, pady=5, cursor='hand2')
        self._open_btn.pack(side='right', padx=(5,0))
        self._open_btn.bind('<Button-1>', lambda e: self._open_result())
        self._open_btn.bind('<Enter>', lambda e: self._open_btn.config(bg=CYAN, fg=BG))
        self._open_btn.bind('<Leave>', lambda e: self._open_btn.config(bg=CYAN_D, fg=CYAN))
        self._open_btn.config(state='disabled')

        _clr = tk.Label(btn_f, text='CLEAR', font=(MONO,13,'bold'),
                 bg=BORDER, fg=TEXT3, padx=8, pady=5, cursor='hand2')
        _clr.pack(side='right')
        _clr.bind('<Button-1>', lambda e: self._clear())

        # Amber separator
        tk.Frame(tp, bg=AMBER_D, height=1).grid(row=1, column=0, sticky='ew')

        # Progress bar
        self._pb_canvas = tk.Canvas(tp, height=2, bg=SURFACE,
                                    highlightthickness=0)
        self._pb_canvas.grid(row=2, column=0, sticky='ew')

        # Terminal text
        t_frame = tk.Frame(tp, bg=SURFACE)
        t_frame.grid(row=3, column=0, sticky='nsew')
        tp.rowconfigure(3, weight=1)

        self._term = tk.Text(t_frame, bg=SURFACE, fg=TEXT2,
                             font=(MONO, 11), relief='flat', bd=0,
                             state='disabled', wrap='word',
                             insertbackground=AMBER,
                             selectbackground=AMBER_DIM,
                             padx=18, pady=14,
                             spacing1=2, spacing3=2)

        vsb = tk.Scrollbar(t_frame, orient='vertical',
                           command=self._term.yview,
                           bg=SURFACE, troughcolor=PANEL,
                           activebackground=BORDER_B, width=10)
        self._term.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self._term.pack(side='left', fill='both', expand=True)

        # Tags
        self._term.tag_config('amber', foreground=AMBER,
                              font=(MONO,13,'bold'))
        self._term.tag_config('amber_l', foreground=AMBER_L,
                              font=(MONO,13,'bold'))
        self._term.tag_config('cyan',  foreground=CYAN)
        self._term.tag_config('green', foreground=GREEN)
        self._term.tag_config('red',   foreground=RED)
        self._term.tag_config('yellow',foreground=YELLOW)
        self._term.tag_config('dim',   foreground=TEXT3)
        self._term.tag_config('white', foreground=WHITE,
                              font=(MONO,13,'bold'))
        self._term.tag_config('url',   foreground=BLUE)

    def _statusbar(self):
        sb = tk.Frame(self, bg=PANEL, height=26)
        sb.pack(fill='x', side='bottom')
        sb.pack_propagate(False)
        tk.Frame(sb, bg=AMBER, width=3).pack(side='left', fill='y')
        self._sb_lbl = tk.Label(sb, text='Ready',
                                font=(MONO,11), bg=PANEL, fg=TEXT3,
                                anchor='w', padx=10)
        self._sb_lbl.pack(side='left', fill='y')

    # ─── HELPERS ─────────────────────────────────────────────────────────────

    def _divider(self, parent, label):
        f = tk.Frame(parent, bg=SURFACE)
        f.pack(fill='x', padx=18, pady=(16,0))
        tk.Frame(f, bg=AMBER_D, height=1).pack(fill='x')
        tk.Label(parent, text=label, font=(MONO,13,'bold'),
                 bg=SURFACE, fg=AMBER_D).pack(anchor='w', padx=18, pady=(3,0))

    def _browse(self):
        d = filedialog.askdirectory(initialdir=self.folder_var.get())
        if d: self.folder_var.set(d)

    def _log(self, text, tag=None):
        self._q.put((text, tag))

    def _flush(self, text, tag):
        if tag is None:
            if any(x in text for x in ['✓','🏆','SUCCESS']): tag = 'green'
            elif '🟢' in text:  tag = 'green'
            elif '🟡' in text:  tag = 'yellow'
            elif '🔴' in text:  tag = 'red'
            elif any(x in text for x in ['❌','ERROR']):  tag = 'red'
            elif any(x in text for x in ['⚠']):          tag = 'yellow'
            elif any(x in text for x in ['━','─','═']):  tag = 'dim'
            elif any(x in text for x in ['▶','◈','◆']):  tag = 'amber'

        self._term.configure(state='normal')
        # Prefix timestamp for main events
        self._term.insert('end', text + '\n', tag or '')
        self._term.see('end')
        self._term.configure(state='disabled')

    def _clear(self):
        if not hasattr(self, '_term'): return
        self._term.configure(state='normal')
        self._term.delete('1.0', 'end')
        self._term.configure(state='disabled')

    def _parse_int(self, w, default):
        try:    return int(re.sub(r'[^\d]','', w.get()) or default)
        except: return default

    # ─── ANIMATION LOOP ──────────────────────────────────────────────────────

    def _loop(self):
        # Drain log queue
        while True:
            try:
                t, g = self._q.get_nowait()
                self._flush(t, g)
            except queue.Empty:
                break

        self._frame += 1

        # Clock
        if self._frame % 10 == 0:
            self._clock_lbl.config(
                text=datetime.now().strftime('%Y-%m-%d  %H:%M:%S'))

        # Progress shimmer
        if self._running:
            self._pb_canvas.delete('all')
            w = self._pb_canvas.winfo_width() or 600
            pos = ((self._frame % 120) / 120) * (w + 200) - 100
            # gradient-like shimmer: 3 overlapping rects
            for offset, alpha_col in [(-20, AMBER_DIM), (0, AMBER), (20, AMBER_DIM)]:
                self._pb_canvas.create_rectangle(
                    pos+offset, 0, pos+offset+80, 2,
                    fill=alpha_col, outline='')

            # Scan line in header
            sc = self._scan_canvas
            sc.delete('all')
            sp = ((self._frame % 60) / 60) * 60
            sc.create_rectangle(0,6, sp,10, fill=AMBER_D, outline='')
            sc.create_rectangle(sp,6, sp+5,10, fill=AMBER, outline='')

            # Button pulse
            self._run_btn.tick()

        self.after(80, self._loop)

    # ─── DEPENDENCY CHECK ────────────────────────────────────────────────────

    def _check_deps(self):
        missing = []
        try: import playwright
        except ImportError: missing.append('playwright')
        try: import openpyxl
        except ImportError: missing.append('openpyxl')

        if missing or not SCRAPER_AVAILABLE:
            if not SCRAPER_AVAILABLE:
                self._log('❌  centris_app.py not found — same folder required', 'red')
            if missing:
                self._log(f'❌  Missing packages: pip install {" ".join(missing)}', 'red')
                self._log('    Then run: playwright install chromium', 'red')
            self._run_btn.set_state('✕   MISSING DEPENDENCIES', enabled=False)
        else:
            self._log('◈  CENTRIS INTELLIGENCE  —  SYSTEM READY', 'amber')
            self._log('─' * 55, 'dim')
            self._log('   Configure search parameters and execute scan.', 'dim')
            self._sb_lbl.config(text='System ready  •  All dependencies verified')

    # ─── RUN ─────────────────────────────────────────────────────────────────

    def _on_run(self):
        if self._running: return
        if not hasattr(self, '_term'): return

        cities = [(s, s) for s, w in self._cities.items() if w.get()]
        if not cities:
            self._log('⚠  Select at least one market to scan.', 'yellow')
            return

        filters = {
            'cities':     cities,
            'max_price':  self._parse_int(self.e_price, 2_000_000),
            'min_units':  self._parse_int(self.e_units, 5),
            'days':       self._parse_int(self.e_days,  7),
            'max_props':  self._parse_int(self.e_max,   50),
            'output_dir': self.folder_var.get(),
        }

        self._running = True
        self._result  = None
        self._clear()
        self._run_btn.set_state('⏳  SCANNING...', enabled=False)
        self._status_lbl.config(text='◉  ACTIVE', fg=CYAN)
        self._sb_lbl.config(text='Scan in progress...')
        self._open_btn.config(state='disabled', bg=CYAN_D, fg=CYAN)
        for t in self._tiles: t.set('—')

        threading.Thread(target=self._worker, args=(filters,), daemon=True).start()

    def _worker(self, filters):
        class Pipe:
            def __init__(self, fn): self._fn=fn; self._buf=''
            def write(self, s):
                self._buf += s
                while '\n' in self._buf:
                    line, self._buf = self._buf.split('\n', 1)
                    self._fn(line)
            def flush(self): pass

        old_out = sys.stdout
        sys.stdout = Pipe(self._log)

        try:
            self._log('▶  SCAN INITIATED', 'amber')
            self._log(f'   Markets  :  {" · ".join(c[0].upper() for c in filters["cities"])}', 'dim')
            self._log(f'   Price    :  ≤ ${filters["max_price"]:,}', 'dim')
            self._log(f'   Units    :  {filters["min_units"]}+', 'dim')
            w = f'{filters["days"]}d window' if filters['days'] else 'all time'
            self._log(f'   Listing  :  {w}', 'dim')
            self._log('━' * 55, 'dim')

            props = scrape_all(filters)
            n = len(props)
            self.after(0, lambda: self._tiles[0].set(str(n), CYAN))

            if not props:
                self._log('⚠  No properties matched the criteria.', 'yellow')
                return

            self._log(f'\n◈  {n} properties collected', 'cyan')
            self._log('━' * 55, 'dim')
            self._log('▶  TGA ANALYSIS', 'amber')

            analyses, no_inc, best_r = [], [], 0

            for prop in props:
                if prop.get('gross_income'):
                    res = analyze(prop)
                    if res:
                        analyses.append(res)
                        r = res['best_ratio']
                        if r > best_r: best_r = r
                        icon = '🟢' if r>=1.15 else '🟡' if r>=1.0 else '🔴'
                        self._log(f'  {icon}  {prop.get("address","?")[:44]}')
                        self._log(f'       {r:.1%}  ·  CF ${res["best_cf"]:,.0f}/mo  ·  ROI {res["best_roi"]:.1%}', 'dim')
                    else:
                        no_inc.append(prop)
                else:
                    no_inc.append(prop)
                    self._log(f'  ◌  {prop.get("address","?")[:44]}  — no income data', 'dim')

            for p in no_inc:
                analyses.append({**p, 'best_ratio':0,'noi':None,'total_exp':None,
                                  'exp_ratio':None,'best_scenario':None,'best_econ':None,
                                  'best_down':None,'best_pmt':None,'best_cf':None,
                                  'best_roi':None,'all_sc':{},'expenses':{}})

            green = sum(1 for a in analyses if a.get('best_ratio',0) >= 1.15)
            self.after(0, lambda: self._tiles[1].set(str(len(analyses)), AMBER))
            self.after(0, lambda: self._tiles[2].set(str(green), GREEN))
            if best_r:
                br = best_r
                self.after(0, lambda: self._tiles[3].set(f'{br:.1%}', AMBER_L))

            # Excel
            self._log('\n▶  GENERATING REPORT', 'amber')
            out  = filters.get('output_dir', os.path.expanduser('~/Desktop'))
            orig = os.getcwd()
            os.chdir(out)
            try:
                fname       = generate_excel(analyses, filters)
                self._result = os.path.join(out, fname)
            finally:
                os.chdir(orig)

            self._log(f'\n✓  Report saved', 'green')
            self._log(f'   {self._result}', 'cyan')

            scored = [a for a in analyses if a.get('noi')]
            if scored:
                best = max(scored, key=lambda x: x['best_ratio'])
                self._log('\n' + '━'*55, 'dim')
                self._log('🏆  TOP OPPORTUNITY', 'amber_l')
                self._log(f'    {best.get("address","?")}', 'white')
                self._log(f'    ${best["price"]:,.0f}  ·  ratio {best["best_ratio"]:.1%}  ·  CF ${best["best_cf"]:,.0f}/mo', 'green')
                self._log(f'    {best.get("url","")}', 'url')
            self._log('━'*55, 'dim')

        except Exception as ex:
            import traceback
            self._log(f'\n❌  {ex}', 'red')
            self._log(traceback.format_exc(), 'red')
        finally:
            sys.stdout = old_out
            self._running = False
            self._pb_canvas.delete('all')
            self.after(0, self._finish)

    def _finish(self):
        self._run_btn.set_state('▶   EXECUTE SCAN', enabled=True)
        self._status_lbl.config(text='◎  COMPLETE', fg=GREEN)
        self._sb_lbl.config(text='Scan complete  •  ' +
                            datetime.now().strftime('%H:%M:%S'))
        if self._result:
            self._open_btn.config(state='normal')

    def _open_result(self):
        if self._open_btn.cget('state') == 'disabled': return
        if self._result and os.path.exists(self._result):
            if sys.platform == 'win32':    os.startfile(self._result)
            elif sys.platform == 'darwin': subprocess.Popen(['open', self._result])
            else:                          subprocess.Popen(['xdg-open', self._result])


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  ENTRY POINT
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if __name__ == '__main__':
    try:
        App().mainloop()
    except Exception:
        import traceback
        err = traceback.format_exc()
        try:
            root = tk.Tk()
            root.title('Startup Error')
            root.geometry('640x360')
            root.configure(bg='#080A0F')
            tk.Label(root, text='❌  Failed to start — error below:',
                     fg='#E05555', bg='#080A0F',
                     font=('Georgia',11,'bold')).pack(pady=(16,6), padx=18, anchor='w')
            t = tk.Text(root, bg='#0D1018', fg='#D8DCE8',
                        font=('Courier New',9), relief='flat', bd=12)
            t.pack(fill='both', expand=True, padx=18, pady=(0,8))
            t.insert('end', err)
            t.insert('end', '\n\nFix:\n  pip install playwright openpyxl\n  playwright install chromium')
            t.configure(state='disabled')
            tk.Button(root, text='Close', command=root.destroy,
                      bg='#E05555', fg='white', relief='flat',
                      padx=18, pady=6).pack(pady=(0,16))
            root.mainloop()
        except Exception:
            with open(os.path.join(os.path.dirname(__file__),'centris_error.log'),'w') as f:
                f.write(err)
