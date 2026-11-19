"""
Business Booms & Depressions, 1945-2025 -- a continuation of the 1944 chart "Business Booms & Depressions since 1775".

What it draws (all from one Excel workbook, read by column letter -- see COLS below):
 * frame about 6.5 : 1 with a presidents row and a year row on top, one grid column per year
 * business volume (black area): % deviation of industrial production from its centred 10-year average
 * commodity prices (dashed line, green / red areas): % change of the producer price index over 12 months
 * federal debt, national income, Dow Jones and bond yield on hand-shaped scales, one scale per line
 * wars as shaded bands, CPI "pillars" (what a base-year dollar bought at the end year), and labels placed
   by a collision check so that they do not sit on each other or on the lines
 * an optional QR code under the chart (QR_URL)

Run:    python chart_original_style.py [workbook.xlsx]
Needs:  pip install -r requirements.txt   (pandas, openpyxl, matplotlib, numpy, qrcode)
Fonts:  put static .ttf files (e.g. Oswald-Bold.ttf) in this folder or in a "fonts" folder next to the script.
Writes: chart_original_style.png and chart_original_style.pdf
"""
import glob, math, os, sys, textwrap
import numpy as np
import pandas as pd
import matplotlib
matplotlib.rcParams['pdf.fonttype'] = 42      # embed real TrueType fonts in the PDF (best for printers)
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patheffects as pe
from matplotlib import font_manager as fm
from matplotlib.lines import Line2D
from matplotlib.patches import Patch, Rectangle
from matplotlib.text import Text
from matplotlib.transforms import Bbox

# =====================================================================
# 1. SETTINGS YOU CAN CHANGE
# =====================================================================
XLSX = sys.argv[1] if len(sys.argv) > 1 else "us_stats_02test.xlsx"
OUT = "chart_original_style"
START, END = pd.Timestamp("1945-01-01"), pd.Timestamp("2026-01-01")

# ---- page layout (inches).  FIG_W and FRAME_RATIO decide the shape of the chart itself.
FIG_W = 40.0               # width of the whole page
FRAME_RATIO = 6.5          # frame width : frame height (sketch: 8:1; the original is about 9.3:1; lower = taller)
PAGE_MARGIN_IN = 0.45      # white margin left and right of the outline (the legend sits in the right one)
OUTLINE_GAP_IN = 0.07      # white gap between the chart and its outline (left, right, bottom; none at the top)
OUTLINE_LW = 3.2           # thickness of the outline, points
LEGEND_W_IN = 0.26         # legend column outside the outline, on the right
MARGIN_L = MARGIN_R = (PAGE_MARGIN_IN + OUTLINE_GAP_IN) / FIG_W
TITLE_IN = 0.35            # empty space above the frame
TEXT_BELOW_IN = None       # height of the text strip: worked out from the text (or set a number of inches)

# ---- text strip: the text itself is in TEXT_FILE (plain text: "# HEADING", paragraphs, **bold**, "@@ source", tables)
TEXT_FILE = "chart_text.md"
LEFT_BLOCK = "INFLATION–DEFLATION AREAS"   # this block flows over LEFT_COLS columns on the left; the others go right
LEFT_COLS = 3
TITLE = "BUSINESS BOOMS & DEPRESSIONS"
YEARS_TEXT = "1945–2025"
CENTRE_SHARE = 0.19        # share of the width for the title block in the middle (about as in the original)
MAX_RIGHT_RATIO = 1.25     # no right-hand block wider than this times the average
TABLE_GAP_IN = 0.18        # gap between the two halves of the debt table
TITLE_WIDTH_SHARE = 0.85   # the title fills at most this share of the centre
TITLE_MIN_PT = 30.0        # it may shrink to this so that the title group is no taller than the text
COL_GAP_IN, BLOCK_GAP_IN = 0.22, 0.45
TEXT_GAP_TOP_IN = 0.62     # gap between the chart and the text (the outline and the small print sit in it)
TEXT_MARGIN_BOTTOM_IN = 0.35
TEXT_PT = 5.6              # body text size on the 40-inch page (printed 1.5 m wide everything is about 1.48 times larger)
TEXT_LEADING = 1.2
HEAD_SCALE = 3.1           # block headings, relative to the body text
STRIPE_RED, STRIPE_IN = "#D9534F", 0.055
LEGEND_FS = 5.0
TEXT_FONTS = dict(body="texgyrebonum-regular.ttf", bold="texgyrebonum-bold.ttf", italic="texgyrebonum-italic.ttf",
                  heading="LibreFranklin-Bold.ttf", title="LibreFranklin-ExtraBold.ttf", since="Jost-MediumItalic.ttf")

# ---- QR code + "how it was made" lines, under the chart in the middle (set QR_URL = None to leave them out)
QR_URL = "https://enricoronca.github.io/Business-Booms-Depressions-1945-2025/"     # the "how it was made" page; the code is drawn as vector squares
QR_SIZE_IN = 0.6           # width of the QR code on the page, in inches
QR_LINES = ["CONTINUED 1945-2025 FROM THE 1944 ORIGINAL, BUILT WITH CLAUDE (ANTHROPIC), NOVEMBER 2026",
            "37 VERSIONS  ·  4,071 PLOTTED VALUES  ·  1,239 LINES OF CODE",
            "SCAN FOR THE DATA, THE CODE AND HOW IT WAS MADE"]
PRES_FRAC, YEAR_FRAC = 0.072, 0.060 # height of the presidents row / year row, as a share of the frame
BASE_LW = 1.1              # thickness of the two baselines (and of the axis name boxes that sit on one)
NGRID = 12                 # number of horizontal grid rows in the plot
PNG_DPI = 450              # PNG resolution on the 40-inch page (450 = about 300 dpi when printed 1.5 m wide)
FONT_SCALE = 1.0           # multiply every font size

# ---- commodity prices: PPIACO is already 1982 = 100 (workbook column O), plotted LINEAR
COMM_SMOOTH = 3            # months of smoothing of the 12-month change (1 = none)
BORDER_LW = 0.9             # extra width of the yellow outline around each commodity dash over the black area
COMM_BAND = "#F3E07A"         # colour of that band (same yellow as the background)
RED_MIN_MONTHS = 3          # shorter red spells are left green
RED_YOY = 10.0             # the part of the curve above this % a year is drawn red

# ---- hand-shaped axes.  A label sits ON a grid row: row 0 = bottom line, row NGRID = top line.
AXBOX_ROWS = 1.5             # height of the name box beside each scale, in grid rows
AXES = {
    "money":     dict(labels=[0, 100, 250, 500, 1000, 2000, 4000, 7000, 10000, 15000, 20000, 30000, 40000],
                      first_row=0, fmt="{:,.0f}", box="FEDERAL DEBT &\nNATIONAL INCOME\n$ BILLIONS"),
    # business volume: one grid row per 5 points, zero on row 8; values beyond the labels are cut off at the edge
    "business":  dict(labels=list(range(-30, 25, 5)), first_row=2, fmt="{:+.0f}", show=lambda v: v % 10 == 0,
                      box=None),
    # commodity prices: % change over 12 months, zero on row 2; 10 points per row below zero, 5 per row above
    "commodity": dict(labels=[-20, -10, 0, 5, 10, 15, 20, 25, 30, 35], first_row=0, fmt="{:+.0f}", show=lambda v: v % 10 == 0,
                      box=None),
    "dow":       dict(labels=[0, 100, 250, 500, 1000, 2500, 5000, 10000, 20000, 30000, 40000, 50000, 60000],
                      first_row=0, fmt="{:,.0f}", box="INDUSTRIAL STOCK PRICES\n(DOW JONES AVERAGE)"),
    "yield":     dict(labels=[0, 0.04, 0.08, 0.12, 0.16], first_row=3, invert=True, fmt="{:.0%}",
                      box="BOND YIELD"),
}
WHITE_AXES = {"money", "dow", "yield"}            # drawn as white one-year strips; the others stay yellow
# One scale per line, each at its own place (year where the scale starts).  White strips snap to a year column.
# The name box is attached to the RIGHT side of the scale, near the bottom.
AXIS_X = {"commodity": 1957.0, "money": 1964.0, "yield": 1985.0, "dow": 1996.0, "business": 2013.0}
AXBOX_Y0_ROWS = 0.5          # bottom of the name boxes, in grid rows (0.5 + 1.5 = top exactly on grid line 2)
AXBOX_W_YEARS = 2.4        # width of the name box, in years

# ---- wars: shaded in the band under the commodity baseline (name, start year, end year, height of its label in the band)
WARS = [("WORLD WAR II", 1945.0, 1945.7, 0.5), ("KOREAN WAR", 1950.5, 1953.6, 0.5), ("VIETNAM WAR", 1965.0, 1973.3, 0.5),
        ("GULF WAR", 1990.6, 1991.2, 0.5), ("WAR IN AFGHANISTAN", 2001.8, 2021.6, 0.74), ("IRAQ WAR", 2003.2, 2011.9, 0.26),
        ("WAR IN UKRAINE", 2022.1, 2025.99, 0.5)]
WAR_COLOR = "#C2958C"      # dusty rose, as the war bands in the original; overlapping wars come out darker
WAR_ALPHA = 0.75
WAR_FS = 8.0               # size of the war names

# ---- labels on the lines: "moments" (series, year).  debt / income values come from the workbook.
MOMENTS = ([("debt", y) for y in (1947, 1960, 1981, 1995, 2009, 2017, 2025)] +
           [("income", y) for y in (1950, 1960, 1969, 1980, 1990, 2000, 2025)] +
           [("dow", y) for y in (1972.9, 1999.9, 2002.75, 2009.2, 2015.4, 2020.9)] + [("dow", 1987.76, "BLACK MONDAY CRASH")] +
           [("yield", y) for y in (1954.3, 1966.7, 1981.7, 2012.6, 2016.5, 2020.6, 2025.8)])
# labels to push further below their point: (series, year) -> minimum drop, as a share of the plot height
MOMENT_DROP = {("yield", 2016.5): 0.13, ("dow", 2002.75): 0.13}
# series names with arrows: (text, series, candidate years to try, in order)
NAME_LABELS = []   # (old fixed-position name labels, replaced by the labels next to each axis)
_OLD_NAME_LABELS = [("BUSINESS VOLUME", "bus", [1956.3, 1964.0, 1977.5, 1996.0, 2006.5]),
               ("BOND YIELD", "yield", [1950.0, 1958.0, 1967.0, 1985.0]),
               ("NATIONAL INCOME", "income", [1952.5, 1957.5, 1966.5, 1975.5]),
               ("FEDERAL DEBT", "debt", [1951.5, 1956.5, 1965.5, 1976.5]),
               ("DOW JONES", "dow", [1959.5, 1968.0, 1979.0, 1989.5]),
               ("COMMODITY PRICES 1982 = 100", "ppi", [1948.0, 1955.0, 1962.0]),
               ("COMMODITY PRICES", "ppi", [2024.3, 2023.5, 2022.0])]
EXTRA_EVENTS = [("1955-57", "Auto & Capital-Spending Boom"), ("1972-73", "Wage-Price Controls Boom"),
                ("1976-79", "Late-1970s Recovery")]
UP_WORDS = ("boom", "expansion", "recovery", "rebound")     # events with these words point to a PEAK (label above)
EVENT_DX = {"OPEC Oil Embargo Recession": +1}
EVENT_DXMAX = {"OPEC Oil Embargo Recession": 3.0}   # keep these labels close to their point (years)   # -1 = keep the label LEFT of its point, +1 = right
EVENT_SIDE = {"Dot-com Bust & 9/11": +1, "Postwar Readjustment": +1}   # +1 = label above the business-volume curve, -1 = below
RED_LABEL_COUNT = 3        # how many of the longest red inflation periods get a "PRICE INFLATION AREA" label

# ---- where things sit in the workbook (column letters; row 1 = headers)
COLS = dict(year="A", debt="B", income="D",
            mdate="F", dow="G", yld="I",
            ppi_date="M", ppi="O",
            bus_date="AX", bus="AY",
            pill_first="AG", pill_last="AM",            # base year, end year, then five pillar columns
            ev_period="AO", ev_name="AP")               # event labels: "1973-75" | "OPEC Oil Embargo Recession"
PILLAR_RENAME = {"TOTAL COST OF LIVING": "TOTAL C. of L."}   # header text -> text printed on the pillar
PILLAR_W_YEARS, PILLAR_GAP_YEARS = 0.32, 0.06     # bar width and gap, in years (thinner than before)
PILLAR_ARROW_GAP = 0.35   # years kept free between an arrow head and another pillar block
CAPTION_EDGE_GAP = 1.45   # years kept free right of a caption near the edge (room for the last war name)
PILLAR_EDGE_GAP = 0.9     # years kept free between a caption and the right edge of the chart
PILLAR_MAX_SHIFT = 0.8    # how far (years) a block may slide left to make room for an arrow
PILLAR_100C_IN = 0.80                             # height of a 100-cent bar, in inches (bars stand on the 0% line)

# ---- fonts: the first one found is used.  Static .ttf files in this folder (or ./fonts) are picked up automatically.
HEAD_FONTS = ["Franklin Gothic Demi Cond", "Franklin Gothic Medium Cond", "League Gothic", "Oswald",
              "Bahnschrift SemiBold Condensed", "Anton", "Roboto Condensed", "Arial Narrow", "Impact"]

PROVISIONAL_FROM = "2021-10-01"      # business activity after this date uses a shortened averaging window

# =====================================================================
# 2. READ THE WORKBOOK (by column letter)
# =====================================================================
raw = pd.read_excel(XLSX, header=None)          # row 0 of this table = row 1 of the sheet (headers)

def cidx(letter):
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

def col(letter, n=None):
    s = raw.iloc[1:, cidx(letter)]
    return s.dropna() if n is None else s.iloc[:n]

years = col(COLS["year"]).astype(int).reset_index(drop=True)
n_a = len(years)
annual_dates = pd.to_datetime(years.astype(str) + "-07-01")                 # plotted at mid-year
debt = pd.to_numeric(col(COLS["debt"], n_a)).reset_index(drop=True)
income = pd.to_numeric(col(COLS["income"], n_a)).reset_index(drop=True)

m_dates = pd.to_datetime(col(COLS["mdate"])).reset_index(drop=True)
n_m = len(m_dates)
dow = pd.to_numeric(col(COLS["dow"], n_m)).reset_index(drop=True)
yld = pd.to_numeric(col(COLS["yld"], n_m)).reset_index(drop=True)
p_dates = pd.to_datetime(col(COLS["ppi_date"])).reset_index(drop=True)
ppi = pd.to_numeric(col(COLS["ppi"], len(p_dates))).reset_index(drop=True)
b_dates = pd.to_datetime(col(COLS["bus_date"])).reset_index(drop=True)
bus = pd.to_numeric(col(COLS["bus"], len(b_dates))).reset_index(drop=True)
if bus.isna().any():
    sys.exit("The business-activity column has empty cells. Open the workbook in Excel and save it once, then re-run.")

def trim(dates, values):
    keep = (dates >= START) & (dates < END)               # drop months after the chart's end
    return dates[keep].reset_index(drop=True), values[keep].reset_index(drop=True)

m_dates_t, dow = trim(m_dates, dow)
_, yld = trim(m_dates, yld)
m_dates = m_dates_t
_, ppi_yoy = trim(p_dates, ppi.pct_change(12) * 100)    # 12-month change, computed before the start is cut
p_dates, ppi = trim(p_dates, ppi)
b_dates, bus = trim(b_dates, bus)
print(f"{len(m_dates)} monthly rows, {len(p_dates)} PPI rows, {len(b_dates)} business-activity rows, {n_a} annual rows")

# pillars (live formulas must have been calculated: save the workbook in Excel once)
pf, pl = cidx(COLS["pill_first"]), cidx(COLS["pill_last"])
tbl = raw.iloc[1:, pf:pl + 1].reset_index(drop=True)
tbl.columns = ["base", "end", "p1", "p2", "p3", "p4", "p5"]
tbl = tbl[pd.to_numeric(tbl.base, errors="coerce").notna() & pd.to_numeric(tbl.end, errors="coerce").notna()]
pillars = tbl if len(tbl) else None
names = [str(raw.iloc[0, pf + k]).strip().upper() for k in range(2, 7)]
PILLAR_LABELS = [PILLAR_RENAME.get(n, n) for n in names]
if pillars is not None:
    print("Pillars read:")
    print(pillars.round(1).to_string(index=False))

# event labels
events = []
ep, en = cidx(COLS["ev_period"]), cidx(COLS["ev_name"])
for per, nm in zip(raw.iloc[1:, ep], raw.iloc[1:, en]):
    if isinstance(per, str) and isinstance(nm, str) and "-" in per:
        a, b = per.strip().split("-")
        y0 = int(a)
        y1 = int(b) if len(b) == 4 else int(a[:2] + b.zfill(2))
        events.append(dict(start=pd.Timestamp(f"{y0}-01-01"), end=pd.Timestamp(f"{y1}-12-31"), name=nm.strip()))
for per, nm in EXTRA_EVENTS:                      # extra labels added in the script (move them to the sheet if you like)
    a, b = per.split("-")
    y1 = int(b) if len(b) == 4 else int(a[:2] + b.zfill(2))
    events.append(dict(start=pd.Timestamp(f"{int(a)}-01-01"), end=pd.Timestamp(f"{y1}-12-31"), name=nm))
print(f"{len(events)} event labels read")

# =====================================================================
# 3. HELPERS
# =====================================================================
YELLOW, LGREEN, DGREEN, RED = "#F3E07A", "#A8D5A2", "#3E9B5A", "#E2735D"
DEBT_C, INC_C = "#C1272D", "#1D7A3E"
try:
    HERE = os.path.dirname(os.path.abspath(__file__))
except NameError:
    HERE = os.getcwd()
for pat in ("*.ttf", "*.otf", os.path.join("fonts", "*.ttf"), os.path.join("fonts", "*.otf")):
    for f in glob.glob(os.path.join(HERE, pat)) + glob.glob(os.path.join(os.getcwd(), pat)):
        try:
            fm.fontManager.addfont(f)
        except Exception:
            pass
_avail = {f.name for f in fm.fontManager.ttflist}
HEAD_FONT = next((f for f in HEAD_FONTS if f in _avail), "DejaVu Sans")
print(f"Heading font: {HEAD_FONT}" + ("" if HEAD_FONT != "DejaVu Sans" else
      "  (no condensed font found: put Oswald-Bold.ttf or LeagueGothic-Regular.ttf next to this script)"))
FS = lambda x: x * FONT_SCALE
HALO = [pe.withStroke(linewidth=2.2, foreground=YELLOW)]       # thin yellow outline so numbers stay readable over lines

frame_w_in = FIG_W * (1 - MARGIN_L - MARGIN_R)
frame_h_in = frame_w_in / FRAME_RATIO
plot_h_in = frame_h_in * (1 - PRES_FRAC - YEAR_FRAC)
# ---- text strip: measured before the page is made, so that the page height follows the text
def font_file(name):
    for d in (os.path.join(HERE, "fonts"), HERE, os.path.join(os.getcwd(), "fonts"), os.getcwd()):
        p_ = os.path.join(d, name)
        if os.path.exists(p_):
            return p_
    print(f"font {name} not found: using a default font instead")
    return None

_fp_cache = {}
def fprop(name, size):
    key = (name, round(size, 3))
    if key not in _fp_cache:
        f = font_file(name)
        _fp_cache[key] = fm.FontProperties(fname=f, size=size) if f else fm.FontProperties(family="serif", size=size)
    return _fp_cache[key]

F_BODY, F_BOLD, F_ITAL = TEXT_FONTS["body"], TEXT_FONTS["bold"], TEXT_FONTS["italic"]
F_HEAD, F_TITLE, F_SINCE = TEXT_FONTS["heading"], TEXT_FONTS["title"], TEXT_FONTS["since"]
_mfig = plt.figure(dpi=144)
_tr = _mfig.canvas.get_renderer()
_wcache = {}

def text_w(s_, font, size):
    """width of a string in inches"""
    key = (s_, font, round(size, 3))
    if key not in _wcache:
        w, _, _ = _tr.get_text_width_height_descent(s_, fprop(font, size), ismath=False)
        _wcache[key] = w / _mfig.dpi
    return _wcache[key]

def read_blocks(path):
    """text file: "# HEADING", paragraphs, **bold**, "@@ source line", "| table | rows |" """
    blocks_, cur = {}, None
    for raw_ln in open(path, encoding="utf-8").read().split("\n\n"):
        ln = raw_ln.strip()
        if not ln:
            continue
        if ln.startswith("# "):
            head, _, rest = ln.partition("\n")
            cur = head[2:].strip()
            blocks_[cur] = []
            ln = rest.strip()
            if not ln:
                continue
        if ln.startswith("|"):
            blocks_[cur].append(("table", [[c.strip() for c in r.strip().strip("|").split("|")] for r in ln.split("\n")]))
        elif ln.startswith("@@"):
            blocks_[cur].append(("src", ln[2:].strip()))
        else:
            blocks_[cur].append(("p", " ".join(ln.split("\n"))))
    return blocks_

def words_of(text, base_font, bold_font):
    out, bold = [], False
    for k, part in enumerate(text.split("**")):
        if k:
            bold = not bold
        for w in part.split():
            out.append((w, bold_font if bold else base_font))
    return out

def layout(items, width, size):
    """break paragraphs into lines for a column of `width` inches; returns line records"""
    lh = size * TEXT_LEADING / 72
    lines = []
    for kind, content in items:
        if kind == "table":
            # as in the original: the rows are set in two side-by-side halves, the total at the foot of the right half
            lines.append(dict(kind="gap", h=lh * 0.25))
            head = content[0] if not content[0][0] else None
            rows = content[1:] if head else content
            total = [r for r in rows if "**" in r[0]]
            data = [r for r in rows if "**" not in r[0]]
            k = (len(data) + 1) // 2
            left_rows, right_rows = data[:k], data[k:] + total
            half = (width - TABLE_GAP_IN) / 2
            def cell(r):
                if r is None:
                    return None
                return dict(label=r[0].replace("**", ""), val=r[1].replace("**", "") if len(r) > 1 else "",
                            font=F_BOLD if "**" in r[0] else F_BODY)
            if head:
                lines.append(dict(kind="row2", h=lh * 0.9, size=size * 0.82, half=half, head=head[1].replace("**", ""),
                                  cells=(None, None)))
            for i in range(max(len(left_rows), len(right_rows))):
                lines.append(dict(kind="row2", h=lh * 0.98, size=size * 0.92, half=half, head=None,
                                  cells=(cell(left_rows[i]) if i < len(left_rows) else None,
                                         cell(right_rows[i]) if i < len(right_rows) else None)))
            lines.append(dict(kind="gap", h=lh * 0.35))
            continue
        if kind == "src":
            sz, fnt, fbold, indent, justify = size * 0.8, F_ITAL, F_ITAL, 0.0, False
            lines.append(dict(kind="gap", h=lh * 0.2))
        else:
            sz, fnt, fbold, indent, justify = size, F_BODY, F_BOLD, size * 1.4 / 72, True
        sp = text_w("a a", fnt, sz) - text_w("aa", fnt, sz)
        cur, cur_w, first, avail = [], 0.0, True, width - indent
        for w, f in words_of(content, fnt, fbold):
            ww = text_w(w, f, sz)
            need = ww if not cur else cur_w + sp + ww
            if cur and need > avail:
                lines.append(dict(kind="text", h=sz * TEXT_LEADING / 72, words=cur, size=sz, indent=indent if first else 0,
                                  width=width, justify=justify))
                first, cur, avail, need = False, [], width, ww
            cur.append((w, f, ww))
            cur_w = need
        if cur:
            lines.append(dict(kind="text", h=sz * TEXT_LEADING / 72, words=cur, size=sz, indent=indent if first else 0,
                              width=width, justify=False))
        if kind == "src":
            lines.append(dict(kind="gap", h=lh * 0.15))
    return lines

def height(lines):
    return sum(l["h"] for l in lines)

def flow(lines, ncols, col_h):
    cols, cur, hh = [], [], 0.0
    for l in lines:
        if hh + l["h"] > col_h + 1e-9 and cur:
            cols.append(cur)
            cur, hh = [], 0.0
            if l["kind"] == "gap":
                continue
        cur.append(l)
        hh += l["h"]
    cols.append(cur)
    return cols if len(cols) <= ncols else None

def head_h(size):
    return size * HEAD_SCALE * 1.5 / 72

def fits(items, w, H, ncols):
    ls_ = layout(items, w, TEXT_PT)
    return height(ls_) <= H if ncols == 1 else flow(ls_, ncols, H) is not None

def min_width(items, H, ncols, cap):
    lo, hi = 0.6, cap
    if not fits(items, hi, H, ncols):
        return None
    for _ in range(16):
        mid = (lo + hi) / 2
        lo, hi = (lo, mid) if fits(items, mid, H, ncols) else (mid, hi)
    return hi

blocks = read_blocks(os.path.join(HERE, TEXT_FILE) if not os.path.exists(TEXT_FILE) else TEXT_FILE)
RIGHT = [h for h in blocks if h not in (LEFT_BLOCK, "SMALL PRINT")]
OUT_L = PAGE_MARGIN_IN                              # outline, inches from the left of the page
OUT_R = FIG_W - PAGE_MARGIN_IN
TW = OUT_R - OUT_L
# Widths as in the original: the title block gets CENTRE_SHARE of the width; the rest is shared between the blocks and
# adjusted step by step until they all end at about the same depth.
def block_h(name, w):
    ls_ = layout(blocks[name], w, TEXT_PT)
    if name != LEFT_BLOCK:
        return height(ls_)
    lo, hi = 0.2, 20.0                                   # smallest column height that takes the block in LEFT_COLS columns
    for _ in range(18):
        mid = (lo + hi) / 2
        lo, hi = (lo, mid) if flow(ls_, LEFT_COLS, mid) is not None else (mid, hi)
    return hi

# The title block sits in the exact middle of the page, so the left block and the four right blocks get the same width.
# The right blocks share their width so that they end at about the same depth.
CENTRE_W = TW * CENTRE_SHARE
side = (TW - CENTRE_W - 2 * BLOCK_GAP_IN) / 2
COL_W_LEFT = (side - (LEFT_COLS - 1) * COL_GAP_IN) / LEFT_COLS
r_room = side - (len(RIGHT) - 1) * COL_GAP_IN
wr = np.full(len(RIGHT), r_room / len(RIGHT))
for _ in range(40):
    hs_r = np.array([block_h(nm, w) for nm, w in zip(RIGHT, wr)])
    wr = wr * (hs_r / hs_r.mean()) ** 0.6
    wr = np.minimum(wr, MAX_RIGHT_RATIO * r_room / len(RIGHT))
    wr = wr * r_room / wr.sum()
COL_WS = [float(w) for w in wr]
hs = np.array([block_h(LEFT_BLOCK, COL_W_LEFT)] + [block_h(nm, w) for nm, w in zip(RIGHT, COL_WS)])
TEXT_H = float(hs.max())
# title block in the middle: title as large as fits its width, made smaller if the group would be taller than the text
sp_paras = [c for _, c in blocks.get("SMALL PRINT", [])]      # each paragraph of the small print is one line
sp_text = " ".join(sp_paras)
def centre_stack(t_size):
    t_w = text_w(TITLE, F_TITLE, t_size)
    sp_l = []                                     # the small print now runs under the outline
    qr_h = (QR_SIZE_IN + 0.08 + sum(TEXT_PT * (1.15 if i == 0 else 1.0) * 1.4 / 72 for i in range(len(QR_LINES)))) if QR_URL else 0
    return 0.10 + t_size * 0.72 / 72 + 0.11 + 0.20 + qr_h, t_w, sp_l

T_SIZE = 60.0
while text_w(TITLE, F_TITLE, T_SIZE) > CENTRE_W * TITLE_WIDTH_SHARE and T_SIZE > 8:
    T_SIZE -= 0.5
while centre_stack(T_SIZE)[0] > head_h(TEXT_PT) + TEXT_H and T_SIZE > TITLE_MIN_PT:
    T_SIZE -= 0.5
STACK_H, T_W, SP_LINES = centre_stack(T_SIZE)
STRIP_H = max(head_h(TEXT_PT) + TEXT_H, STACK_H)
if TEXT_BELOW_IN is None:
    TEXT_BELOW_IN = TEXT_GAP_TOP_IN + STRIP_H + TEXT_MARGIN_BOTTOM_IN
print(f"TEXT: {TEXT_PT} pt, heights " + ", ".join(f"{h_:.2f}" for h_ in hs) + f" in, text height {TEXT_H:.2f} in, left columns {COL_W_LEFT:.2f} in, right columns "
      + ", ".join(f"{w_:.2f}" for w_ in COL_WS) + f" in, centre {CENTRE_W:.2f} in, strip {TEXT_BELOW_IN:.2f} in")

FIG_H = TITLE_IN + frame_h_in + TEXT_BELOW_IN
YEAR_IN = frame_w_in / ((END - START).days / 365.25)             # inches per year on the page

def heights(name):
    spec = AXES[name]
    rows = spec["first_row"] + np.arange(len(spec["labels"]))
    if spec.get("invert"):
        rows = rows[::-1]
    return rows / NGRID

def pos(name, values):
    """real values -> plot heights (0-1) on the hand-shaped axis `name`"""
    spec = AXES[name]
    lab = np.asarray(spec["labels"], float)
    assert np.all(np.diff(lab) > 0), f"labels of '{name}' must increase"
    v = np.asarray(values, float)
    if name not in ("business", "commodity") and (np.nanmin(v) < lab[0] or np.nanmax(v) > lab[-1]):
        print(f"WARNING: '{name}' has values outside its labels ({np.nanmin(v):.4g} to {np.nanmax(v):.4g})")
    return np.interp(v, lab, heights(name))

def fmt_val(spec, v):
    t = spec["fmt"].format(v)
    return t.replace("+", "") if v == 0 else t

def lab_ok(spec, v, h):
    return 0.001 < h < 0.999 and (spec.get("show") is None or spec["show"](v))     # no zeros on the bottom line

def ts(year_float):
    y = int(year_float)
    return pd.Timestamp(f"{y}-01-01") + pd.Timedelta(days=(year_float - y) * 365.25)

def num(t):
    return mdates.date2num(t)

def yf(t):
    return t.year + (t.dayofyear - 1) / 365.25

# ---- where the scale groups go (computed once; used by the drawing code and by the label placer)
def layout_groups():
    cols, reserved = [], []
    for nm, x in AXIS_X.items():
        x0 = float(math.ceil(x - 1e-9)) if nm in WHITE_AXES else x     # white strips start exactly on a year line
        cols.append(dict(nm=nm, white=nm in WHITE_AXES, x0=x0, x1=x0 + 1.0))
        reserved.append((x0 - 0.15, x0 + 1.0 + 0.15))                 # the scale itself, full height
    return cols, reserved

scale_cols, reserved = layout_groups()

# =====================================================================
# 4. FIGURE
# =====================================================================
fig = plt.figure(figsize=(FIG_W, FIG_H))
fx0 = MARGIN_L
fw = 1 - MARGIN_L - MARGIN_R
y_frame_bottom = TEXT_BELOW_IN / FIG_H
h_plot = plot_h_in / FIG_H
h_year = frame_h_in * YEAR_FRAC / FIG_H
h_pres = frame_h_in * PRES_FRAC / FIG_H
base = fig.add_axes([fx0, y_frame_bottom, fw, h_plot])
yrs = fig.add_axes([fx0, y_frame_bottom + h_plot, fw, h_year])
pres = fig.add_axes([fx0, y_frame_bottom + h_plot + h_year, fw, h_pres])
for a in (base, yrs, pres):
    a.set_xlim(START, END)
    a.set_ylim(0, 1)
for a in (yrs, pres):
    a.axis("off")
base.set_facecolor(YELLOW)
base.set_xticks([]); base.set_yticks([])
for sp in base.spines.values():
    sp.set_visible(False)
fig.add_artist(Rectangle((fx0, y_frame_bottom), fw, h_plot + h_year + h_pres, transform=fig.transFigure,
                         fill=False, ec="k", lw=1.0, zorder=50))
# outer outline: white gap left, right (around the legend) and bottom, none at the top
_ft = y_frame_bottom + h_plot + h_year + h_pres
fig.add_artist(Rectangle((OUT_L / FIG_W, y_frame_bottom - OUTLINE_GAP_IN / FIG_H), (OUT_R - OUT_L) / FIG_W,
                         _ft - y_frame_bottom + OUTLINE_GAP_IN / FIG_H, transform=fig.transFigure,
                         fill=False, ec="k", lw=OUTLINE_LW, zorder=50, joinstyle="miter"))

# ---- presidents row (top)
TERMS = [("Harry S.", "Truman", "Democrat", "1945-01-01", "1953-01-20"),
         ("Dwight D.", "Eisenhower", "Republican", "1953-01-20", "1961-01-20"), ("John F.", "Kennedy", "Democrat", "1961-01-20", "1963-11-22"),
         ("Lyndon B.", "Johnson", "Democrat", "1963-11-22", "1969-01-20"), ("Richard M.", "Nixon", "Republican", "1969-01-20", "1974-08-09"),
         ("Gerald R.", "Ford", "Republican", "1974-08-09", "1977-01-20"), ("Jimmy", "Carter", "Democrat", "1977-01-20", "1981-01-20"),
         ("Ronald", "Reagan", "Republican", "1981-01-20", "1989-01-20"), ("George H. W.", "Bush", "Republican", "1989-01-20", "1993-01-20"),
         ("Bill", "Clinton", "Democrat", "1993-01-20", "2001-01-20"), ("George W.", "Bush", "Republican", "2001-01-20", "2009-01-20"),
         ("Barack", "Obama", "Democrat", "2009-01-20", "2017-01-20"), ("Donald", "Trump", "Republican", "2017-01-20", "2021-01-20"),
         ("Joe", "Biden", "Democrat", "2021-01-20", "2025-01-20"), ("D.", "Trump", "Republican", "2025-01-20", "2026-01-01")]
_pr = fig.canvas.get_renderer()

def snap(t):                                   # nearest 1 January, so the dividers sit exactly on the year lines
    t = pd.Timestamp(t)
    return pd.Timestamp(f"{t.year + (1 if t.month >= 7 else 0)}-01-01")

for first, name, party, s_, e_ in TERMS:
    s_, e_ = max(snap(s_), START), min(snap(e_), END)
    if e_ <= s_:
        continue
    pres.axvspan(s_, e_, facecolor=YELLOW, edgecolor="k", lw=0.9)
    if not name:
        continue
    mid = s_ + (e_ - s_) / 2
    if (e_ - s_).days < 500:                                  # too narrow: surname sideways
        pres.text(mid, 0.5, name.upper(), ha="center", va="center", fontsize=FS(8), fontweight="bold",
                  rotation=90, family=HEAD_FONT)
    else:
        # first name next to the surname on one line; shrink, then use initials, then the surname alone, until it fits
        width_px = pres.transData.transform((num(e_), 0))[0] - pres.transData.transform((num(s_), 0))[0]
        initials = " ".join(w[0] + "." for w in first.replace(".", " ").split())
        label, size = name.upper(), 7.0
        done = False
        for cand in (f"{first} {name}", f"{initials} {name}", name):
            for fs_ in (11.5, 10.5, 9.5, 8.5, 7.5):
                t_ = pres.text(0, 0, cand.upper(), fontsize=FS(fs_), fontweight="bold", family=HEAD_FONT)
                w_ = t_.get_window_extent(_pr).width
                t_.remove()
                if w_ <= width_px * 0.92:
                    label, size, done = cand.upper(), fs_, True
                    break
            if done:
                break
        pres.text(mid, 0.66, label, ha="center", va="center", fontsize=FS(size), fontweight="bold", family=HEAD_FONT)
        pres.text(mid, 0.2, party, ha="center", va="center", fontsize=FS(7), family=HEAD_FONT)

# ---- year row: yellow boxes, one per year, exactly over the year columns of the grid
for y in range(START.year, END.year):
    a, b = pd.Timestamp(f"{y}-01-01"), pd.Timestamp(f"{y + 1}-01-01")
    yrs.axvspan(a, b, facecolor="white", edgecolor="k", lw=0.8)
    yrs.text(a + (b - a) / 2, 0.5, str(y), ha="center", va="center", fontsize=FS(10), fontweight="bold",
             family=HEAD_FONT)

# ---- grid: one vertical line per year (heavier each decade), NGRID+1 horizontal lines
for yr in range(START.year, END.year + 1):
    heavy = yr % 10 == 0
    base.axvline(pd.Timestamp(f"{yr}-01-01"), color="k", lw=0.9 if heavy else 0.4,
                 alpha=0.85 if heavy else 0.55, zorder=1)
for r in range(0, NGRID + 1):
    base.axhline(r / NGRID, color="k", lw=0.4, alpha=0.55, zorder=1)
zero_h = float(pos("business", [0])[0])
base_h = float(pos("commodity", [0])[0])
base.axhline(zero_h, color="k", lw=BASE_LW, zorder=4.45)   # business-volume baseline drawn over the commodity area: the red stops exactly at it
base.axhline(base_h, color="k", lw=BASE_LW, zorder=2.6)   # commodity baseline over the green/red/war areas, so it never shifts
base.axvspan(pd.Timestamp(PROVISIONAL_FROM), END, color="grey", alpha=0.10, zorder=1)
for _n, _w0, _w1, _f in WARS:
    base.axvspan(ts(_w0), ts(_w1), ymin=0, ymax=base_h, facecolor=WAR_COLOR, alpha=WAR_ALPHA, lw=0, zorder=1.6)

# ---- series (all drawn in plot-height units)
bus_h = pos("business", bus)
bus_fill = base.fill_between(b_dates, zero_h, bus_h, color="black", lw=0, zorder=3)

yoy = ppi_yoy.rolling(COMM_SMOOTH, min_periods=1).mean()
valid = yoy.notna().values
comm_h = pos("commodity", yoy.fillna(0).clip(-20, 40))
thr_h = float(pos("commodity", [RED_YOY])[0])
above = valid & (yoy.fillna(0).values > 0)
red = valid & (yoy.fillna(0).values >= RED_YOY)
_i = 0
while _i < len(red):                                  # drop red runs shorter than RED_MIN_MONTHS
    if red[_i]:
        _j = _i
        while _j + 1 < len(red) and red[_j + 1]:
            _j += 1
        if _j - _i + 1 < RED_MIN_MONTHS:
            red[_i:_j + 1] = False
        _i = _j + 1
    else:
        _i += 1
# three clean layers: light green = prices falling, dark green = rising up to RED_YOY, red = rising faster than that
base.fill_between(p_dates, base_h, np.minimum(comm_h, base_h), color=LGREEN, lw=0, zorder=2)
base.fill_between(p_dates, base_h, np.maximum(comm_h, base_h), color=DGREEN, lw=0, zorder=2.1)
# months with more than RED_YOY inflation are red all the way down to the baseline
base.fill_between(p_dates, base_h, np.maximum(comm_h, base_h), where=red, color=RED, lw=0, zorder=2.2)
base.plot(p_dates, np.where(valid, comm_h, np.nan), color="k", ls=(0, (4, 2)), lw=1.0, zorder=4)

base.plot(annual_dates, pos("money", debt), color=DEBT_C, lw=1.9, zorder=6)
base.plot(annual_dates, pos("money", income), color=INC_C, lw=1.9, zorder=6)
base.plot(m_dates, pos("dow", dow), color="k", ls=(0, (6, 2, 1, 2)), lw=1.0, zorder=5)
base.plot(m_dates, pos("yield", yld), color="k", ls=":", lw=1.4, zorder=5)
# where a black line runs over the black business-volume area, draw it again in light yellow so it stays visible (as in the original)
from matplotlib.path import Path as _Path
_bp = [pp for pp in bus_fill.get_paths() if len(pp.vertices)]
_clip = _Path.make_compound_path(*_bp) if _bp else None
if _clip is not None:
    # over the black area: Dow and bond yield switch to yellow dots/dashes; the commodity dashes get a yellow outline
    for _x, _y, _ls, _lw in [(m_dates, pos("dow", dow), (0, (6, 2, 1, 2)), 1.0), (m_dates, pos("yield", yld), ":", 1.4)]:
        _ln, = base.plot(_x, _y, color=YELLOW, ls=_ls, lw=_lw, zorder=5.5)
        _ln.set_clip_path(_clip, base.transData)
    # over the black: each commodity dash gets its own outline in the background yellow (same dash pattern and phase
    # as the main line, so the dashes stay separate and line up exactly with the ones outside the black)
    _e = BORDER_LW / 2                                   # outline thickness on each side, in points
    with matplotlib.rc_context({"lines.scale_dashes": False}):   # dash lengths in points, independent of line width
        _ln, = base.plot(p_dates, np.where(valid, comm_h, np.nan), color=COMM_BAND, lw=1.0 + BORDER_LW, zorder=4.4,
                         ls=(_e, (4 + 2 * _e, 2 - 2 * _e)), dash_capstyle="butt")   # a yellow box around every dash
    _ln.set_clip_path(_clip, base.transData)
    _ln, = base.plot(p_dates, np.where(valid, comm_h, np.nan), color="k", ls=(0, (4, 2)), lw=1.0, zorder=4.5,
                     dash_capstyle="butt")
    _ln.set_clip_path(_clip, base.transData)

# ---- scale columns
for c in scale_cols:
    spec, hts = AXES[c["nm"]], heights(c["nm"])
    if c["white"]:
        # white strip, exactly one year wide; grid lines continue across it and are cut only behind each number
        base.add_patch(Rectangle((num(ts(c["x0"])), 0), num(ts(c["x1"])) - num(ts(c["x0"])), 1, facecolor="white",
                                 edgecolor="k", lw=0.8, zorder=20))
        for r in range(NGRID + 1):
            base.plot([ts(c["x0"]), ts(c["x1"])], [r / NGRID] * 2, color="k", lw=0.4, alpha=0.55, zorder=20.1)
        for v, h in zip(spec["labels"], hts):
            if not lab_ok(spec, v, h):
                continue
            va, dy = ("top", -0.003) if h >= 0.999 else (("bottom", 0.003) if h <= 0.001 else ("center", 0))
            base.text(ts((c["x0"] + c["x1"]) / 2), h + dy, fmt_val(spec, v), fontsize=FS(6.6), ha="center",
                      va=va, zorder=21, family=HEAD_FONT,
                      bbox=dict(boxstyle="square,pad=0.08", fc="white", ec="none"))
    else:
        for v, h in zip(spec["labels"], hts):                           # yellow column: number sits just above its line
            if not lab_ok(spec, v, h):
                continue
            va, dy = ("top", -0.004) if h >= 0.999 else ("bottom", 0.004)
            base.text(ts((c["x0"] + c["x1"]) / 2), h + dy, fmt_val(spec, v), fontsize=FS(6.6), ha="center",
                      va=va, zorder=15, family=HEAD_FONT, path_effects=HALO)
    # name box attached to the LEFT side of the scale, near the bottom (as in the sketch)
    if not spec.get("box"):
        continue
    bx0, bx1 = num(ts(c["x0"] - AXBOX_W_YEARS)), num(ts(c["x0"]))
    by0, bh = AXBOX_Y0_ROWS / NGRID, AXBOX_ROWS / NGRID
    base.add_patch(Rectangle((bx0, by0), bx1 - bx0, bh, facecolor="white", edgecolor="k", lw=BASE_LW, zorder=22,
                           joinstyle="miter"))   # same weight as the baseline, so its top edge is the baseline
    # biggest font that fits inside the box (with a small margin)
    _rr = fig.canvas.get_renderer()
    bw_px = abs(base.transData.transform((bx1, 0))[0] - base.transData.transform((bx0, 0))[0])
    bh_px = abs(base.transData.transform((0, by0 + bh))[1] - base.transData.transform((0, by0))[1])
    for fs_ in np.arange(11.0, 4.9, -0.25):
        t_ = base.text((bx0 + bx1) / 2, by0 + bh / 2, spec["box"], ha="center", va="center", fontsize=FS(fs_),
                       fontweight="bold", family=HEAD_FONT, zorder=23, linespacing=1.1)
        e_ = t_.get_window_extent(_rr)
        if e_.width <= bw_px * 0.88 and e_.height <= bh_px * 0.80:
            break
        t_.remove()

# =====================================================================
# 5. LABELS: pillars first, then a collision-checked placer for everything else
# =====================================================================
renderer = fig.canvas.get_renderer()
def tbox(t):
    t.update_positions(renderer)                # the text box only, without the arrow
    return Text.get_window_extent(t, renderer)

occ = []                                        # display-space boxes that are already taken

def d_box(x0, x1, y0, y1, pad=0):
    (px0, py0), (px1, py1) = (base.transData.transform((num(ts(x0)), y0)), base.transData.transform((num(ts(x1)), y1)))
    return Bbox.from_extents(min(px0, px1) - pad, min(py0, py1) - pad, max(px0, px1) + pad, max(py0, py1) + pad)

for r0, r1 in reserved:
    occ.append(d_box(r0, r1 + 0.1, 0, 1))
for c in [c_ for c_ in scale_cols if AXES[c_['nm']].get('box')]:   # the name boxes beside the scales (bottom part only)
    occ.append(d_box(c['x0'] - AXBOX_W_YEARS - 0.1, c['x0'], AXBOX_Y0_ROWS / NGRID - 0.004, (AXBOX_Y0_ROWS + AXBOX_ROWS) / NGRID + 0.02))
reserved_wide = [(c['x0'] - (AXBOX_W_YEARS if AXES[c['nm']].get('box') else 0) - 0.15, c['x1'] + 0.15) for c in scale_cols]   # scale + its name box
plot_bb = base.get_window_extent(renderer)

# ---- CPI pillars (white bars), with a caption that has an arrow to the base year on the commodity curve
ppi_s = pd.Series(yoy.fillna(0).values, index=p_dates)      # 12-month % change
pillar_centres = []
px_per_year = base.transData.transform((num(ts(2001)), 0))[0] - base.transData.transform((num(ts(2000)), 0))[0]
plan = []
if pillars is not None:
    for _, r in pillars.iterrows():
        have = [(PILLAR_LABELS[i], float(r[f"p{i + 1}"])) for i in range(5)
                if pd.notna(pd.to_numeric(r[f"p{i + 1}"], errors="coerce"))]
        if not have:
            continue
        b_year, e_year = int(r.base), int(r.end)
        block = len(have) * PILLAR_W_YEARS + (len(have) - 1) * PILLAR_GAP_YEARS
        # centre the bars on the biggest inflation hump near the end of the episode: they stand ON the baseline, inside the area
        win = ((p_dates >= pd.Timestamp(f"{e_year - 5}-01-01")) & (p_dates <= pd.Timestamp(f"{e_year + 1}-12-31"))).values
        k = int(np.argmax(np.where(win & valid, yoy.fillna(-99).values, -99)))
        cx = yf(p_dates[k])
        _t = base.text(0, 0, f"THE {b_year} DOLLAR BOUGHT\nTHE ABOVE AMOUNTS IN {e_year}", fontsize=FS(5.6),
                       fontweight="bold", family=HEAD_FONT)
        _cw = _t.get_window_extent(renderer).width / px_per_year + 0.45
        _t.remove()
        cx = min(cx, END.year - block / 2 - PILLAR_EDGE_GAP)               # bars not too close to the right edge
        span = max(block, 2.4)                                   # the caption is wider than the bars
        lefts = [l for l in np.arange(START.year + 0.3, END.year - block - 0.2, 0.25)
                 if not any(l + block / 2 - span / 2 < r1 and l + block / 2 + span / 2 > r0 for r0, r1 in reserved_wide)
                 and l + block / 2 - span / 2 > START.year and l + block / 2 + span / 2 < END.year]
        left = min(lefts, key=lambda l: abs(l + block / 2 - cx)) if lefts else cx - block / 2
        _t = base.text(0, 0, f"THE {b_year} DOLLAR BOUGHT\nTHE ABOVE AMOUNTS IN {e_year}", fontsize=FS(5.6),
                       fontweight="bold", family=HEAD_FONT)
        cap_w = _t.get_window_extent(renderer).width / px_per_year + 0.45      # caption width in years, with its round ends
        _t.remove()
        plan.append(dict(have=have, b_year=b_year, e_year=e_year, block=block, left=left,
                         half=max(block, cap_w) / 2, target=b_year + 0.5, cap_w=cap_w))
    # an arrow must not land on another block: first slide that block left (up to PILLAR_MAX_SHIFT), then the arrow right
    for pk in plan:
        for pj in plan:
            if pj is pk:
                continue
            cj = pj["left"] + pj["block"] / 2
            hi = cj + pj["half"] + PILLAR_ARROW_GAP
            if cj - pj["half"] - PILLAR_ARROW_GAP < pk["target"] < hi:
                move = min(PILLAR_MAX_SHIFT, hi - pk["target"])
                pj["left"] -= move
                hi -= move
                if pk["target"] < hi:
                    pk["target"] = hi
    for pl in plan:
        have, b_year, e_year, block, left = pl["have"], pl["b_year"], pl["e_year"], pl["block"], pl["left"]
        pillar_centres.append(left + block / 2)
        y0 = base_h
        tallest = 0
        for i, (lab, cents) in enumerate(have):
            xa = ts(left + i * (PILLAR_W_YEARS + PILLAR_GAP_YEARS))
            xb = ts(left + i * (PILLAR_W_YEARS + PILLAR_GAP_YEARS) + PILLAR_W_YEARS)
            hgt = cents / 100 * PILLAR_100C_IN / plot_h_in
            tallest = max(tallest, hgt)
            base.add_patch(Rectangle((num(xa), y0), num(xb) - num(xa), hgt, facecolor="white", edgecolor="k",
                                     lw=0.7, zorder=30))
            base.text(xa + (xb - xa) / 2, y0 + hgt / 2, lab, rotation=90, fontsize=FS(5.6 if len(lab) < 9 else 4.6),
                      fontweight="bold", family=HEAD_FONT, ha="center", va="center", zorder=31)
            base.text(xa + (xb - xa) / 2, y0 + hgt + 0.009, f"{cents:.0f}c", fontsize=FS(5.6), fontweight="bold",
                      bbox=dict(boxstyle="round,pad=0.12,rounding_size=0.3", fc="white", ec="none"),
                      family=HEAD_FONT, ha="center", va="bottom", zorder=31)
        # caption: rounded "pill" box just under the baseline, two lines; the arrow ends ON the baseline at the base year
        capt = f"THE {b_year} DOLLAR BOUGHT\nTHE ABOVE AMOUNTS IN {e_year}"
        cap_fs, cap_pad = FS(5.6), 0.40
        _tmp = base.text(0, 0, capt, fontsize=cap_fs, fontweight="bold", family=HEAD_FONT, linespacing=1.05)
        _te = _tmp.get_window_extent(renderer)
        _tmp.remove()
        fs_px = cap_fs * fig.dpi / 72
        box_h_px = _te.height + 2 * cap_pad * fs_px
        yc_disp = base.transData.transform((0, base_h))[1] - 4 - box_h_px / 2
        cap_xy = (num(ts(min(left + block / 2, END.year - pl['cap_w'] / 2 - CAPTION_EDGE_GAP))), base.transData.inverted().transform((0, yc_disp))[1])
        cap = base.annotate(capt, xy=(num(ts(pl["target"])), base_h), xytext=cap_xy, ha="center", va="center",
                            fontsize=cap_fs, fontweight="bold", family=HEAD_FONT, zorder=32, linespacing=1.05,
                            bbox=dict(boxstyle=f"round,pad={cap_pad},rounding_size={box_h_px / fs_px / 2:.2f}",
                                      fc="white", ec="k", lw=0.7),
                            arrowprops=dict(arrowstyle="simple,head_length=0.8,head_width=1.0,tail_width=0.28", fc="white",
                                            ec="k", lw=0.8, mutation_scale=8, connectionstyle="arc3,rad=-0.15",
                                            shrinkA=0, shrinkB=1))
        occ.append(d_box(left - 0.2, left + block + 0.2, base_h - 0.003, y0 + tallest + 0.05, pad=2))
        _cb = tbox(cap)
        _p = cap_pad * fs_px + box_h_px / 2 * 0.3
        occ.append(Bbox.from_extents(_cb.x0 - _p - 4, _cb.y0 - cap_pad * fs_px - 4, _cb.x1 + _p + 4, _cb.y1 + cap_pad * fs_px + 4))
        cx0 = cap_xy[0]
        cap_y0, cap_h = cap_xy[1], 0.0
        # the arrow path (a quadratic curve, like matplotlib's arc3 with rad=-0.15) is an obstacle for the labels placed later
        P0 = base.transData.transform((cx0, cap_y0 + cap_h / 2)); P2 = base.transData.transform((num(ts(pl["target"])), base_h))
        Cc = np.array([(P0[0] + P2[0]) / 2 + (-0.15) * (P2[1] - P0[1]), (P0[1] + P2[1]) / 2 - (-0.15) * (P2[0] - P0[0])])
        for tt in np.linspace(0, 1, 60):
            q = (1 - tt) ** 2 * P0 + 2 * tt * (1 - tt) * Cc + tt ** 2 * P2
            occ.append(Bbox.from_extents(q[0] - 9, q[1] - 9, q[0] + 9, q[1] + 9))

# ---- label placer
def anchor_for(series, yr):
    """(year, plot height) of a point on the named line"""
    if series in ("debt", "income"):
        arr = debt if series == "debt" else income
        i = int(np.clip(round(yr - 0.5) - years.iloc[0], 0, n_a - 1))
        return years.iloc[i] + 0.5, float(pos("money", [arr.iloc[i]])[0]), float(arr.iloc[i])
    ser, ax_ = {"bus": (pd.Series(bus.values, index=b_dates), "business"),
                "yield": (pd.Series(yld.values, index=m_dates), "yield"),
                "dow": (pd.Series(dow.values, index=m_dates), "dow"),
                "ppi": (ppi_s, "commodity")}[series]
    j = ser.index.get_indexer([ts(yr)], method="nearest")[0]
    return yf(ser.index[j]), float(pos(ax_, [ser.iloc[j]])[0]), float(ser.iloc[j])

def candidates(prefer=0, dxmax=8.0, dymax=0.55):
    out = []
    for dx in np.arange(-dxmax, dxmax + 0.01, 0.5):
        for dy in np.arange(0.035, dymax, 0.025):
            for sgn in (1, -1):
                pen = 0 if (prefer == 0 or sgn == prefer) else 0.35
                out.append((abs(dx) * 0.035 + dy + pen, dx, sgn * dy))
    out.sort()
    return out

def inside_plot(bb):
    return bb.x0 > plot_bb.x0 + 3 and bb.x1 < plot_bb.x1 - 3 and bb.y0 > plot_bb.y0 + 3 and bb.y1 < plot_bb.y1 - 3

# ---- the lines and the black area as points in display space, so labels can avoid covering them
def dense(dates, hts, step=5):
    xy = base.transData.transform(np.column_stack([mdates.date2num(pd.to_datetime(dates)), np.asarray(hts, float)]))
    xy = xy[np.isfinite(xy[:, 1])]
    xs = np.arange(xy[0, 0], xy[-1, 0], step)
    return np.column_stack([xs, np.interp(xs, xy[:, 0], xy[:, 1])])

line_xy = np.vstack([dense(annual_dates, pos("money", debt)), dense(annual_dates, pos("money", income)),
                     dense(m_dates, pos("dow", dow)), dense(m_dates, pos("yield", yld)),
                     dense(p_dates[valid], comm_h[valid])])
_bx = dense(b_dates, bus_h)
_z = base.transData.transform((0, zero_h))[1]
area_xy = np.vstack([np.column_stack([_bx[:, 0], _z + f * (_bx[:, 1] - _z)]) for f in np.linspace(0, 1, 9)])
both_xy = np.vstack([line_xy, area_xy])
_cx = dense(p_dates[valid], comm_h[valid])
_cb = base.transData.transform((0, base_h))[1]
comm_area_xy = np.vstack([np.column_stack([_cx[:, 0], _cb + f * (_cx[:, 1] - _cb)]) for f in np.linspace(0, 1, 9)])
event_xy = np.vstack([both_xy, comm_area_xy])     # event labels also stay off the green / red commodity area
key_xy = np.vstack([dense(annual_dates, pos("money", debt)), dense(annual_dates, pos("money", income))])

def hits(bb, pts, m=2):
    return np.any((pts[:, 0] >= bb.x0 - m) & (pts[:, 0] <= bb.x1 + m) & (pts[:, 1] >= bb.y0 - m) & (pts[:, 1] <= bb.y1 + m))

def line_backgrounds(t, txt, fs, padf):
    """yellow background behind each line, as wide as that line (instead of one rectangle for the whole label)"""
    bb = tbox(t)
    lines = txt.split("\n")
    n = len(lines)
    lh = bb.height / n
    pad = padf * fs * fig.dpi / 72
    cx = (bb.x0 + bb.x1) / 2
    inv = base.transData.inverted()
    for i, ln in enumerate(lines):
        tt = base.text(0, 0, ln, fontsize=fs, fontweight="bold", family=HEAD_FONT)
        w = tt.get_window_extent(renderer).width
        tt.remove()
        y_top = bb.y1 - i * lh + (pad if i == 0 else 0)
        y_bot = bb.y1 - (i + 1) * lh - (pad if i == n - 1 else 0)
        (dx0, dy0), (dx1, dy1) = inv.transform((cx - w / 2 - pad, y_bot)), inv.transform((cx + w / 2 + pad, y_top))
        base.add_patch(Rectangle((dx0, dy0), dx1 - dx0, dy1 - dy0, facecolor=YELLOW, edgecolor="none", alpha=0.92,
                                 zorder=12.9))

_cache = {}
def place(txt, xa, ha_, prefer, kind, obst=None, dxmax=8.0, arrow_check=True, min_h=None, dx_sign=0, min_drop=0.0):
    """try positions around the anchor (xa year, ha_ height); keep the first free one"""
    if kind == "event":                      # business-volume event names: yellow, no border
        kw = dict(fontsize=FS(6.4), bbox=dict(boxstyle="square,pad=0.14", fc="none", ec="none"))
    elif kind == "moment":                   # values of debt / income: yellow with a black border
        kw = dict(fontsize=FS(5.4), bbox=dict(boxstyle="square,pad=0.28", fc=YELLOW, ec="k", lw=0.7))
    else:                                    # names, inflation labels: white
        kw = dict(fontsize=FS(5.4), bbox=dict(boxstyle="square,pad=0.28", fc="white", ec="k", lw=0.5))
    a_disp = base.transData.transform((num(ts(xa)), ha_))
    t = base.annotate(txt, xy=(num(ts(xa)), ha_), xytext=(num(ts(xa)), ha_), ha="center", va="center", fontweight="bold",
                      family=HEAD_FONT, zorder=13, linespacing=1.0, annotation_clip=False,
                      arrowprops=dict(arrowstyle="-|>,head_length=0.3,head_width=0.12", lw=0.4, color="k",
                                      shrinkA=0, shrinkB=0.5), **kw)
    key = (prefer, dxmax)
    if key not in _cache:
        _cache[key] = candidates(prefer, dxmax)
    for _, dx, dy in _cache[key]:
        if dx_sign and dx * dx_sign < -0.01:
            continue
        if min_drop and dy > -min_drop:
            continue
        t.xyann = (num(ts(xa + dx)), min(max(ha_ + dy, 0.03), 0.97))
        bb = tbox(t)
        if not inside_plot(bb):
            continue
        if min_h is not None and bb.y0 < base.transData.transform((0, min_h))[1] + 3:
            continue
        if kind == "event" and bb.x0 - 4 < a_disp[0] < bb.x1 + 4 and bb.y0 - 13 < a_disp[1] < bb.y1 + 13:
            continue                                     # leave room for the pointer between the label and its point
        if obst is not None and hits(bb, obst, 7 if kind == "moment" else 2):
            continue
        pad = Bbox.from_extents(bb.x0 - 3, bb.y0 - 3, bb.x1 + 3, bb.y1 + 3)
        if any(pad.overlaps(o) for o in occ):
            continue
        c = np.array([(bb.x0 + bb.x1) / 2, (bb.y0 + bb.y1) / 2])
        pts = [c + (a_disp - c) * s_ for s_ in np.linspace(0, 1, 60)]
        ok = True
        for q in (pts if arrow_check else []):
            if bb.x0 <= q[0] <= bb.x1 and bb.y0 <= q[1] <= bb.y1:
                continue
            if any(o.x0 <= q[0] <= o.x1 and o.y0 <= q[1] <= o.y1 for o in occ):
                ok = False
                break
        if not ok:
            continue
        occ.append(pad)
        if kind == "event":
            line_backgrounds(t, txt, FS(6.4), 0.14)
            gap_px = (bb.y0 - a_disp[1]) if a_disp[1] < bb.y0 else (a_disp[1] - bb.y1)
            if bb.x0 + 2 < a_disp[0] < bb.x1 - 2 and gap_px < 40:   # point right above/below: small triangle, no arrow
                t.arrow_patch.set_visible(False)
                below = a_disp[1] < bb.y0
                tri = 4.2 * fig.dpi / 72                     # triangle size in pixels; its tip just touches the curve
                ty = (a_disp[1] + tri * 0.55) if below else (a_disp[1] - tri * 0.55)
                tx_, ty_ = base.transData.inverted().transform((a_disp[0], ty))
                base.plot([tx_], [ty_], marker="v" if below else "^", ms=4.2, color="k", zorder=13, clip_on=False)
        for q in pts:                                                # the arrow also becomes an obstacle
            occ.append(Bbox.from_extents(q[0] - 4, q[1] - 4, q[0] + 4, q[1] + 4))
        return True
    t.remove()
    return False

def place_tiers(txt, xa, ha_, prefer, kind, tiers, dxmax=8.0, last_resort=False, min_h=None, dx_sign=0, min_drop=0.0):
    if min_drop and any(place(txt, xa, ha_, -1, kind, obst=o, dxmax=dxmax, min_drop=min_drop) for o in tiers):
        return True
    for sgn in ([dx_sign, 0] if dx_sign else [0]):            # the asked-for direction first, then anywhere
        if min_h is not None and any(place(txt, xa, ha_, prefer, kind, obst=o, dxmax=dxmax, min_h=min_h, dx_sign=sgn) for o in tiers):
            return True
        if any(place(txt, xa, ha_, prefer, kind, obst=o, dxmax=dxmax, dx_sign=sgn) for o in tiers):
            return True
    # last resort: the arrow may cross another label, but the label itself still never overlaps one
    return last_resort and place(txt, xa, ha_, prefer, kind, obst=None, dxmax=dxmax, arrow_check=False)

skipped = []
# ---- war names inside the shaded band, wherever they fit (the shading itself is always drawn)
band_y1 = base.transData.transform((0, base_h))[1]
band_y0 = plot_bb.y0
drawn_w = []
WAR_SHORT = {"WAR IN UKRAINE": "UKRAINE WAR", "WAR IN AFGHANISTAN": "AFGHAN WAR", "WORLD WAR II": "WWII"}
for wname_full, w0, w1, frac in sorted(WARS, key=lambda w: w[2] - w[1]):
    # full name first, then a shorter one, then a smaller font, until it fits
    for wname, wfs in [(wname_full, WAR_FS), (WAR_SHORT.get(wname_full, wname_full), WAR_FS),
                       (WAR_SHORT.get(wname_full, wname_full), WAR_FS - 1.5)]:
        t = base.text(0, 0, wname, fontsize=FS(wfs), fontweight="bold", family=HEAD_FONT)
        bb = t.get_window_extent(renderer)
        tw, th = bb.width, bb.height
        t.remove()
        xc0 = base.transData.transform((num(ts(w0)), 0))[0]
        xc1 = base.transData.transform((num(ts(w1)), 0))[0]
        yc = (band_y0 + band_y1) / 2                          # always centred in height; only moved left/right
        best = None
        for relaxed in (False, True):
            for xc in sorted(np.arange(xc0 - tw * (0.6 if xc1 - xc0 < tw else 0), xc1 + tw * (0.6 if xc1 - xc0 < tw else 0) + 1, 6),
                             key=lambda v: abs(v - (xc0 + xc1) / 2)):
                if xc - tw / 2 < xc0 - tw * 0.7 or xc + tw / 2 > xc1 + tw * 0.7:
                    continue
                box = Bbox.from_extents(xc - tw / 2 - 2, yc - th / 2 - 2, xc + tw / 2 + 2, yc + th / 2 + 2)
                if not inside_plot(box) or any(box.overlaps(o) for o in occ):
                    continue
                if not relaxed and hits(box, line_xy):
                    continue
                best = (xc, box)
                break
            if best:
                break
        if best is None:
            continue
        xd = base.transData.inverted().transform((best[0], yc))
        base.text(xd[0], xd[1], wname, ha="center", va="center", fontsize=FS(wfs), fontweight="bold", family=HEAD_FONT, zorder=14)
        occ.append(best[1])
        drawn_w.append(wname)
        break
    else:
        skipped.append(f"{wname_full} (label; shading drawn)")
print("war names drawn:", ", ".join(drawn_w) or "none")

# ---- 1. events (from your sheet), arrows to the peak / trough of the business-activity curve.
#         They may not cover any line or the black area; if no such spot exists the label is skipped.
bus_s = pd.Series(bus.values, index=b_dates)
used = []
for e in sorted(events, key=lambda e: e["start"]):
    seg = bus_s[e["start"]:e["end"]]
    boom = any(w in e["name"].lower() for w in UP_WORDS)
    cand = seg[[all(abs((d - u).days) > 300 for u in used) for d in seg.index]]
    if cand.empty:
        cand = seg
    free = cand[[not any(r0 - 0.3 < yf(d) < r1 + 0.5 for r0, r1 in reserved) for d in cand.index]]
    if not free.empty:                                         # keep arrows out of the scale columns if possible
        cand = free
    when = cand.idxmax() if boom else cand.idxmin()
    used.append(when)
    val = float(seg.loc[when])
    anchor_h = float(pos("business", [val])[0]) + (0.004 if boom else -0.004)
    txt = "\n".join(textwrap.wrap(e["name"].upper(), 15)) + f"\n({val:+.0f}%)"
    side = EVENT_SIDE.get(e["name"], +1 if boom else -1)
    if not place_tiers(txt, yf(when), anchor_h, side, "event", [event_xy, both_xy, line_xy, key_xy], dxmax=EVENT_DXMAX.get(e["name"], 16.0), last_resort=True,
                       min_h=zero_h if (side == +1 and not boom) else None, dx_sign=EVENT_DX.get(e["name"], 0)):
        skipped.append(e["name"])

# ---- 2. red inflation areas: white label BELOW the baseline, arrow up into the red part
def red_spans():
    spans, i, n = [], 0, len(red)
    while i < n:
        if red[i]:
            j = i
            while j + 1 < n and red[j + 1]:
                j += 1
            spans.append((j - i + 1, i, j))
            i = j + 1
        else:
            i += 1
    return sorted(spans, reverse=True)

comm_xy = dense(p_dates[valid], comm_h[valid], step=2)
base_y_disp = base.transData.transform((0, base_h))[1]

def place_inside_area(k, layouts):
    """put plain text inside the commodity area around index k; True if it fits entirely inside"""
    xa = yf(p_dates[k])
    for relaxed in (False, True):
        for txt in layouts:
            for fs_ in (6.6, 6.0, 5.4, 4.8):
                t = base.text(num(ts(xa)), base_h, txt, ha="center", va="center", fontsize=FS(fs_), fontweight="bold",
                              family=HEAD_FONT, zorder=12, linespacing=1.0)
                for dx in sorted(np.arange(-1.5, 1.51, 0.1), key=abs):
                    for fy in np.arange(0.5, 0.05, -0.05):
                        top = float(comm_h[k])
                        t.set_position((num(ts(xa + dx)), base_h + fy * (top - base_h)))
                        bb = t.get_window_extent(renderer)
                        sel = (comm_xy[:, 0] >= bb.x0 - 3) & (comm_xy[:, 0] <= bb.x1 + 3)
                        if not sel.any() or comm_xy[sel, 1].min() < bb.y1 + 3 or bb.y0 < base_y_disp + 3:
                            continue                          # must sit entirely between the baseline and the curve
                        pad_ = Bbox.from_extents(bb.x0 - 3, bb.y0 - 3, bb.x1 + 3, bb.y1 + 3)
                        if any(pad_.overlaps(o) for o in occ):
                            continue
                        others = line_xy[~np.isin(np.arange(len(line_xy)), [])]
                        if not relaxed and hits(pad_, key_xy):
                            continue
                        occ.append(pad_)
                        return True
                t.remove()
    return False

n_red = 0
for _, i, j in red_spans():
    if n_red >= RED_LABEL_COUNT:
        break
    k = i + int(np.argmax(comm_h[i:j + 1]))
    xa = yf(p_dates[k])
    if any(abs(xa - c) < 2.5 for c in pillar_centres) or any(r0 < xa < r1 for r0, r1 in reserved):
        continue                                   # that hump already has pillars (or a scale) on it
    pk = f"(PEAK {float(ppi_yoy.iloc[i:j + 1].max()):+.0f}%)"   # unsmoothed peak, as quoted in the text
    if place_inside_area(k, [f"PRICE\nINFLATION\nAREA\n{pk}", f"PRICE INFLATION\nAREA {pk}",
                             f"PRICE\nINFLATION\nPEAK\n{float(yoy.iloc[k]):+.0f}%"]):
        n_red += 1
    else:
        skipped.append(f"price inflation area {int(xa)}")

# ---- 4. moments on the lines (values read from the workbook): yellow boxes with a black border
def money_txt(v):
    return f"{v:,.0f} BILLION" if v < 1000 else f"{v / 1000:.1f} TRILLION"

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
for mom in MOMENTS:
    series, yr = mom[0], mom[1]
    title = mom[2] if len(mom) > 2 else None
    if series in ("debt", "income"):
        xa, ha_, v = anchor_for(series, yr + 0.5)
        nm = "FED. DEBT" if series == "debt" else "NATIONAL INCOME"
        txt = f"{nm} {yr}\n{money_txt(v)}"
    else:
        xa, ha_, v = anchor_for(series, yr)
        when = f"{MONTHS[min(int(round((xa % 1) * 12)), 11)]} {int(xa)}"
        txt = f"DOW JONES {when}\n{v:,.0f}" if series == "dow" else f"BOND YIELD {when}\n{v:.1%}"
        if title:
            txt = f"{title}\n{txt}"
    if not place_tiers(txt, xa, ha_, 0, "moment", [both_xy, line_xy, None], min_drop=MOMENT_DROP.get((series, yr), 0.0)):
        skipped.append(txt.replace("\n", " "))
if skipped:
    print("No room for these labels (add space or remove them in the settings):", "; ".join(skipped))

# =====================================================================
# 6. LEGEND, TEXT STRIP UNDER THE CHART (title, QR code, text blocks)
# =====================================================================
# legend as in the original: six names, each written above a short sample of its line, between chart and outline
LEGEND = [("Federal\nDebt", dict(color=DEBT_C, lw=1.6)), ("National\nIncome", dict(color=INC_C, lw=1.6)),
          ("Business\nActivity", "area"), ("Bonds", dict(color="k", ls=":", lw=1.4)),
          ("Stocks", dict(color="k", ls=(0, (6, 2, 1, 2)), lw=1.1)), ("Commodity\nPrices", dict(color="k", ls=(0, (4, 2)), lw=1.1))]
lx0 = (OUT_R + 0.08) / FIG_W
lx1 = (OUT_R + 0.08 + LEGEND_W_IN) / FIG_W
ly_top, ly_bot = y_frame_bottom + h_plot * 0.93, y_frame_bottom + h_plot * 0.12
for i, (name, st) in enumerate(LEGEND):
    yy = ly_top - i * (ly_top - ly_bot) / (len(LEGEND) - 1)
    fig.text((lx0 + lx1) / 2, yy, name, fontsize=FS(LEGEND_FS), family=HEAD_FONT, fontweight="bold", va="bottom", ha="center",
             linespacing=1.0)
    ys = yy - 0.05 / FIG_H
    if st == "area":
        fig.add_artist(Rectangle((lx0, ys - 0.04 / FIG_H), lx1 - lx0, 0.08 / FIG_H, transform=fig.transFigure,
                                 facecolor="k", edgecolor="none"))
    else:
        fig.add_artist(Line2D([lx0, lx1], [ys, ys], transform=fig.transFigure, **st))

def draw_lines(lines, x, y_top):
    y = y_top
    for l in lines:
        y -= l["h"]
        base_y = y + l["h"] * 0.22
        if l["kind"] == "text":
            ws = l["words"]
            sp_nat = text_w("a a", ws[0][1], l["size"]) - text_w("aa", ws[0][1], l["size"])
            total = sum(w for _, _, w in ws)
            room = l["width"] - l["indent"]
            gap = (room - total) / (len(ws) - 1) if (l["justify"] and len(ws) > 1) else sp_nat
            xx = x + l["indent"]
            for w, f, ww in ws:
                fig.text(xx / FIG_W, base_y / FIG_H, w, fontproperties=fprop(f, l["size"]), va="baseline", ha="left")
                xx += ww + gap
        elif l["kind"] == "row2":
            for j in range(2):
                x_ = x + j * (l["half"] + TABLE_GAP_IN)
                vx = x_ + l["half"]
                if l["head"]:
                    fig.text(vx / FIG_W, base_y / FIG_H, l["head"], fontproperties=fprop(F_BOLD, l["size"]),
                             va="baseline", ha="right")
                    continue
                c = l["cells"][j]
                if not c:
                    continue
                p_ = fprop(c["font"], l["size"])
                fig.text(x_ / FIG_W, base_y / FIG_H, c["label"], fontproperties=p_, va="baseline")
                fig.text(vx / FIG_W, base_y / FIG_H, c["val"], fontproperties=p_, va="baseline", ha="right")
                lw_ = text_w(c["label"] + " ", c["font"], l["size"])
                vw_ = text_w(" " + c["val"], c["font"], l["size"])
                dw = text_w(" .", F_BODY, l["size"])
                n = max(0, int((vx - x_ - lw_ - vw_) / dw))
                fig.text((vx - vw_) / FIG_W, base_y / FIG_H, " ." * n, fontproperties=fprop(F_BODY, l["size"]),
                         va="baseline", ha="right")
    return y

def heading(txt, x, y_top):
    sz = TEXT_PT * HEAD_SCALE
    fig.text(x / FIG_W, (y_top - sz * 1.0 / 72) / FIG_H, txt, fontproperties=fprop(F_HEAD, sz), va="baseline")
    return y_top - head_h(TEXT_PT)

strip_top = TEXT_BELOW_IN - TEXT_GAP_TOP_IN
_sp_y = TEXT_BELOW_IN - OUTLINE_GAP_IN - 0.15     # small print: one line per paragraph between the outline and the text
for _k, _para in enumerate(sp_paras):
    fig.text(((OUT_L + OUT_R) / 2) / FIG_W, _sp_y / FIG_H, _para, ha="center", va="baseline",
             fontproperties=fprop(F_BOLD if _k == len(sp_paras) - 1 and len(sp_paras) > 1 else F_ITAL, TEXT_PT * 0.85))
    _sp_y -= TEXT_PT * 0.85 * 1.45 / 72
strip_bot = strip_top - head_h(TEXT_PT) - TEXT_H
xx = OUT_L
y0 = heading(LEFT_BLOCK, xx, strip_top)
for col in flow(layout(blocks[LEFT_BLOCK], COL_W_LEFT, TEXT_PT), LEFT_COLS, TEXT_H):
    short = TEXT_H - height(col)                  # spread a small shortfall over the lines so all columns end level
    if 0 < short < 3 * TEXT_PT * TEXT_LEADING / 72:
        k_ = 1 + short / height(col)
        col = [dict(l, h=l["h"] * k_) for l in col]
    draw_lines(col, xx, y0)
    xx += COL_W_LEFT + COL_GAP_IN
centre_x0 = xx - COL_GAP_IN + BLOCK_GAP_IN
centre_x1 = centre_x0 + CENTRE_W
xx = centre_x1 + BLOCK_GAP_IN
for h, w_ in zip(RIGHT, COL_WS):
    draw_lines(layout(blocks[h], w_, TEXT_PT), xx, heading(h, xx, strip_top))
    xx += w_ + COL_GAP_IN

# ---- centre: red stripes, title, years, QR code, small print -- the whole group centred in the strip
cw = centre_x1 - centre_x0
cx = (centre_x0 + centre_x1) / 2
t_size, t_w, sp_lines, stack = T_SIZE, T_W, SP_LINES, STACK_H
yr_size = t_size * 0.42
strip_bot = strip_top - STRIP_H
y_s1 = (strip_top + strip_bot) / 2 + stack / 2
y_title = y_s1 - 0.10 - t_size * 0.72 / 72
y_s2 = y_title - 0.11
stripe_x0, stripe_x1 = cx - t_w / 2 - 0.02, cx + t_w / 2 + 0.02
fig.text(cx / FIG_W, y_title / FIG_H, TITLE, fontproperties=fprop(F_TITLE, t_size), ha="center", va="baseline")
yr_w = text_w(YEARS_TEXT, F_SINCE, yr_size)
stripe_lw = STRIPE_IN * 72
fig.add_artist(Line2D([stripe_x0 / FIG_W, stripe_x1 / FIG_W], [y_s1 / FIG_H] * 2, color=STRIPE_RED, lw=stripe_lw,
                      solid_capstyle="butt", transform=fig.transFigure))
for a_, b_ in ((stripe_x0, cx - yr_w / 2 - 0.08), (cx + yr_w / 2 + 0.08, stripe_x1)):
    fig.add_artist(Line2D([a_ / FIG_W, b_ / FIG_W], [y_s2 / FIG_H] * 2, color=STRIPE_RED, lw=stripe_lw,
                          solid_capstyle="butt", transform=fig.transFigure))
fig.text(cx / FIG_W, y_s2 / FIG_H, YEARS_TEXT, fontproperties=fprop(F_SINCE, yr_size), ha="center", va="center_baseline")
y_next = y_s2 - 0.20
if QR_URL:
    try:
        import qrcode
    except ImportError:
        sys.exit("QR code: run  pip install qrcode  (or set QR_URL = None)")
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, border=0)
    qr.add_data(QR_URL)
    qr.make(fit=True)
    mat = np.array(qr.get_matrix(), dtype=bool)
    n = mat.shape[0]
    qax = fig.add_axes([(cx - QR_SIZE_IN / 2) / FIG_W, (y_next - QR_SIZE_IN) / FIG_H, QR_SIZE_IN / FIG_W, QR_SIZE_IN / FIG_H])
    qax.set_xlim(0, n); qax.set_ylim(n, 0); qax.axis("off")
    for (r, c) in zip(*np.nonzero(mat)):
        qax.add_patch(Rectangle((c, r), 1, 1, facecolor="k", edgecolor="k", lw=0.15))
    y_next -= QR_SIZE_IN + 0.08
    for i, ln in enumerate(QR_LINES):
        sz = TEXT_PT * (1.15 if i == 0 else 1.0)
        y_next -= sz * 1.4 / 72
        fig.text(cx / FIG_W, y_next / FIG_H, ln, fontproperties=fprop(F_HEAD if i == 0 else F_BODY, sz),
                 ha="center", va="baseline")
y_next -= 0.10
for l in sp_lines:
    y_next -= l["h"]
    fig.text(cx / FIG_W, (y_next + l["h"] * 0.22) / FIG_H, " ".join(w for w, _, _ in l["words"]),
             fontproperties=fprop(F_ITAL, l["size"]), ha="center", va="baseline")

fig.savefig(f"{OUT}.png", dpi=PNG_DPI)
fig.savefig(f"{OUT}.pdf")                         # vector: sharp at any print size
print(f"page {FIG_W:g} x {FIG_H:.1f} in, frame {frame_w_in:.1f} x {frame_h_in:.2f} in (ratio {frame_w_in / frame_h_in:.1f}:1)")
print(f"saved {OUT}.png and {OUT}.pdf")
