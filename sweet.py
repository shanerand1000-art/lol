#!/usr/bin/env python3
"""
Sweet -- Finance Screenshot -> Excel worksheet
===============================================
Install:
  pip install google-genai Pillow keyboard pyautogui pyperclip

API key (either one):
  * put it in a file named sweet_key.txt next to this script, OR
  * set the GEMINI_API_KEY environment variable

Run:
  python sweet.py             background mode: tiny dot only, no taskbar entry
  python sweet.py --taskbar   same, plus a small status bar in the taskbar
  (double-click sweet_background.pyw / sweet_taskbar.pyw to run with no console window)

Hotkeys (system-wide):
  Ctrl+Shift+S  screenshot -> Gemini -> worksheet copied to clipboard
                dot in top-right = working; dot gone = done and copied; red ! = error
  Ctrl+Shift+T  auto-type the worksheet into Excel, one character every 0.5 s
  Ctrl+Shift+X  stop the auto-typer immediately
  Ctrl+Shift+Q  quit Sweet

Clipboard format: tab-separated grid laid out by cell address. Click cell A1 in Excel
and press Ctrl+V to drop the whole worksheet (formulas included) in the right cells.
"""

import sys
import os
import re
import json
import time
import queue
import argparse
import threading
import tkinter as tk

# ─── dependency check ──────────────────────────────────────────────────────────
_need = []
try:
    from google import genai
    from google.genai import types as genai_types
except ImportError:
    _need.append("google-genai")
try:
    from PIL import ImageGrab
except ImportError:
    _need.append("Pillow")
try:
    import keyboard
except ImportError:
    _need.append("keyboard")
try:
    import pyautogui
    pyautogui.FAILSAFE = True      # slam the mouse into a screen corner to abort typing
    pyautogui.PAUSE = 0.0
except ImportError:
    _need.append("pyautogui")
try:
    import pyperclip
except ImportError:
    _need.append("pyperclip")

if _need:
    print("Missing packages. Install with:")
    print(f"  pip install {' '.join(_need)}")
    sys.exit(1)

# Correct dot placement / screenshot size on scaled Windows displays
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ─── configuration ─────────────────────────────────────────────────────────────
HERE = os.path.dirname(os.path.abspath(__file__))


# Paste your Gemini API key between the quotes, replacing the placeholder text.
# (Or skip this and use a sweet_key.txt file instead. Never upload this file to GitHub with your real key in it.)
API_KEY_HERE = "PASTE_YOUR_GEMINI_API_KEY_HERE"


def _load_api_key() -> str:
    if API_KEY_HERE.strip() and API_KEY_HERE != "PASTE_YOUR_GEMINI_API_KEY_HERE":
        return API_KEY_HERE.strip()
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if key:
        return key
    try:
        with open(os.path.join(HERE, "sweet_key.txt"), encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return ""


API_KEY = _load_api_key()

# Newest first. Override the first choice with:  set SWEET_MODEL=gemini-3.8-flash
# If a model is unavailable for your key, Sweet automatically tries the next one.
MODEL_CHAIN = [
    m for m in [
        os.environ.get("SWEET_MODEL", "").strip(),
        "gemini-3.8-flash",
        "gemini-3.7-flash",
        "gemini-3.5-flash",
        "gemini-3-flash-preview",
    ] if m
]

HOTKEY_ANALYZE = "ctrl+shift+s"
HOTKEY_TYPE    = "ctrl+shift+t"
HOTKEY_STOP    = "ctrl+shift+x"
HOTKEY_QUIT    = "ctrl+shift+q"

TYPE_START_DELAY = 3      # seconds to click into Excel after pressing the type hotkey
TYPE_CHAR_DELAY  = 0.5    # seconds between typed characters
CELL_PAUSE       = 0.5    # seconds between finishing one cell and moving to the next
NAV_KEY_DELAY    = 0.03   # per-character delay when typing a cell address (navigation only)
LOG_FILE         = os.path.join(HERE, "sweet.log")

# ─── system prompt (the full finance worksheet prompt) ─────────────────────────
SYSTEM_PROMPT = """You are an Excel finance worksheet assistant.
You will get ONE finance problem at a time, usually as a screenshot with extra browser or app stuff around it.
Only use the finance problem.
Ignore everything else.
Your job is to solve it directly in Excel cells.
Do not give a long written explanation.
Keep the worksheet short, natural, and to the point.
Keep it compact.
Do not make it look like a report, template, or polished model.
Do not over-label things.
Do not use long phrases if a short finance abbreviation works.
Do not add extra checks unless they are actually useful.
Do not add duplicate methods just to prove the answer.
Use a worksheet that feels natural:

* short labels
* simple layout
* only needed rows
* a little loose, not overly perfect
* still clear enough to follow

Use short labels like these:
r
g
n
PV
FV
PMT
YTM
APR
EAR
HPR
CY
Div
DIV0
DIV1
CF
P0
P1
P2
P5
TV
ROE
SGR
EPS
If a slightly longer label helps, keep it short:
Par
Coupon
Price
Ret
Beg Bal
End Bal
Int
Prin
Payout
Plowback
Sale Px
Do NOT use long labels like:
Present Value
Future Value
Number of Periods
Required Return for Investors
Total Annual Cash Flow Present Values
Prefer:
PV
FV
n
r
CF
PVs
Usually work in columns B through E.
Do not force everything into the same exact format.
Use what fits the problem.
For simple TVM / bond / loan problems, a small vertical setup is fine:
B3 r
C3 [value or formula]
B4 n
C4 [value or formula]
B5 PMT
C5 [value or formula]
B6 PV
C6 [value or formula]
B7 FV
C7 [value or formula]
B8 Type
C8 [only if needed]
For stock problems, usually use a small timeline across columns.
Example style:

```
    C      D      E      F
```

5 Year 1 2 3
6 Div 1.32 1.47 =E6*(1+$C$3)
7 P2 =F6/($C$2-$C$3)
8 CF =D6 =E6+E7
9 PV =D8/(1+$C$2)^D5 =E8/(1+$C$2)^E5
11 P0 =SUM(D9)
That kind of setup is preferred for stock questions.
Do NOT automatically add:
VARIABLES
VALUES
Using NPV function
Price = sum of PVs
Check
Verification
Answer check
unless they are actually useful.
Usually one method is enough.
If a SUM row is needed, just label it:
P0
Price
PV
or
Ans
Do not add a second method unless the problem naturally calls for it.
If something needs to be converted, show it in the cell where it belongs.
Examples:
r
=0.068/12
n
=5*12
PMT
=50/2
YTM
=SemiRate*2
Do not create extra input rows for every raw number unless they are reused a lot.
Short helper formulas are better.
Use these when needed:
=PV(rate,nper,pmt,fv,type)
=FV(rate,nper,pmt,pv,type)
=RATE(nper,pmt,pv,fv,type)
=NPER(rate,pmt,pv,fv,type)
=PMT(rate,nper,pv,fv,type)
=IPMT(rate,period,nper,pv)
=PPMT(rate,period,nper,pv)
=NPV(rate,value1,value2,...)
=EFFECT(nominal_rate,npery)
=NOMINAL(effect_rate,npery)
=EXP(rate)-1
=SUM(...)
For bond date problems, also allowed:
=DATE(year,month,day)
=PRICE(settlement,maturity,rate,yld,redemption,frequency)
=YIELD(settlement,maturity,rate,pr,redemption,frequency)
Do not use other fancy methods.
Keep rate and periods matched.
Monthly:
r = APR/12
n = Years*12
Quarterly:
r = APR/4
n = Years*4
Semiannual:
r = AnnualRate/2
n = Years*2
Weekly:
r = APR/52
n = Years*52
Use short labels:
r
n
PMT
PV
FV
Unused optional arguments can be left blank.
Examples:
=PV(0.08,2,,3000)
=RATE(10,,-2000,4000)
Use normal Excel sign rules.
Money out and money in should have opposite signs.
Examples:
investment today = negative
payments made = negative
loan received = positive
FV received = positive
bond price paid = negative
coupon received = positive
par received = positive
Do not force every answer positive.
If a later step needs the positive version of a negative PV cell, use:
=-that_cell
Use:
Type = 0 for ordinary annuity
Type = 1 for annuity due
Use:
PV = PMT/r
for perpetuities.
For delayed annuities or perpetuities, do it in steps.
Do not cram everything into one giant formula.
Use a compact setup:
r
n
PMT
PV
FV
For amortization rows, use short headers:
Yr
Beg Bal
Pmt
Int
Prin
End Bal
Use:
=PMT(...)
=IPMT(...)
=PPMT(...)
Use short setup:
r
n
PMT
PV
FV
For semiannual bonds:
r
=AnnualYTM/2
n
=Years*2
PMT
=AnnualCoupon/2
FV
1000
PV
=PV(r_cell,n_cell,PMT_cell,FV_cell)
If solving YTM:
semi rate
=RATE(n_cell,PMT_cell,PV_cell,FV_cell)
YTM
=semi_rate_cell*2
Do not forget:
coupon rate gives PMT
YTM gives r
If par is not given, use 1000.
For current yield:
CY
=AnnualCoupon/Price
For total return, keep it short:
Coupon CF
Sale Px
Init Px
TR
=(CouponCF+SalePx-InitPx)/InitPx
Keep it short:
Beg Px
End Px
HPR
=(EndPx-BegPx)/BegPx
If annualized:
AF
=365/Days
Ann Ret
=HPR*AF
For stock problems, usually use a timeline.
Use short row labels such as:
Year
Div
g
r
P2
P5
TV
CF
PVs
P0
P0
=(Div+P1)/(1+r)
Ret
=(Div+P1-P0)/P0
DY
=Div/P0
CG
=(P1-P0)/P0
P0
=DIV1/r
or if solving for return:
r
=DIV1/P0
If dividend just paid is DIV0:
DIV1
=DIV0*(1+g)
P0
=DIV1/(r-g)
If solving return:
r
=(DIV1/P0)+g
Use a small timeline.
Keep labels short.
Example:
Year 0 1 2 3 4 5
Div 2.20
r 6.9%
P5
=2.20/0.069
P0
=(P5+2.20)/(1+0.069)^5
Do not over-explain it.
Use a timeline.
Typical rows:
Year
Div
g
TV
CF
PVs
P0
If there is terminal growth:

1. calculate the next dividend
2. calculate terminal value at the right year
3. add it to the cash flow in that year
4. discount the cash flows back

Example style:
r 15.1%
g 2.6%
Year 1 2 3
Div 1.32 1.47 =D6*(1+$C$3)
P2 =E6/($C$2-$C$3)
CF =C6 =D6+D7
PV =C8/(1+$C$2)^C5 =D8/(1+$C$2)^D5
P0 =SUM(C9)
That style is preferred.
Do not automatically add an extra “Using NPV” row unless it is genuinely helpful.
Use a horizontal timeline.
Rows can be:
Year
g
Div
TV
CF
PVs
P0
Each dividend should grow from the prior one.
Terminal value should use the NEXT dividend.
Do not use a big explanation block.
Keep it simple:
Payout
=Div/EPS
Plowback
=1-Payout
SGR
=ROE*Plowback
Keep formatting normal and basic.
Okay:

* bold short labels if needed
* % format for rates
* currency for prices/dividends
* standard gridlines
* normal alignment
* one or two blank rows between sections

Do not:

* add decorative colors
* add lots of borders
* merge cells everywhere
* center giant titles
* create big polished headers
* add long written notes

It should feel like a quick clean working sheet.
Do not round intermediate calculations.
Keep full Excel precision.
Only round the final displayed answer if the question asks for it.
Before finishing, silently check:

* Only the finance problem was used
* Labels are short
* Worksheet is compact
* No unnecessary verification rows were added
* Rate and n match
* Signs are correct
* Semiannual bond setup is correct
* YTM is annualized when needed
* Gordon growth uses DIV1
* Terminal value uses the next dividend
* Time 0 is not included in NPV
* Intermediate values are not rounded
* There is one clear final answer

After filling the Excel cells, if a chat reply is needed, reply only:
Done.

OUTPUT FORMAT FOR THE AUTO-TYPER (this replaces "fill the Excel cells" and "reply only Done" above):
You cannot edit Excel directly here. Instead, print the finished worksheet as short plain text first. Then output this exact block, with no markdown fences around it:

AUTO_TYPE_JSON_START
[{"cell":"B3","value":"r"},{"cell":"C3","value":"=0.08/2"},{"cell":"B4","value":"n"},{"cell":"C4","value":"=5*2"}]
AUTO_TYPE_JSON_END

List EVERY cell that should contain something, with its address and exact contents, in the order it should be typed.
Formula cells start with = exactly as typed in Excel. Put only cell contents in the JSON, no formatting. A plain rate may be written as a percent like 15.1% (Excel turns it into 0.151); otherwise no dollar signs or thousands commas.
Keep column A and row 1 completely empty: no labels, no stray 0, no year or timeline numbers there. Start the worksheet in column B at row 3 or lower. For a timeline, put the year numbers in a row below row 1 starting at column C, with the row labels in column B.
Follow the short label style above (r, n, PV, FV, PMT, CF, P0, Sale Px). Never use FORMULATEXT.
Use valid JSON (double quotes, no trailing commas) and do not wrap the block in code fences.
"""

# ─── shared state ──────────────────────────────────────────────────────────────
_lock   = threading.Lock()
_busy   = False                 # an analysis or auto-type run is in progress
_cells  = []                    # [(addr, value), ...] from the last analysis
_stop   = threading.Event()     # set by HOTKEY_STOP to abort auto-typing


def log(msg: str = ""):
    print(msg)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except OSError:
        pass


# ─── UI window ─────────────────────────────────────────────────────────────────
class SweetWindow:
    """
    Background mode (default): a tiny dot in the top-right corner, no taskbar entry.
      hidden = idle / done,  orange pulsing = working,  red "!" = error
    Taskbar mode (--taskbar): a slim status bar that also shows in the taskbar.

    All Tk calls happen on the main thread; other threads post work with call().
    """

    DOT_SIZE = 14
    PAD      = 2

    def __init__(self, taskbar: bool):
        self.taskbar = taskbar
        self.root    = tk.Tk()
        self._state  = "idle"
        self._phase  = 0.0
        self._dir    = 1.0
        self._q      = queue.Queue()
        self.actions = {}          # filled in by main(): snap / type / stop
        self._build()
        self.root.after(40, self._poll)

    def _build(self):
        r  = self.root
        sw = r.winfo_screenwidth()

        if self.taskbar:
            r.title("Sweet")
            r.resizable(False, False)
            r.wm_attributes("-topmost", True)
            r.configure(bg="#111111")
            self._var = tk.StringVar(value="Sweet | ready")
            self._lbl = tk.Label(r, textvariable=self._var, bg="#111111", fg="#777777",
                                 font=("Consolas", 9), anchor="w", padx=8, width=15)
            self._lbl.pack(side="left", fill="y")
            for key, text in (("snap", "Snap"), ("type", "Type"), ("stop", "Stop")):
                tk.Button(r, text=text, width=6, relief="flat", bg="#2d2d2d", fg="white",
                          activebackground="#444444", activeforeground="white",
                          command=lambda k=key: self.actions[k]()
                          ).pack(side="left", padx=2, pady=3)
            r.update_idletasks()
            w, h = r.winfo_reqwidth() + 6, 30
            r.geometry(f"{w}x{h}+{sw - w - 6}+6")
        else:
            r.overrideredirect(True)            # no title bar, not in the taskbar
            r.wm_attributes("-topmost", True)
            sz = self.DOT_SIZE + self.PAD * 2
            r.geometry(f"{sz}x{sz}+{sw - sz - 3}+3")
            r.configure(bg="#000000")
            cv = tk.Canvas(r, width=sz, height=sz, bg="#000000", highlightthickness=0)
            cv.pack()
            self._cv  = cv
            self._dot = cv.create_oval(self.PAD, self.PAD,
                                       self.PAD + self.DOT_SIZE, self.PAD + self.DOT_SIZE,
                                       fill="#2a2a2a", outline="")
            self._txt = cv.create_text(sz // 2, sz // 2, text="", fill="white",
                                       font=("Arial", 8, "bold"))
            r.withdraw()

    # ── thread-safe plumbing ────────────────────────────────────────────────────
    def call(self, fn):
        self._q.put(fn)

    def _poll(self):
        try:
            while True:
                self._q.get_nowait()()
        except queue.Empty:
            pass
        self.root.after(40, self._poll)

    def hide(self):
        """Hide the dot / status bar so it is not in the screenshot."""
        self.call(self.root.withdraw)

    def set_state(self, state: str):
        """state in {'idle', 'working', 'error'}; safe to call from any thread."""
        self._state = state
        self.call(self._apply)

    # ── rendering (main thread only) ────────────────────────────────────────────
    def _apply(self):
        s = self._state
        if self.taskbar:
            cfg = {
                "idle":    ("Sweet | ready",   "#777777"),
                "working": ("Sweet | working", "#FFA020"),
                "error":   ("Sweet | error !", "#FF3333"),
            }
            txt, col = cfg[s]
            self.root.deiconify()
            self._var.set(txt)
            self._lbl.configure(fg=col)
            return
        if s == "idle":
            self.root.withdraw()
            return
        self.root.deiconify()
        if s == "working":
            self._cv.itemconfig(self._dot, fill="#FFA500")
            self._cv.itemconfig(self._txt, text="")
            self._phase, self._dir = 0.0, 1.0
            self._tick()
        else:
            self._cv.itemconfig(self._dot, fill="#FF3333")
            self._cv.itemconfig(self._txt, text="!")

    def _tick(self):
        if self._state != "working":
            return
        self._phase += self._dir * 0.09
        if   self._phase >= 1.0: self._phase, self._dir = 1.0, -1.0
        elif self._phase <= 0.0: self._phase, self._dir = 0.0,  1.0
        g = int(80 + self._phase * 120)
        self._cv.itemconfig(self._dot, fill=f"#ff{g:02x}00")
        self.root.after(55, self._tick)

    def mainloop(self):
        self.root.mainloop()


# ─── cell helpers ──────────────────────────────────────────────────────────────
_ADDR = re.compile(r"^([A-Z]{1,3})(\d{1,7})$")


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _addr_key(addr: str):
    m = _ADDR.match(addr)
    return (int(m.group(2)), _col_to_num(m.group(1)))


def _clean_value(v) -> str:
    s = str(v).strip()
    for a, b in (("−", "-"), ("–", "-"), ("—", "-"),
                 ("‘", "'"), ("’", "'"), ("“", '"'), ("”", '"')):
        s = s.replace(a, b)
    return s.replace("\t", " ").replace("\r", " ").replace("\n", " ")


def _extract_cells(text: str) -> list:
    """Parse the AUTO_TYPE_JSON block; returns [(addr, value)] sorted top-to-bottom, left-to-right."""
    m = re.search(r"AUTO_TYPE_JSON_START\s*(\[[\s\S]*?\])\s*AUTO_TYPE_JSON_END",
                  text, re.IGNORECASE)
    if not m:
        log("[Sweet] AUTO_TYPE_JSON block not found in response.")
        return []
    try:
        data = json.loads(m.group(1))
    except Exception as e:
        log(f"[Sweet] JSON parse error: {e}")
        return []
    cells = {}
    for item in data:
        try:
            addr = str(item["cell"]).strip().upper().replace("$", "")
            val  = _clean_value(item["value"])
        except Exception:
            continue
        if not _ADDR.match(addr):
            log(f"[Sweet] Skipping invalid cell address: {addr!r}")
            continue
        if val != "":
            cells[addr] = val          # last entry for an address wins
    return sorted(cells.items(), key=lambda kv: _addr_key(kv[0]))


def _stray_cells(cells: list) -> list:
    """Addresses in column A or row 1 (the worksheet must start at B2 or lower)."""
    out = []
    for addr, _ in cells:
        row, col = _addr_key(addr)
        if row == 1 or col == 1:
            out.append(addr)
    return out


def _num_to_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _to_clipboard_grid(cells: list) -> str:
    """Tab-separated grid positioned by cell address (A1 is the top-left of the text)."""
    rows, cols = {}, 0
    maxrow = 0
    for addr, val in cells:
        row, col = _addr_key(addr)
        rows.setdefault(row, {})[col] = val
        cols, maxrow = max(cols, col), max(maxrow, row)
    lines = []
    for r in range(1, maxrow + 1):
        cur = rows.get(r, {})
        lines.append("\t".join(cur.get(c, "") for c in range(1, cols + 1)).rstrip("\t"))
    return "\n".join(lines)


# ─── Gemini screenshot analysis ────────────────────────────────────────────────
def _generate(client, config, contents) -> str:
    """Call the newest available Gemini model and return the reply text."""
    last_err = None
    for model_name in MODEL_CHAIN:
        try:
            log(f"[Sweet] Asking {model_name}...")
            resp = client.models.generate_content(model=model_name,
                                                  contents=contents, config=config)
            return resp.text or ""
        except Exception as e:
            last_err = e
            m = str(e).lower()
            if "404" in m or "not found" in m or "not supported" in m:
                log(f"[Sweet] {model_name} unavailable, trying next model...")
                continue
            raise
    raise RuntimeError(f"No Gemini model available: {last_err}")


def analyze(win: SweetWindow):
    global _busy, _cells

    with _lock:
        if _busy:
            log("[Sweet] Busy -- wait for the current run to finish.")
            return
        _busy = True

    success = False
    try:
        if not API_KEY:
            raise RuntimeError("No API key. Put it in sweet_key.txt next to sweet.py "
                               "or set GEMINI_API_KEY.")

        # Dot hidden while the picture is taken, then it appears to show work has started.
        win.hide()
        time.sleep(0.3)
        screenshot = ImageGrab.grab()
        win.set_state("working")

        client = genai.Client(api_key=API_KEY)
        config = genai_types.GenerateContentConfig(system_instruction=SYSTEM_PROMPT,
                                                   temperature=0.0)
        contents = ["Analyze the finance problem shown in this screenshot and produce "
                    "the Excel worksheet:", screenshot]

        text  = _generate(client, config, contents)
        cells = _extract_cells(text)

        stray = _stray_cells(cells)
        if stray:
            log(f"[Sweet] Gemini used column A / row 1 ({', '.join(stray)}); asking it to redo...")
            fix = contents + [
                "Your previous answer was:\n" + text,
                "That put cells in column A or row 1 (" + ", ".join(stray) + "). Redo the ENTIRE answer "
                "with nothing in column A or row 1: start in column B at row 3 or lower, and update "
                "every formula reference to match the new cell positions. Same output format as before.",
            ]
            text2  = _generate(client, config, fix)
            cells2 = _extract_cells(text2)
            if cells2 and len(_stray_cells(cells2)) < len(stray):
                text, cells = text2, cells2
            if _stray_cells(cells):
                log(f"[Sweet] WARNING: still cells in column A / row 1: {', '.join(_stray_cells(cells))}")
        clean = re.sub(r"AUTO_TYPE_JSON_START[\s\S]*?AUTO_TYPE_JSON_END", "",
                       text, flags=re.IGNORECASE).strip()

        log("\n" + "=" * 66)
        log(clean)
        log("=" * 66)

        if not cells:
            pyperclip.copy(clean)
            raise RuntimeError("Gemini returned no cell data for the auto-typer "
                               "(worksheet text was copied instead). Try again.")

        with _lock:
            _cells = cells
        pyperclip.copy(_to_clipboard_grid(cells))
        chars = sum(len(v) for _, v in cells)
        log(f"[Sweet] Copied {len(cells)} cells to clipboard.  "
            f"{HOTKEY_TYPE.upper()} auto-types them (~{chars * TYPE_CHAR_DELAY / 60:.1f} min).")
        success = True

    except Exception as exc:
        log(f"\n[Sweet] ERROR: {exc}")

    finally:
        win.set_state("idle" if success else "error")
        with _lock:
            _busy = False


# ─── auto-typer ────────────────────────────────────────────────────────────────
class _Stopped(Exception):
    pass


def _wait(seconds: float):
    """Sleep, but abort immediately if the stop hotkey was pressed."""
    if _stop.wait(seconds):
        raise _Stopped()


def _goto_cell(addr: str):
    """Jump to a cell with Excel's Go To dialog (Ctrl+G)."""
    pyautogui.hotkey("ctrl", "g")
    _wait(0.35)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("delete")
    pyautogui.typewrite(addr, interval=NAV_KEY_DELAY)
    pyautogui.press("enter")
    _wait(0.35)


def _type_value(value: str):
    """Type one character at a time, TYPE_CHAR_DELAY apart, then commit the cell."""
    for ch in value:
        if _stop.is_set():
            raise _Stopped()
        pyautogui.typewrite(ch)
        _wait(TYPE_CHAR_DELAY)
    # Excel's AutoComplete can tack extra text onto labels like "Price"; Delete removes it.
    pyautogui.press("delete")
    pyautogui.press("enter")


def auto_type(win: SweetWindow):
    global _busy

    with _lock:
        cells = list(_cells)
        if _busy:
            log("[Sweet] Busy -- wait for the current run to finish.")
            return
        if not cells:
            log(f"[Sweet] Nothing to type yet -- run {HOTKEY_ANALYZE.upper()} first.")
            win.set_state("error")
            return
        _busy = True

    _stop.clear()
    success = False
    try:
        chars = sum(len(v) for _, v in cells)
        log(f"[Sweet] Click your Excel sheet now -- typing {len(cells)} cells "
            f"(~{chars * TYPE_CHAR_DELAY / 60:.1f} min) starts in {TYPE_START_DELAY}s.  "
            f"{HOTKEY_STOP.upper()} stops it.")
        _wait(TYPE_START_DELAY)
        win.set_state("working")

        for addr, value in cells:
            _goto_cell(addr)
            _type_value(value)
            log(f"  {addr:<6} {value}")
            _wait(CELL_PAUSE)

        log("[Sweet] Auto-type complete.")
        success = True

    except _Stopped:
        log("[Sweet] Auto-type stopped.")
        success = True
    except Exception as exc:
        log(f"[Sweet] Auto-type error: {exc}")

    finally:
        win.set_state("idle" if success else "error")
        with _lock:
            _busy = False


# ─── hotkeys ───────────────────────────────────────────────────────────────────
def _spawn(fn, win):
    threading.Thread(target=fn, args=(win,), daemon=True).start()


def register_hotkeys(win: SweetWindow):
    keyboard.add_hotkey(HOTKEY_ANALYZE, lambda: _spawn(analyze, win), suppress=True)
    keyboard.add_hotkey(HOTKEY_TYPE,    lambda: _spawn(auto_type, win), suppress=True)
    keyboard.add_hotkey(HOTKEY_STOP,    _stop.set, suppress=True)
    keyboard.add_hotkey(HOTKEY_QUIT,
                        lambda: (_stop.set(), keyboard.unhook_all(), win.call(win.root.quit)),
                        suppress=True)
    log(f"[Sweet] Running.  {HOTKEY_ANALYZE.upper()} analyze | {HOTKEY_TYPE.upper()} auto-type | "
        f"{HOTKEY_STOP.upper()} stop | {HOTKEY_QUIT.upper()} quit")


# ─── entry point ───────────────────────────────────────────────────────────────
def main(argv=None):
    ap = argparse.ArgumentParser(description="Sweet -- Finance Screenshot -> Excel worksheet")
    ap.add_argument("--taskbar", action="store_true",
                    help="show a small status bar in the taskbar (default: tiny dot only)")
    args = ap.parse_args(argv)

    if not API_KEY:
        log("[Sweet] WARNING: no API key found (sweet_key.txt or GEMINI_API_KEY).")

    win = SweetWindow(taskbar=args.taskbar)
    win.actions = {"snap": lambda: _spawn(analyze, win),
                   "type": lambda: _spawn(auto_type, win),
                   "stop": _stop.set}
    register_hotkeys(win)
    win.mainloop()
    log("[Sweet] Goodbye.")


if __name__ == "__main__":
    main()
