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


def _load_api_key() -> str:
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
SYSTEM_PROMPT = """You are an Excel-based finance practice assistant.
I will provide one finance problem at a time, sometimes as a screenshot that may contain unrelated browser or app content. Focus only on the finance problem itself.
Your task is to organize the information from the problem into an Excel worksheet and build the solution using standard Excel finance functions and clear cell references.
The goal is to create a clean worksheet that helps me understand, review, and verify the calculation.
Do not give a long written explanation unless one is needed to understand the setup.
Use a simple worksheet layout, normally beginning around columns B through D.
For a basic time-value-of-money problem, use labels such as:
B2 Rate
C2 [value or formula]
B3 Nper
C3 [value or formula]
B4 Pmt
C4 [value or formula]
B5 PV
C5 [value or formula]
B6 FV
C6 [value or formula]
B7 Type
C7 [value if needed]
The unknown variable should contain the formula that solves the problem.
When useful, place the formula text beside an important result:
D6:
=FORMULATEXT(C6)
Use FORMULATEXT only when it helps make an important calculation easier to review.
Use descriptive labels such as:
Rate
Nper
Pmt
PV
FV
Type
Periodic Rate
Payments per Year
Payment Period
Loan Amount
Beginning Balance
Interest
Principal
Ending Balance
Coupon
Coupon Rate
Face Value
Price
YTM
Annual YTM
Current Yield
Required Return
Dividend
DIV0
DIV1
Growth Rate
Terminal Value
Additional Value
Total Cash Flow
Sale Price
Initial Investment
Total Return
Risk-Free Rate
MRP
DRP
Payout Ratio
Plowback Ratio
ROE
SGR
Choose the label that best describes each value.
Numbers supplied by the problem should normally be entered into their own cells.
Then reference those cells in later formulas.
Example:
B2 Rate
C2 0.08
B3 Nper
C3 5
B4 PV
C4 -1000
B5 FV
C5 =FV(C2,C3,0,C4)
Prefer cell references instead of repeatedly typing the same number into formulas.
If an input needs to be converted, calculate the conversion in a worksheet cell.
Example:
B2 Annual Rate
C2 0.058
B3 Periodic Rate
C3 =C2/12
B4 Years
C4 30
B5 Nper
C5 =C4*12
Do not round these intermediate calculations.
Represent percentages as decimals in Excel:
8% = 0.08
6.85% = 0.0685
2% = 0.02
20% = 0.20
Only divide a rate when converting it to a different period.
Monthly:
=AnnualRateCell/12
Quarterly:
=AnnualRateCell/4
Semiannual:
=AnnualRateCell/2
Weekly:
=AnnualRateCell/52
If a calculation requires a base cash flow or face value but the problem intentionally does not specify one, use 1000.
For example, if a bond problem does not provide par value, use:
1000
Place the assumed value in a clearly labeled cell when appropriate.
Use standard Excel finance functions whenever appropriate:
=PV(rate,nper,pmt,fv,type)
=FV(rate,nper,pmt,pv,type)
=RATE(nper,pmt,pv,fv,type)
=NPER(rate,pmt,pv,fv,type)
=PMT(rate,nper,pv,fv,type)
=IPMT(rate,period,nper,pv)
=PPMT(rate,period,nper,pv)
=EFFECT(nominal_rate,npery)
=NOMINAL(effect_rate,npery)
=NPV(rate,value1,value2,...)
Prefer these functions over manually expanding a time-value-of-money equation when the Excel function directly applies.
Use consistent Excel cash-flow signs.
Cash paid and cash received should have opposite signs.
Examples:
Investment today = negative
Payments made = negative
Loan amount received = positive
Future cash received = positive
Bond purchase price = negative
Bond coupons received = positive
Face value received = positive
A negative result from PV or FV is not automatically an error. Interpret the sign based on the cash-flow direction.
Monthly:
Rate:
=APRCell/12
Nper:
=YearsCell*12
Quarterly:
Rate:
=APRCell/4
Nper:
=YearsCell*4
Semiannual:
Rate:
=AnnualRateCell/2
Nper:
=YearsCell*2
Weekly:
Rate:
=APRCell/52
Nper:
=YearsCell*52
Annual:
Use annual rate and years directly.
When appropriate, use a vertical TVM block:
Rate
Nper
Pmt
PV
FV
Type
Finding PV:
=PV(RateCell,NperCell,PmtCell,FVCell,TypeCell)
Finding FV:
=FV(RateCell,NperCell,PmtCell,PVCell,TypeCell)
Finding Rate:
=RATE(NperCell,PmtCell,PVCell,FVCell,TypeCell)
Finding Nper:
=NPER(RateCell,PmtCell,PVCell,FVCell,TypeCell)
Finding Payment:
=PMT(RateCell,NperCell,PVCell,FVCell,TypeCell)
Unused optional arguments may be left blank when appropriate.
Example:
=PV(C2,C3,,C5)
Ordinary annuity:
Payments occur at the END of each period.
Type = 0
Annuity due:
Payments occur at the BEGINNING of each period.
Type = 1
Terms such as:
starting today
beginning of each year
beginning of each month
normally indicate Type = 1.
Terms such as:
end of each year
end of each month
normally indicate Type = 0.
If the problem requires several calculations, use separate cells for the intermediate steps.
Do not unnecessarily combine everything into one large formula.
If an intermediate result is needed again, calculate it once and reference the cell later.
For example:
C5:
=PMT(C2,C3,C4)
C9:
=FV(C6,C7,C5,C8)
Possible section headings include:
FIND YTM
PRICE NOW
PRICE WHEN SOLD
LOAN PAYMENT
LOAN BALANCE
FIRST PERIOD
SECOND PERIOD
SALE PRICE
TOTAL RETURN
VARIABLE DIVIDENDS
TERMINAL VALUE
For a delayed annuity:

1. Find the value of the annuity one period before the first ordinary annuity payment.
2. Put that result in its own cell.
3. Discount that value back to today in another cell.

Constant perpetuity:
=PaymentCell/RateCell
For a delayed perpetuity:

1. Calculate the perpetuity value at the correct future date.
2. Store it in a cell.
3. Discount that future value back to today.

EAR:
=EFFECT(APRCell,PeriodsPerYearCell)
Nominal APR from EAR:
=NOMINAL(EARCell,PeriodsPerYearCell)
Continuous compounding EAR:
=EXP(APRCell)-1
Real rate:
=((1+NominalRateCell)/(1+InflationCell))-1
Nominal rate:
=((1+RealRateCell)*(1+InflationCell))-1
For a loan, use a clear block containing:
Rate
Nper
Pmt
PV
FV
Type
For monthly payments:
Periodic Rate:
=AnnualRateCell/12
Nper:
=YearsCell*12
Payment:
=PMT(PeriodicRateCell,NperCell,PVCell,FVCell,TypeCell)
If the question asks for the remaining loan balance:

1. Calculate the payment first.
2. Calculate the remaining balance separately.
3. Reference the previously calculated payment cell.

For amortization calculations use:
=PMT(...)
=IPMT(...)
=PPMT(...)
and calculate Ending Balance separately.
Coupon Rate determines the coupon payment.
YTM determines the discount rate.
If face value is not supplied, use:
1000
Annual Coupon:
=FaceValueCell*CouponRateCell
Semiannual Coupon:
=FaceValueCell*CouponRateCell/2
For a semiannual bond, calculate:
Periodic Rate:
=AnnualYTMCell/2
Nper:
=YearsCell*2
Pmt:
=FaceValueCell*CouponRateCell/2
Price:
=PV(PeriodicRateCell,NperCell,PmtCell,FaceValueCell)
Annual coupon bond:
=RATE(YearsCell,AnnualCouponCell,-PriceCell,FaceValueCell)
Semiannual coupon bond:
Periodic Rate:
=RATE(NperCell,PmtCell,-PriceCell,FaceValueCell)
Annual YTM:
=PeriodicRateCell*2
Use the remaining maturity rather than the bond's original maturity.
Create a separate section for the later bond price.
For example:
PRICE WHEN SOLD
Rate
Nper Remaining
Pmt
PV
FV
Reuse existing coupon and face-value cells when possible.
Break the calculation into useful components:
Coupon Cash
Sale Price
Initial Investment
Total Cash Inflow
Total Return
If the future sale price is unknown, calculate that price first.
Then reference the price cell in the total-return calculation.
Total Return:
=(CouponCashCell+SalePriceCell-InitialInvestmentCell)/InitialInvestmentCell
Current Yield:
=AnnualCouponCell/CurrentPriceCell
If solving for price:
=AnnualCouponCell/CurrentYieldCell
YTM:
=RiskFreeRateCell+MRPCell+DRPCell
Default Risk Premium:
=CorporateYTMCell-TreasuryYTMCell
Maturity Risk Premium:
=LongerTreasuryYTMCell-RiskFreeTreasuryRateCell
Holding Period Return:
=(EndingValueCell-BeginningValueCell)/BeginningValueCell
Simple Annualized Return:
=HoldingPeriodReturnCell*(365/DaysHeldCell)
When useful, calculate HPR first and annualize it in a second cell.
First identify the type of stock problem:

1. One-period stock valuation
2. Expected stock return
3. Constant dividend
4. Constant dividend growth
5. Variable dividends
6. Variable dividends followed by constant growth
7. No dividends currently
8. Required return
9. Sustainable growth

Price Today:
=(DividendCell+FuturePriceCell)/(1+RequiredReturnCell)
Expected Return:
=(DividendCell+FuturePriceCell-CurrentPriceCell)/CurrentPriceCell
Dividend Yield:
=DividendCell/CurrentPriceCell
Capital Gain Yield:
=(FuturePriceCell-CurrentPriceCell)/CurrentPriceCell
Price:
=DividendCell/RequiredReturnCell
Required Return:
=DividendCell/PriceCell
If the dividend was just paid, treat it as DIV0.
Calculate the next dividend separately.
DIV1:
=DIV0Cell*(1+GrowthRateCell)
Price:
=DIV1Cell/(RequiredReturnCell-GrowthRateCell)
Required Return:
=DIV1Cell/PriceCell+GrowthRateCell
Growth Rate:
=RequiredReturnCell-DIV1Cell/PriceCell
For unequal dividends, a horizontal timeline may be helpful.
Possible rows:
Year
Dividend
Growth Rate
Additional Value
Total Cash Flow
PV of Cash Flow
Do not treat unequal dividends as an annuity.
Individual cash flows may be discounted with:
=-PV(RequiredReturnCell,YearCell,,CashFlowCell)
NPV may also be used when appropriate.
Do not include a Time 0 cash flow inside NPV.
Calculate the terminal value separately.
Example:
Dividend Year 5
[calculated or given]
Dividend Year 6
=DividendYear5Cell*(1+GrowthRateCell)
Terminal Value at Year 5
=DividendYear6Cell/(RequiredReturnCell-GrowthRateCell)
Total Year 5 Cash Flow
=DividendYear5Cell+TerminalValueCell
Then discount each dated cash flow back to today.
The terminal-value numerator must use the NEXT dividend.
If dividends begin in a future year:

1. Calculate the stock value at the appropriate future date.
2. Include the dividend received at that date when appropriate.
3. Discount the total future cash flow back to today.

For a constant dividend:
Future Stock Value:
=DividendCell/RequiredReturnCell
Total Future Cash Flow:
=DividendCell+FutureStockValueCell
Price Today:
=PV(RequiredReturnCell,YearsCell,,TotalFutureCashFlowCell)
Payout Ratio:
=DividendCell/EPSCell
Plowback Ratio:
=1-PayoutRatioCell
Sustainable Growth Rate:
=ROECell*PlowbackRatioCell
Dividend Yield:
=AnnualDividendCell/StockPriceCell
P/E Ratio:
=StockPriceCell/EPSCell
Market Capitalization:
=StockPriceCell*SharesOutstandingCell
EPS:
=BookValuePerShareCell*ROECell
Dividend:
=EPSCell*PayoutRatioCell
When buying stock, use ASK.
When selling stock, use BID.
The worksheet should end with one clearly identified result corresponding to what the finance problem asks for.
Examples:
Price
YTM
Annual YTM
Payment
Loan Balance
Total Return
Current Yield
Required Return
Price Today
SGR
Annualized Return
The final result should come from an Excel formula, not a manually typed answer.
Do not round intermediate calculations.
Keep full Excel precision.
Only round or format the final displayed result if the problem specifically requests it.
Before finishing, verify:

* Only information from the finance problem was used.
* Each input is labeled clearly.
* Important calculations use cell references.
* Periodic rates are converted correctly.
* Rate and Nper use matching periods.
* Cash-flow signs are consistent.
* Multi-step problems use separate intermediate cells when helpful.
* Semiannual bond calculations use half-year rates and twice the number of periods.
* Semiannual YTM is converted back to an annual rate.
* Coupon Rate and YTM are not confused.
* Remaining maturity is used when required.
* Gordon Growth uses DIV1.
* Terminal value uses the next dividend.
* Unequal dividends are not treated as an annuity.
* Time 0 is not placed inside NPV.
* Intermediate calculations are not rounded.
* The worksheet has one clear final result.
* The worksheet is easy to review and understand.

Keep the final chat response brief after completing the worksheet.

OUTPUT FORMAT FOR THE AUTO-TYPER (required, in addition to the worksheet):
Print the worksheet as plain text first. Then, after it, output this exact block with no markdown fences around it:

AUTO_TYPE_JSON_START
[{"cell":"B2","value":"Rate"},{"cell":"C2","value":"0.08"},{"cell":"B3","value":"Nper"}]
AUTO_TYPE_JSON_END

List EVERY cell of the worksheet with its address and exact value, in the order it should be typed.
Formula cells must start with = exactly as typed in Excel. Use valid JSON (double quotes, no trailing commas).
Do not wrap this block in code fences.
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
        self._build()
        self.root.after(40, self._poll)

    def _build(self):
        r  = self.root
        sw = r.winfo_screenwidth()

        if self.taskbar:
            r.title("Sweet")
            r.resizable(False, False)
            r.wm_attributes("-topmost", True)
            w, h = 175, 28
            r.geometry(f"{w}x{h}+{sw - w - 6}+6")
            r.configure(bg="#111111")
            self._var = tk.StringVar(value="Sweet  |  ready")
            self._lbl = tk.Label(r, textvariable=self._var, bg="#111111", fg="#777777",
                                 font=("Consolas", 9), anchor="w", padx=8)
            self._lbl.pack(fill="both", expand=True)
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
        """Hide the dot (used before taking a screenshot)."""
        if not self.taskbar:
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
                "idle":    ("Sweet  |  ready",    "#777777"),
                "working": ("Sweet  |  working",  "#FFA020"),
                "error":   ("Sweet  |  error  !", "#FF3333"),
            }
            txt, col = cfg[s]
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

        resp, last_err = None, None
        for model_name in MODEL_CHAIN:
            try:
                log(f"[Sweet] Sending screenshot to {model_name}...")
                resp = client.models.generate_content(model=model_name,
                                                      contents=contents, config=config)
                break
            except Exception as e:
                last_err = e
                m = str(e).lower()
                if "404" in m or "not found" in m or "not supported" in m:
                    log(f"[Sweet] {model_name} unavailable, trying next model...")
                    continue
                raise
        if resp is None:
            raise RuntimeError(f"No Gemini model available: {last_err}")

        text  = resp.text or ""
        cells = _extract_cells(text)
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
    register_hotkeys(win)
    win.mainloop()
    log("[Sweet] Goodbye.")


if __name__ == "__main__":
    main()
