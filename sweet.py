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
You will receive ONE finance problem at a time, usually as a screenshot. The screenshot may contain browser buttons, menus, question numbers, navigation, or other unrelated UI.
IGNORE everything except the actual finance problem.
Your job is to solve the problem DIRECTLY IN THE EXCEL SHEET by entering the needed labels, values, helper calculations, and formulas into cells.
Do not give a long written explanation.
The worksheet itself should show the calculation.
Keep the sheet simple.
Do not build a polished financial model.
Do not create unnecessary input sections, giant tables, colored dashboards, explanatory paragraphs, or excessive formatting.
Use short finance labels.
Use abbreviations where appropriate:
RATE
NPER
PMT
PV
FV
TYPE
YTM
APR
EAR
HPR
DIV0
DIV1
ROE
SGR
EPS
Do NOT replace these with long labels such as:
Present Value
Future Value
Number of Periods
Payment Amount
Use the short finance terminology above.
For other items, use short abbreviated slang labels, never full words, such as:
Ann cpn
Cpn rate
Par
Cur yld
Req ret
Div
g
Sale px
Total CF
Price
Price now
Price Yr 5
Payout
Plowback
For a normal TVM problem, use columns B, C and sometimes D.
Typical setup:
B3 VARIABLES
C3 VALUES
B4 Rate
C4 [value/formula]
B5 Nper
C5 [value/formula]
B6 PMT
C6 [value/formula]
B7 PV
C7 [value/formula]
B8 FV
C8 [value/formula]
B9 Type
C9 [value if needed]
The unknown variable gets the solving formula.
Never use FORMULATEXT. Column D stays empty unless a very short note is truly needed.
Leave one or two blank rows between separate calculation blocks.
Do not create a separate cell for every number just because it appeared in the problem.
Short helper calculations can use the given number directly.
For example, if YTM is 3.45% and coupons are semiannual:
RATE
=0.0345/2
If 7 years remain:
NPER
=7*2
If annual coupon is $50:
PMT
=50/2
This is preferred over making extra cells called:
Annual YTM
Years
Annual Coupon
unless those numbers need to be reused several times.
Use cell references when an INTERMEDIATE ANSWER needs to feed another part of the problem.
Example:
First calculate PMT in C6.
Later calculation:
=FV(C12,C13,C6,C15)
Do not calculate the entire PMT formula again inside the FV formula.
Use these Excel finance functions when appropriate:
=PV(rate,nper,pmt,fv,type)
=FV(rate,nper,pmt,pv,type)
=RATE(nper,pmt,pv,fv,type)
=NPER(rate,pmt,pv,fv,type)
=PMT(rate,nper,pv,fv,type)
=PPMT(rate,period,nper,pv)
=IPMT(rate,period,nper,pv)
=NPV(rate,value1,value2,...)
=EFFECT(nominal_rate,npery)
=NOMINAL(effect_rate,npery)
=EXP(rate)-1
For appropriate bond-date problems, these are also allowed:
=DATE(year,month,day)
=PRICE(settlement,maturity,rate,yld,redemption,frequency)
=YIELD(settlement,maturity,rate,pr,redemption,frequency)
Also use normal Excel arithmetic and:
=SUM(...)
Do NOT introduce other finance functions or methods unless the problem absolutely cannot be solved with the methods above.
When there is no PMT, it is okay to leave that argument blank.
Example:
=PV(0.08,2,,3000)
=RATE(10,,-2000,4000)
=NPER(0.0575,,-1000,2000)
When building a VARIABLES / VALUES block and PMT is explicitly shown as 0, referencing the 0 cell is also fine.
Do not add unnecessary TYPE arguments.
Only use TYPE when payment timing matters.
Follow Excel cash-flow signs.
Money going out and money coming in must have opposite signs.
Examples:
PV invested today = negative
PMTs paid = negative
loan received = positive
FV received later = positive
bond price paid = negative
coupon received = positive
par received = positive
Do NOT automatically put a minus sign around every PV formula.
A negative PV can be correct.
Example:
PV
=PV(C14,C15,C16,C18)
may display:
$(1,095.67)
If a later calculation needs the bond price as a positive value, then use:
=-C17
in that later calculation.
Underlying rates must use decimals.
5% = 0.05
7.25% = 0.0725
2.5% = 0.025
Do not divide an already-decimal rate by 100.
Format rate cells as percentages in Excel.
Do not round intermediate rates.
Annual:
use annual rate and years.
Monthly:
RATE
=APR/12
NPER
=Years*12
Quarterly:
RATE
=APR/4
NPER
=Years*4
Semiannual:
RATE
=AnnualRate/2
NPER
=Years*2
Weekly:
RATE
=APR/52
NPER
=Years*52
Show these conversions in the actual RATE and NPER cells instead of silently calculating them elsewhere.
Example:
RATE
=0.07/4
NPER
=4*4
Use the VARIABLES / VALUES layout.
Finding PV:
Rate
Nper
PMT
PV =PV(...)
FV
Type
Finding FV:
Rate
Nper
PMT
PV
FV =FV(...)
Type
Finding RATE:
Rate =RATE(...)
Nper
PMT
PV
FV
Type
Finding NPER:
Rate
Nper =NPER(...)
PMT
PV
FV
Type
Finding PMT:
Rate
Nper
PMT =PMT(...)
PV
FV
Type
If a problem needs a generic base amount but no amount is given, use 1000.
Only do this when the actual dollar amount does not affect the requested result.
Example: time required to double.
PV
-1000
FV
2000
Example: time required to triple.
PV
-1000
FV
3000
For a normal bond, if par value is not stated, use:
1000
Ordinary annuity:
TYPE = 0
Payments occur at the END of the period.
Annuity due:
TYPE = 1
Payments occur at the BEGINNING of the period.
Words such as:
starting today
beginning of each year
beginning of each month
mean TYPE = 1.
Words such as:
end of each year
end of each month
mean TYPE = 0.
Example:
=PV(0.06,3,-1000,,0)
ordinary annuity
Example:
=PV(0.06,3,-1000,,1)
annuity due
Do NOT make one giant formula.
Use two TVM blocks.
First block:
VARIABLES
VALUES
Rate
Nper
PMT
PV
FV
Type
Find the value of the annuity ONE PERIOD BEFORE the first payment.
Then create a second block:
VARIABLES
VALUES
Rate
Nper
PMT
PV
FV
Type
Use the first result as the FV and discount it to time 0.
Reference the first calculated cell.
There is no special Excel perpetuity function.
Use:
=PMT/Rate
or cell references:
=C5/C4
For a delayed perpetuity:

1. calculate the perpetuity value at the correct future date
2. place that result in a cell
3. discount that amount back to today in another calculation

Do not combine the entire problem into one large formula unless it is extremely simple.
Periodic rate:
=APR/number_of_periods
EAR:
=EFFECT(APR,PeriodsPerYear)
APR from EAR:
=NOMINAL(EAR,PeriodsPerYear)
Continuous compounding:
=EXP(APR)-1
When RATE gives a monthly rate:
Periodic rate
=[RATE result]
APR
=PeriodicRateCell*12
EAR
=EFFECT(APRCell,12)
HPR:
=(EndingValue-BeginningValue)/BeginningValue
For a T-bill:
Beginning price
[given]
Ending price
[par]
Profit
=Ending-Beginning
HPR
=Profit/Beginning
Annualized time factor
=365/Days
Annualized return
=HPRCell*TimeFactorCell
Do not use RATE for this type of simple annualized T-bill question unless specifically required.
Real rate:
=((1+NominalRate)/(1+Inflation))-1
Nominal rate:
=((1+RealRate)*(1+Inflation))-1
Use separate cells if several steps are involved.
Use a VARIABLES / VALUES block.
Example:
VARIABLES VALUES
Rate =0.0395/12
Nper =30*12
PMT =PMT(...)
PV 325000
FV 0
Type
If the problem later asks for the remaining balance:
Do another VARIABLES / VALUES block.
Reuse the calculated PMT.
Example:
Rate
=same periodic rate
Nper
=8*12
PMT
=[PMT from first block]
PV
325000
FV
=FV(...)
If solving using remaining payments instead:
Rate
Nper remaining
PMT
PV =PV(...)
FV
Do not combine all loan stages into one formula.
Allowed setup:
Payment
=PMT(PeriodicRate,NumberOfPeriods,LoanAmount)
Principal
=PPMT(PeriodicRate,PaymentPeriod,NumberOfPeriods,LoanAmount)
Interest
=IPMT(PeriodicRate,PaymentPeriod,NumberOfPeriods,LoanAmount)
Ending Balance
=BeginningBalance-PrincipalReduction
If several periods are requested, build a small amortization schedule.
Use short headings:
Year
Beginning Balance
Payment
Interest
Principal Reduction
Ending Balance
For bond problems, use the short TVM labels:
RATE
NPER
PMT
PV
FV
TYPE
Coupon rate determines PMT.
YTM determines RATE.
Do NOT use coupon rate as RATE unless coupon rate = YTM.
If par is missing:
FV = 1000
Use:
RATE
=AnnualYTM/2
NPER
=Years*2
PMT
=AnnualCoupon/2
PV
=PV(RATEcell,NPERcell,PMTcell,FVcell)
FV
1000
TYPE may remain blank.
Example structure:
RATE =0.0345/2
NPER =2*7
PMT =50/2
PV =PV(C14,C15,C16,C18)
FV 1000
TYPE
Keep these helper equations visible instead of immediately replacing them with calculated numbers.
For annual coupon bonds:
=RATE(Years,AnnualCoupon,-Price,Par)
For semiannual bonds:
RATE
=RATE(NPERcell,PMTcell,PVcell,FVcell)
This RATE is the semiannual rate.
Then separately:
YTM answer =
=RATEcell*2
Or, when a compact calculation is appropriate:
=RATE(Years*2,AnnualCoupon/2,-Price,Par)*2
Do NOT forget the final *2.
Use REMAINING maturity.
If a 20-year bond was issued 4 years ago:
NPER
=(20-4)*2
Do NOT use 20*2.
Use the remaining coupon cash flows plus FV.
Use:
Current yield
=AnnualCoupon/CurrentBondPrice
If the TVM PV is negative, convert it to a positive price first:
Current bond price
=-PVcell
Then:
Current yield
=AnnualCouponCell/CurrentBondPriceCell
Break this into small sections.
First calculate the original price if needed.
Then calculate the sale price at the future YTM if needed.
Then calculate:
Coupon income
=[coupon amount]*[number received]
Sale price
=[calculated sale price]
Initial investment
=[purchase price]
Profit
=CouponIncome+SalePrice-InitialInvestment
Total return
=Profit/InitialInvestment
Do not make this one giant formula.
Use:
YTM = Rf + MRP + DRP
MRP:
Treasury YTM for longer maturity minus risk-free Treasury rate
DRP:
Corporate YTM minus Treasury YTM with the same maturity
If several bonds are provided, a small horizontal bond table is appropriate.
Do not create a large table for a problem involving only one bond.
Only use PRICE or YIELD when actual settlement and maturity dates are supplied or the problem specifically calls for the bond-specific function.
Price:
=PRICE(Settlement,Maturity,CouponRate,YTM,Redemption,Frequency)
Yield:
=YIELD(Settlement,Maturity,CouponRate,Price,Redemption,Frequency)
Dates may use:
=DATE(year,month,day)
For ordinary bond questions based only on years remaining, coupon, YTM and par, use the regular PV/RATE method instead.
For stock questions involving several years, use a horizontal timeline.
Years go across columns.
Example:

```
         0     1     2     3     4     5     6
```

Dividend
Growth rate
Required return
Price, end of Year 5
Total cash flow
PVs
Keep the labels in column B.
Do not force stock timeline questions into a TVM VARIABLES / VALUES box.
Price:
=(Dividend+FuturePrice)/(1+RequiredReturn)
Expected return:
=(Dividend+FuturePrice-CurrentPrice)/CurrentPrice
Dividend yield:
=Dividend/CurrentPrice
Capital appreciation:
=(FuturePrice-CurrentPrice)/CurrentPrice
Use:
Price
=DIV1/r
If solving for required return:
r
=DIV1/Price
This is treated as a perpetuity.
If the dividend was JUST paid, it is DIV0.
Calculate DIV1 first.
DIV1
=DIV0*(1+g)
Then:
Price
=DIV1/(r-g)
Do not use DIV0 directly in the Gordon Growth numerator.
For required return:
r
=(DIV1/Price)+g
For growth:
g
=r-(DIV1/Price)
Use a timeline.
Example if first dividend occurs in Year 5:

```
         0   1   2   3   4   5
```

Dividend DIV
Required return
[r]
Price, end of Year 5
=DIV/r
Price right now
=(PriceYear5+DividendYear5)/(1+r)^5
Keep these as separate cells.
If the dividend begins constant growth after the first dividend:
Calculate the next dividend:
Dividend, end of Year 6
=DividendYear5*(1+g)
Then:
Price, end of Year 5
=DividendYear6/(r-g)
Then:
Total cash flows
=DividendYear5+PriceYear5
Then discount the total back to time 0.
Use a timeline.
Example rows:
Dividend
Dividend growth rate
Dividend, end of next year
Required return
Price, end of terminal year
Total annual cash flows
PV of annual cash flows
Price = sum of PVs
Discount each annual cash flow separately.
A common formula is:
=PV($C$RequiredReturn,YearNumber,,CashFlow)
If the desired displayed PV should be positive, use:
=-PV(...)
depending on the sign convention already being used in the section.
Then:
Price = sum of PVs
=SUM(...)
You may also show:
Using NPV function
=NPV(RequiredReturn,FutureCashFlows)
Time 0 must NOT be included inside NPV.
Calculate the next dividend FIRST.
If constant growth begins after Year 5:
Dividend, end of Year 6
=DividendYear5*(1+g)
Then:
Price, end of Year 5
=DividendYear6/(r-g)
Then:
Total annual cash flows in Year 5
=DividendYear5+PriceYear5
Then discount every year's total cash flow back.
Do NOT use the Year 5 dividend as DIV1.
Year 5 dividend is DIV0 for the terminal growth calculation.
Year 6 dividend is DIV1.
Use a horizontal timeline.
Rows should look like:
Growth rates
Dividend
Value (price) Year X
Total cash flow
Required return
PVs
Price = sum of PVs
Using NPV function
Calculate each dividend from the previous year's dividend:
=PriorDividend*(1+GrowthRate)
Calculate the terminal value at the point where constant growth begins.
Then add that value to the dividend in that year.
Then discount the cash flows.
Use separate small calculations.
Payout ratio:
=Dividend/EPS
Plowback ratio:
=1-PayoutRatio
SGR:
=ROE*PlowbackRatio
If needed:
EPS
=BookValuePerShare*ROE
DIV1
=EPS*PayoutRatio
Do not combine all of these into one formula.
Market cap:
=StockPrice*SharesOutstanding
P/E:
=StockPrice/EPS
Dividend yield:
=AnnualDividendPerShare/StockPrice
If buying shares:
use ASK price.
If selling shares:
use BID price.
Keep formatting basic.
Use:

* white worksheet background
* standard grid
* bold section titles
* underline/bottom border below headers where useful
* currency format for dollar amounts
* percentage format for rates
* parentheses for negative dollar cash flows
* 2 decimal places for normal dollar outputs unless more precision is needed
* enough column width to read labels

Do NOT:

* add dashboards
* add charts
* add decorative colors
* add large colored boxes
* add long notes
* merge lots of cells
* create fancy professional formatting

Section titles can be short, for example:
(a) Price of the bond:
(b) Current yield:
FIND YTM
PRICE WHEN SOLD
VARIABLES
VALUES
Price = sum of PVs
Using NPV function
YTM answer =
Keep wording short.
Do NOT round intermediate calculations.
Keep the full Excel precision in every formula.
Only format the FINAL requested answer as instructed.
If a rate must be entered as a regular percent rounded to two decimals:
Excel result:
0.059146
Displayed/submitted answer:
5.91
Do not change intermediate formulas to rounded values.
Before finishing, silently check:

1. Did I ignore unrelated screenshot UI?
2. Did I identify the exact finance problem?
3. Did I use only the formulas/methods listed in this prompt?
4. Are RATE and NPER in matching periods?
5. Are cash-flow signs correct?
6. Did I use short labels such as PV, FV, PMT, NPER, RATE and YTM?
7. Did I avoid unnecessary extra input cells?
8. Did I show simple conversions such as /2, *2, /12, *12 directly in helper cells?
9. If an intermediate result is reused, did I reference that cell instead of recalculating it?
10. Did I break multi-stage problems into separate sections/cells instead of one giant formula?
11. For a semiannual bond, did I divide RATE and PMT by 2 and multiply NPER by 2?
12. For semiannual YTM, did I multiply the RATE result by 2?
13. Did I use coupon rate for coupon PMT and YTM for RATE?
14. Did I use remaining maturity if time has already passed?
15. Did I use DIV1 rather than DIV0 in Gordon Growth?
16. Did I add terminal value to the cash flow in the correct year?
17. Did I keep Time 0 outside NPV?
18. Did I avoid rounding intermediate calculations?
19. Is there one clear final answer?
20. Is the worksheet simple and easy to follow?

OUTPUT FORMAT FOR THE AUTO-TYPER (this replaces "Make the Excel cell edits" and "reply only Done" above):
You cannot edit Excel directly here. Instead, print the finished worksheet as short plain text first. Then output this exact block, with no markdown fences around it:

AUTO_TYPE_JSON_START
[{"cell":"B3","value":"VARIABLES"},{"cell":"C3","value":"VALUES"},{"cell":"B4","value":"Rate"},{"cell":"C4","value":"=0.08/2"}]
AUTO_TYPE_JSON_END

List EVERY cell that should contain something, with its address and exact contents, in the order it should be typed.
Formula cells start with = exactly as typed in Excel. Put only the cell contents in the JSON: no formatting, no dollar signs or percent signs or thousands commas in numbers (write rates as decimals like 0.05).
Keep column A and row 1 completely empty. Nothing goes in column A or row 1: no labels, no stray 0, no year or timeline numbers. Start the worksheet in column B at row 3 or lower. For a timeline, put the year numbers in a row below row 1, starting at column C, with the row labels in column B.
Every label must be a short abbreviation (PV, PMT, Cpn rate, Req ret, Sale px, Total CF), never a full spelled-out phrase. Never use FORMULATEXT.
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
