#!/usr/bin/env python3
"""
Sweet -- Finance Screenshot Analyzer
======================================
Install dependencies:
  pip install google-generativeai Pillow keyboard pyautogui pyperclip

Set your Gemini API key (required):
  Windows:    set GEMINI_API_KEY=your_key_here
  Mac/Linux:  export GEMINI_API_KEY=your_key_here

Run (background dot only, no taskbar entry):
  python sweet.py

Run (status bar visible in taskbar):
  python sweet.py --taskbar

Hotkeys (work system-wide):
  Ctrl+Shift+S  ->  Screenshot + analyze finance problem
  Ctrl+Shift+T  ->  Auto-type last result into Excel (you have 3 sec to focus Excel)
  Ctrl+Shift+Q  ->  Quit Sweet

Note: On Windows the keyboard library usually works without admin rights.
      On Linux it requires root (sudo python sweet.py).
"""

import sys
import os
import time
import threading
import json
import re
import argparse
import tkinter as tk

# ─── dependency check ──────────────────────────────────────────────────────────
_need = []
try:
    import google.generativeai as genai
except ImportError:
    _need.append("google-generativeai")
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
    pyautogui.FAILSAFE = True
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

# ─── configuration ─────────────────────────────────────────────────────────────
API_KEY          = os.environ.get("GEMINI_API_KEY", "")
MODEL_NAME       = "gemini-2.0-flash"   # user requested "Gemini 3.6 Flash" -- using closest available
HOTKEY_ANALYZE   = "ctrl+shift+s"
HOTKEY_TYPE      = "ctrl+shift+t"
HOTKEY_QUIT      = "ctrl+shift+q"
TYPE_START_DELAY = 3       # seconds before auto-type starts (time to focus Excel)
CELL_PAUSE       = 0.15    # pause between each cell during auto-type (seconds)

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
Rate, Nper, Pmt, PV, FV, Type, Periodic Rate, Payments per Year, Payment Period, Loan Amount,
Beginning Balance, Interest, Principal, Ending Balance, Coupon, Coupon Rate, Face Value, Price,
YTM, Annual YTM, Current Yield, Required Return, Dividend, DIV0, DIV1, Growth Rate,
Terminal Value, Additional Value, Total Cash Flow, Sale Price, Initial Investment, Total Return,
Risk-Free Rate, MRP, DRP, Payout Ratio, Plowback Ratio, ROE, SGR.
Choose the label that best describes each value.
Numbers supplied by the problem should normally be entered into their own cells.
Then reference those cells in later formulas.
If an input needs to be converted, calculate the conversion in a worksheet cell.
Do not round these intermediate calculations.
Represent percentages as decimals in Excel: 8% = 0.08, 6.85% = 0.0685, 2% = 0.02, 20% = 0.20.
Only divide a rate when converting it to a different period:
  Monthly: =AnnualRateCell/12
  Quarterly: =AnnualRateCell/4
  Semiannual: =AnnualRateCell/2
  Weekly: =AnnualRateCell/52
If a calculation requires a base cash flow or face value but the problem intentionally does not specify one, use 1000.
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
Prefer these functions over manually expanding a time-value-of-money equation.
Use consistent Excel cash-flow signs:
  Cash paid and cash received should have opposite signs.
  Investment today = negative, payments made = negative, loan amount received = positive,
  future cash received = positive, bond purchase price = negative,
  bond coupons received = positive, face value received = positive.
Ordinary annuity: payments at END of period, Type = 0.
Annuity due: payments at BEGINNING of period, Type = 1.
For a loan, use Rate, Nper, Pmt, PV, FV, Type block.
For amortization use PMT, IPMT, PPMT and calculate Ending Balance separately.
Coupon Rate determines the coupon payment. YTM determines the discount rate.
Annual Coupon = FaceValue * CouponRate.
Semiannual Coupon = FaceValue * CouponRate / 2.
For a semiannual bond: Periodic Rate = AnnualYTM/2, Nper = Years*2.
Use remaining maturity rather than original maturity when applicable.
For stock valuation:
  Gordon Growth uses DIV1 (the NEXT dividend, not the one just paid).
  Terminal value numerator must use the NEXT dividend.
  For unequal dividends, use individual PV calculations, not annuity.
  Do not place Time 0 inside NPV.
EAR: =EFFECT(APRCell, PeriodsPerYearCell)
Continuous compounding EAR: =EXP(APRCell)-1
Real rate: =((1+NominalRateCell)/(1+InflationCell))-1
The worksheet should end with one clearly identified final result corresponding to what the problem asks for.
Before finishing, verify: periodic rates are converted correctly, Rate and Nper use matching periods,
cash-flow signs are consistent, semiannual YTM is converted back to annual, Gordon Growth uses DIV1,
terminal value uses the next dividend, unequal dividends are not treated as an annuity,
Time 0 is not placed inside NPV, intermediate calculations are not rounded.

CRITICAL: After your worksheet, output this exact block with no markdown fences around it:

AUTO_TYPE_JSON_START
[{"cell":"B2","value":"Rate"},{"cell":"C2","value":"0.08"},{"cell":"B3","value":"Nper"}]
AUTO_TYPE_JSON_END

List EVERY cell in the worksheet with its address and exact value.
Formula cells must start with = exactly as they would be typed in Excel.
Ensure the JSON is valid (double quotes only, no trailing commas).
"""

# ─── shared state ──────────────────────────────────────────────────────────────
_lock  = threading.Lock()
_busy  = False       # True while an analysis is running
_cells = []          # [(addr, value), ...] from last successful analysis


# ─── UI window ─────────────────────────────────────────────────────────────────
class SweetWindow:
    """
    Background mode (default):
      Tiny animated dot in the top-right corner. No taskbar entry.
      Hidden when idle, orange pulsing dot while working, red ! on error.

    Taskbar mode (--taskbar):
      Slim status bar in the top-right corner, visible in the taskbar.
      Text label changes color to reflect state.
    """

    DOT_SIZE = 20

    def __init__(self, taskbar: bool):
        self.taskbar = taskbar
        self.root    = tk.Tk()
        self._state  = "idle"
        self._phase  = 0.0    # animation pulse 0.0-1.0
        self._dir    = 1.0    # animation direction
        self._build()

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
            self._lbl = tk.Label(
                r, textvariable=self._var,
                bg="#111111", fg="#777777",
                font=("Consolas", 9), anchor="w", padx=8
            )
            self._lbl.pack(fill="both", expand=True)

        else:
            r.overrideredirect(True)           # no title bar, not in taskbar
            r.wm_attributes("-topmost", True)
            r.wm_attributes("-alpha", 0.92)
            pad = 6
            sz  = self.DOT_SIZE + pad * 2
            r.geometry(f"{sz}x{sz}+{sw - sz - 4}+4")
            r.configure(bg="#000000")
            cv = tk.Canvas(r, width=sz, height=sz,
                           bg="#000000", highlightthickness=0)
            cv.pack()
            self._cv  = cv
            self._dot = cv.create_oval(pad, pad,
                                       pad + self.DOT_SIZE, pad + self.DOT_SIZE,
                                       fill="#2a2a2a", outline="")
            self._txt = cv.create_text(pad + self.DOT_SIZE // 2,
                                       pad + self.DOT_SIZE // 2,
                                       text="", fill="white",
                                       font=("Arial", 9, "bold"))
            r.withdraw()   # start hidden

    # ── public state API ────────────────────────────────────────────────────────
    def set_state(self, state: str):
        """Thread-safe. state in {'idle', 'working', 'error'}"""
        self._state = state
        self.root.after(0, self._apply)

    def _apply(self):
        s = self._state
        if self.taskbar:
            cfg = {
                "idle":    ("Sweet  |  ready",    "#777777"),
                "working": ("Sweet  |  working.", "#FFA020"),
                "error":   ("Sweet  |  error  !", "#FF3333"),
            }
            txt, col = cfg.get(s, ("Sweet", "#777777"))
            self._var.set(txt)
            self._lbl.configure(fg=col)
        else:
            if s == "idle":
                self.root.withdraw()
            else:
                self.root.deiconify()
                if s == "working":
                    self._cv.itemconfig(self._dot, fill="#FFA500")
                    self._cv.itemconfig(self._txt, text="")
                    self._phase, self._dir = 0.0, 1.0
                    self._tick()
                elif s == "error":
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


# ─── Gemini screenshot analysis ────────────────────────────────────────────────
def analyze(win: SweetWindow):
    global _busy, _cells

    with _lock:
        if _busy:
            print("[Sweet] Still working -- please wait.")
            return
        _busy = True

    win.set_state("working")
    success = False

    try:
        if not API_KEY:
            raise RuntimeError(
                "GEMINI_API_KEY is not set.\n"
                "  Windows:    set GEMINI_API_KEY=your_key_here\n"
                "  Mac/Linux:  export GEMINI_API_KEY=your_key_here"
            )

        # Hide overlay briefly so it's not captured in the screenshot
        win.root.after(0, win.root.withdraw)
        time.sleep(0.22)
        screenshot = ImageGrab.grab()
        win.root.after(0, win.root.deiconify)

        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel(MODEL_NAME, system_instruction=SYSTEM_PROMPT)

        print("[Sweet] Sending screenshot to Gemini...")
        resp  = model.generate_content([
            "Analyze the finance problem shown in this screenshot and produce the Excel worksheet:",
            screenshot,
        ])
        text  = resp.text
        cells = _extract_cells(text)

        with _lock:
            _cells = cells

        clean = re.sub(
            r"AUTO_TYPE_JSON_START[\s\S]*?AUTO_TYPE_JSON_END", "", text, flags=re.IGNORECASE
        ).strip()
        print("\n" + "=" * 66)
        print(clean)
        print("=" * 66)
        print(f"[Sweet] {len(cells)} cells ready.  "
              f"Press {HOTKEY_TYPE.upper()} to auto-type into Excel.")

        success = True

    except Exception as exc:
        print(f"\n[Sweet] ERROR: {exc}")

    finally:
        win.set_state("idle" if success else "error")
        with _lock:
            _busy = False


def _extract_cells(text: str) -> list:
    """Parse AUTO_TYPE_JSON block from Gemini response."""
    m = re.search(
        r"AUTO_TYPE_JSON_START\s*(\[[\s\S]*?\])\s*AUTO_TYPE_JSON_END",
        text, re.IGNORECASE
    )
    if not m:
        print("[Sweet] Warning -- AUTO_TYPE_JSON block not found in response.")
        return []
    try:
        data = json.loads(m.group(1))
        return [
            (item["cell"].strip().upper(), str(item["value"]).strip())
            for item in data
        ]
    except Exception as e:
        print(f"[Sweet] JSON parse error: {e}")
        return []


# ─── auto-typer ────────────────────────────────────────────────────────────────
def auto_type(win: SweetWindow):
    """
    Navigate Excel cell-by-cell via Ctrl+G (Go To) and paste each value.
    Focus Excel before the countdown reaches zero.
    """
    with _lock:
        cells = list(_cells)

    if not cells:
        print(f"[Sweet] No data to type -- run {HOTKEY_ANALYZE.upper()} first.")
        win.set_state("error")
        time.sleep(1.5)
        win.set_state("idle")
        return

    print(f"[Sweet] Click on Excel now -- typing starts in {TYPE_START_DELAY}s...")
    for n in range(TYPE_START_DELAY, 0, -1):
        print(f"  {n}...")
        time.sleep(1.0)

    win.set_state("working")
    print(f"[Sweet] Typing {len(cells)} cells into Excel...")

    try:
        for addr, value in cells:
            _goto_cell(addr)
            _paste_value(value)
            time.sleep(CELL_PAUSE)

        print("[Sweet] Auto-type complete!")
        win.set_state("idle")

    except Exception as exc:
        print(f"[Sweet] Auto-type error: {exc}")
        win.set_state("error")


def _goto_cell(addr: str):
    """Use Excel's Go To dialog (Ctrl+G / F5) to navigate to a cell address."""
    pyautogui.hotkey("ctrl", "g")
    time.sleep(0.25)
    # Select all existing content in Reference field, then clear it
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.06)
    pyautogui.press("delete")
    time.sleep(0.04)
    # Type the cell address (always simple: B2, C10, D6, etc.)
    pyautogui.typewrite(addr, interval=0.04)
    pyautogui.press("enter")
    time.sleep(0.20)


def _paste_value(value: str):
    """Paste a value into the active Excel cell using the clipboard."""
    pyperclip.copy(value)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.08)
    pyautogui.press("enter")   # confirm entry


# ─── hotkeys ───────────────────────────────────────────────────────────────────
def register_hotkeys(win: SweetWindow):
    keyboard.add_hotkey(
        HOTKEY_ANALYZE,
        lambda: threading.Thread(target=analyze, args=(win,), daemon=True).start(),
        suppress=True,
    )
    keyboard.add_hotkey(
        HOTKEY_TYPE,
        lambda: threading.Thread(target=auto_type, args=(win,), daemon=True).start(),
        suppress=True,
    )
    keyboard.add_hotkey(
        HOTKEY_QUIT,
        lambda: (keyboard.unhook_all(), win.root.after(0, win.root.quit)),
        suppress=True,
    )

    bar = "+" + "-" * 40 + "+"
    print(f"\n{bar}")
    print(f"|{'Sweet is running':^40}|")
    print(f"{bar}")
    print(f"|  {HOTKEY_ANALYZE.upper():<14}  analyze screen       |")
    print(f"|  {HOTKEY_TYPE.upper():<14}  auto-type into Excel  |")
    print(f"|  {HOTKEY_QUIT.upper():<14}  quit                  |")
    print(f"{bar}\n")


# ─── entry point ───────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(
        description="Sweet -- Finance Screenshot Analyzer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Set GEMINI_API_KEY before running:\n"
            "  Windows:    set GEMINI_API_KEY=your_key_here\n"
            "  Mac/Linux:  export GEMINI_API_KEY=your_key_here"
        ),
    )
    ap.add_argument(
        "--taskbar", action="store_true",
        help="Show status bar in taskbar (default: tiny background dot only)",
    )
    args = ap.parse_args()

    if not API_KEY:
        print(
            "[Sweet] WARNING: GEMINI_API_KEY not set -- analysis will fail.\n"
            "  Windows:    set GEMINI_API_KEY=your_key_here\n"
            "  Mac/Linux:  export GEMINI_API_KEY=your_key_here\n"
        )

    win = SweetWindow(taskbar=args.taskbar)
    register_hotkeys(win)
    win.mainloop()
    print("[Sweet] Goodbye.")


if __name__ == "__main__":
    main()
