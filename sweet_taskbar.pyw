# Double-click to run Sweet with no console window and a small status bar in the taskbar.
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sweet
sweet.main(["--taskbar"])
