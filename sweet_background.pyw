# Double-click to run Sweet with no console window and no taskbar entry (tiny dot only).
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sweet
sweet.main([])
