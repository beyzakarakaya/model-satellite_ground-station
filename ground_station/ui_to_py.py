#from PyQt4 import uic

from PyQt6 import uic

with open('ground_station/uydu_design.py', 'w', encoding="utf-8") as fout:
   uic.compileUi('ground_station/uydu_design.ui', fout)