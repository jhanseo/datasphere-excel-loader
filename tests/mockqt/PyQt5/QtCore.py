from ._base import pyqtSignal, Obj

class _Enum(int):
    def __new__(cls, v, name=""):
        o = int.__new__(cls, v); o._name = name; return o
    def __repr__(self): return self._name or int.__repr__(self)

class Qt:
    Horizontal = 1; Vertical = 2
    AlignRight = 2
    Unchecked = _Enum(0, "Unchecked"); Checked = _Enum(2, "Checked")
    UserRole = 256
    MatchFixedString = 8
    WaitCursor = 3
    ItemIsUserCheckable = 16; ItemIsEnabled = 32; ItemIsSelectable = 1
    ItemIsEditable = 2
    white = "white"
    AA_EnableHighDpiScaling = 1; AA_UseHighDpiPixmaps = 2

class QObject(Obj):
    pass

class QThread(Obj):
    started = pyqtSignal(); finished = pyqtSignal()
    def __init__(self, parent=None): super().__init__()
    def start(self):
        self.run(); self.finished.emit()
    def run(self): pass
    def isRunning(self): return False
    def wait(self, ms=0): return True
