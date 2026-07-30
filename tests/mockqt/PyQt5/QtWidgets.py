from ._base import Obj, Signal, pyqtSignal
from .QtCore import Qt

# --- 단순 no-op 위젯들 -------------------------------------------------
class QWidget(Obj): pass
class QLabel(Obj): pass
class QGroupBox(Obj): pass
class QSplitter(Obj):
    def __init__(self, *a): super().__init__()
class QVBoxLayout(Obj):
    def __init__(self, *a): super().__init__()
class QHBoxLayout(Obj):
    def __init__(self, *a): super().__init__()
class QFormLayout(Obj):
    def __init__(self, *a): super().__init__(); self.rows = []
    def addRow(self, label, w): self.rows.append((label, w))
class QProgressBar(Obj): pass
class QPlainTextEdit(Obj):
    def __init__(self, *a): super().__init__(); self.lines = []
    def appendHtml(self, h): self.lines.append(h)
    def clear(self): self.lines = []
class QHeaderView(Obj):
    Stretch = 1; ResizeToContents = 2; Interactive = 3
class QAbstractItemView(Obj):
    NoEditTriggers = 0; ExtendedSelection = 3
class QDialogButtonBox(Obj):
    Ok = 1; Cancel = 2; Close = 4; Yes = 8; No = 16
    accepted = pyqtSignal(); rejected = pyqtSignal(); clicked = pyqtSignal(object)
    def __init__(self, *a): super().__init__()
    def button(self, which): return QPushButton()

class QPushButton(Obj):
    clicked = pyqtSignal()
    def __init__(self, text="", *a): super().__init__(text)
    def click(self): self.clicked.emit()

class QCheckBox(Obj):
    toggled = pyqtSignal(bool)
    def __init__(self, text="", *a): super().__init__(text); self._c = False
    def setChecked(self, v): self._c = bool(v); self.toggled.emit(self._c)
    def isChecked(self): return self._c

class QLineEdit(Obj):
    Password = 2; Normal = 0
    def __init__(self, text="", *a): super().__init__(text)
    def setEchoMode(self, m): pass
    def setValidator(self, v): pass

class QComboBox(Obj):
    currentIndexChanged = pyqtSignal(int)
    currentTextChanged = pyqtSignal(str)
    def __init__(self, *a):
        super().__init__(); self._items = []; self._idx = -1; self._editable = False
    def setEditable(self, v): self._editable = v
    def addItem(self, text, data=None): self._items.append((text, data))
    def addItems(self, texts):
        for t in texts: self.addItem(t)
        if self._items and self._idx < 0: self._idx = 0
    def count(self): return len(self._items)
    def findData(self, data):
        for i, (t, d) in enumerate(self._items):
            if d == data: return i
        return -1
    def findText(self, text, flags=0):
        for i, (t, d) in enumerate(self._items):
            if str(t).lower() == str(text).lower(): return i
        return -1
    def setCurrentIndex(self, i):
        if i != self._idx:
            self._idx = i
            self.currentIndexChanged.emit(i)
            self.currentTextChanged.emit(self.currentText())
    def currentIndex(self): return self._idx
    def currentText(self):
        return self._items[self._idx][0] if 0 <= self._idx < len(self._items) else ""
    def currentData(self):
        return self._items[self._idx][1] if 0 <= self._idx < len(self._items) else None
    def setCurrentText(self, text):
        i = self.findText(text)
        if i < 0:
            self._items.append((text, None)); i = len(self._items) - 1
        self.setCurrentIndex(i)

# --- 목록 / 표 / 트리 ---------------------------------------------------
class QListWidgetItem(Obj):
    def __init__(self, text="", *a):
        super().__init__(text); self._data = {}
    def setData(self, role, v): self._data[role] = v
    def data(self, role): return self._data.get(role)
    def setForeground(self, c): pass

class QListWidget(Obj):
    def __init__(self, *a): super().__init__(); self._items = []; self._sel = []
    def addItem(self, it): self._items.append(it)
    def item(self, i): return self._items[i]
    def count(self): return len(self._items)
    def row(self, it): return self._items.index(it)
    def takeItem(self, i): return self._items.pop(i)
    def clear(self): self._items = []; self._sel = []
    def selectedItems(self): return list(self._sel)
    def selectRows(self, idxs): self._sel = [self._items[i] for i in idxs]
    def setSelectionMode(self, m): pass

class QTableWidgetItem(Obj):
    def __init__(self, text="", *a):
        super().__init__(text)
        self._flags = 0; self._check = Qt.Unchecked; self._bg = None
        self._table = None; self._row = -1; self._col = -1
    def setFlags(self, f): self._flags = f
    def flags(self): return self._flags
    def setCheckState(self, s):
        self._check = s
        if self._table: self._table.itemChanged.emit(self)
    def checkState(self): return self._check
    def setBackground(self, b):
        self._bg = b
        if self._table: self._table.itemChanged.emit(self)
    def background(self): return self._bg
    def setText(self, t):
        self._text = t
        if self._table: self._table.itemChanged.emit(self)
    def row(self): return self._row
    def column(self): return self._col

class QTableWidget(Obj):
    itemChanged = pyqtSignal(object)
    def __init__(self, rows=0, cols=0, *a):
        super().__init__(); self._rows = rows; self._cols = cols
        self._cells = {}; self._widgets = {}
    def setRowCount(self, n):
        self._rows = n
        self._cells = {k: v for k, v in self._cells.items() if k[0] < n}
        self._widgets = {k: v for k, v in self._widgets.items() if k[0] < n}
    def rowCount(self): return self._rows
    def setColumnCount(self, n): self._cols = n
    def columnCount(self): return self._cols
    def setItem(self, r, c, item):
        item._table = self; item._row = r; item._col = c
        self._cells[(r, c)] = item
        self.itemChanged.emit(item)
    def item(self, r, c): return self._cells.get((r, c))
    def setCellWidget(self, r, c, w): self._widgets[(r, c)] = w
    def cellWidget(self, r, c): return self._widgets.get((r, c))
    def clear(self): self._cells = {}; self._widgets = {}
    def horizontalHeader(self): return QHeaderView()
    def verticalHeader(self): return QHeaderView()
    def columnWidth(self, c): return 100
    def setColumnWidth(self, c, w): pass
    def setHorizontalHeaderLabels(self, labels): self._headers = labels
    def resizeColumnsToContents(self): pass
    def setEditTriggers(self, t): pass

class QTreeWidgetItem(Obj):
    def __init__(self, texts=None, *a):
        super().__init__(); self._texts = list(texts or []); self._children = []
        self._data = {}; self._checks = {}; self._flags = 0; self._tree = None
    def setFlags(self, f): self._flags = f
    def flags(self): return self._flags
    def addChild(self, ch):
        ch._parent = self; ch._tree = self._tree; self._children.append(ch)
    def childCount(self): return len(self._children)
    def child(self, i): return self._children[i]
    def setData(self, col, role, v): self._data[(col, role)] = v
    def data(self, col, role): return self._data.get((col, role))
    def setCheckState(self, col, s):
        self._checks[col] = s
        if self._tree: self._tree.itemChanged.emit(self, col)
    def checkState(self, col): return self._checks.get(col, Qt.Unchecked)
    def setExpanded(self, v): pass
    def text(self, col): return self._texts[col] if col < len(self._texts) else ""

class QTreeWidget(Obj):
    itemChanged = pyqtSignal(object, int)
    currentItemChanged = pyqtSignal(object, object)
    def __init__(self, *a):
        super().__init__(); self._roots = []; self._current = None
    def clear(self): self._roots = []; self._current = None
    def addTopLevelItem(self, it): it._tree = self; self._roots.append(it)
    def topLevelItemCount(self): return len(self._roots)
    def topLevelItem(self, i): return self._roots[i]
    def setCurrentItem(self, it):
        prev, self._current = self._current, it
        self.currentItemChanged.emit(it, prev)
    def currentItem(self): return self._current
    def setHeaderLabels(self, l): pass
    def setColumnWidth(self, c, w): pass

class QTabWidget(Obj):
    def __init__(self, *a): super().__init__(); self._tabs = []; self._cur = 0
    def addTab(self, w, label): self._tabs.append([w, label]); return len(self._tabs)-1
    def count(self): return len(self._tabs)
    def setCurrentIndex(self, i): self._cur = i
    def currentIndex(self): return self._cur
    def setTabText(self, i, t): self._tabs[i][1] = t
    def tabText(self, i): return self._tabs[i][1]

# --- 대화상자 / 앱 ------------------------------------------------------
class QDialog(Obj):
    Accepted = 1; Rejected = 0
    def __init__(self, parent=None): super().__init__()
    def exec_(self): return self.Accepted
    def accept(self): pass
    def reject(self): pass

class QMainWindow(Obj):
    def __init__(self): super().__init__(); self._sb = _StatusBar()
    def statusBar(self): return self._sb
    def setCentralWidget(self, w): pass
    def setAcceptDrops(self, v): pass

class _StatusBar(Obj):
    def __init__(self): super().__init__(); self.messages = []
    def showMessage(self, m, t=0): self.messages.append(m)
    def clearMessage(self): pass

class QMessageBox(Obj):
    Yes = 1; No = 2; Cancel = 4; Ok = 8
    CALLS = []       # (kind, title, text)
    ANSWERS = []     # question()이 순서대로 꺼내 쓰는 응답
    @classmethod
    def information(cls, p, t, m, *a): cls.CALLS.append(("info", t, m)); return cls.Ok
    @classmethod
    def warning(cls, p, t, m, *a): cls.CALLS.append(("warn", t, m)); return cls.Ok
    @classmethod
    def critical(cls, p, t, m, *a): cls.CALLS.append(("crit", t, m)); return cls.Ok
    @classmethod
    def question(cls, p, t, m, *a):
        cls.CALLS.append(("question", t, m))
        return cls.ANSWERS.pop(0) if cls.ANSWERS else cls.Ok

class QFileDialog(Obj):
    @staticmethod
    def getOpenFileNames(*a, **k): return ([], "")
    @staticmethod
    def getExistingDirectory(*a, **k): return ""

class QApplication(Obj):
    _inst = None
    def __init__(self, argv=None): super().__init__()
    @staticmethod
    def setAttribute(*a): pass
    @staticmethod
    def setOverrideCursor(*a): pass
    @staticmethod
    def restoreOverrideCursor(*a): pass
    @staticmethod
    def instance(): return QApplication._inst
