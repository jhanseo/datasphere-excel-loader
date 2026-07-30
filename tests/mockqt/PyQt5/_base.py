"""아주 단순한 Qt 대역(mock). wizard.py의 로직 경로를 검증하기 위한 용도."""

class Signal:
    def __init__(self): self._slots = []
    def connect(self, fn): self._slots.append(fn)
    def disconnect(self, fn=None): self._slots = []
    def emit(self, *a):
        # PyQt5는 슬롯이 받을 수 있는 만큼만 인자를 넘긴다(초과분 절삭).
        import inspect
        for fn in list(self._slots):
            n = len(a)
            try:
                sig = inspect.signature(fn)
                if not any(pp.kind == pp.VAR_POSITIONAL
                           for pp in sig.parameters.values()):
                    n = min(n, sum(1 for pp in sig.parameters.values()
                                   if pp.kind in (pp.POSITIONAL_ONLY,
                                                  pp.POSITIONAL_OR_KEYWORD)))
            except (TypeError, ValueError):
                pass
            fn(*a[:n])

class SignalDesc:
    def __init__(self, *types): pass
    def __get__(self, obj, cls):
        if obj is None: return self
        key = "_sig_%d" % id(self)
        if not hasattr(obj, key): setattr(obj, key, Signal())
        return getattr(obj, key)

pyqtSignal = SignalDesc

class Obj:
    """알려지지 않은 메서드는 조용히 무시하는 기본 위젯."""
    def __init__(self, *a, **k):
        self._text = a[0] if a and isinstance(a[0], str) else ""
        self._enabled = True
        self._tooltip = ""
    def __getattr__(self, name):
        if name.startswith("_"): raise AttributeError(name)
        return lambda *a, **k: None
    def setEnabled(self, v): self._enabled = v
    def isEnabled(self): return self._enabled
    def setText(self, t): self._text = t
    def text(self): return self._text
    def setToolTip(self, t): self._tooltip = t
    def toolTip(self): return self._tooltip
