from ._base import Obj
class QColor(Obj):
    def __init__(self, v=None): super().__init__(); self.v = v
    def __eq__(self, o): return isinstance(o, QColor) and o.v == self.v
class QFont(Obj):
    def __init__(self, *a): super().__init__()
class QIntValidator(Obj):
    def __init__(self, *a): super().__init__()
