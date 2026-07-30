# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec :: ExcelUploadWizard  (PyInstaller 6.x)

빌드:
    build_exe.bat          (권장)
    또는  python -m PyInstaller --clean --noconfirm wizard.spec

빌드 방식 (아래 ONEFILE 값으로 선택)
    ONEFILE = False (기본, 권장)
        dist/ExcelUploadWizard/ 폴더가 만들어지고 그 안에 exe가 들어간다.
        - exe 자체가 작아서(1~2MB) 백신/DRM이 파일을 잠글 확률이 훨씬 낮다
          -> "Failed to compute PE checksum / Error code 32" 회피
        - 실행할 때마다 %TEMP%에 압축을 풀지 않으므로 실행이 빠르다
        - 배포할 때는 폴더째 zip으로 묶어서 전달한다
    ONEFILE = True
        dist/ExcelUploadWizard.exe 파일 하나만 만들어진다.
        배포는 편하지만 백신에 걸리거나 실행이 느릴 수 있다.
"""

import importlib.util

ONEFILE = False

# 설치되어 있는 모듈만 hiddenimports에 넣는다.
# (없는 모듈을 적으면 빌드 로그에 "Hidden import not found" ERROR가 찍혀
#  빌드가 실패한 것처럼 보인다. 실제 실패는 아니지만 혼동을 준다.)
_CANDIDATES = [
    "hdbcli",
    "hdbcli.dbapi",
    "openpyxl",
    "xlrd",
    # DRM(DRMONE) 파일 처리 - Windows + Excel 환경에서만 설치되어 있다
    "xlwings",
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
]


def _available(name):
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError, AttributeError):
        return False


hidden = [m for m in _CANDIDATES if _available(m)]
print("[spec] ONEFILE =", ONEFILE)
print("[spec] hiddenimports =", hidden)

a = Analysis(
    ["wizard.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "scipy", "tkinter", "notebook", "IPython", "pytest",
        # 다른 Qt 바인딩이 함께 잡히면 exe가 커지고 충돌이 생길 수 있다
        "PySide2", "PySide6", "PyQt6",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

_COMMON = dict(
    name="ExcelUploadWizard",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,              # UPX 압축은 백신 오탐을 크게 늘린다. 끈 채로 둔다.
    console=False,          # GUI 전용 (콘솔 창 숨김)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # 아이콘 파일이 있으면 "icon.ico" 로 지정
)

if ONEFILE:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        upx_exclude=[],
        runtime_tmpdir=None,
        **_COMMON,
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        **_COMMON,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        upx_exclude=[],
        name="ExcelUploadWizard",
    )
