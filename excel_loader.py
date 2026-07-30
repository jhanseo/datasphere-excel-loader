# -*- coding: utf-8 -*-
"""
Excel -> HANA Datasphere Loader (core)
======================================
Excel 파일들을 읽어, 시트 이름이 keyword로 시작하는 시트를
SAP Datasphere(HANA) Open SQL Schema의 테이블로 업로드한다.

- 테이블이 없으면 데이터 타입을 자동 추론하여 CREATE TABLE 후 INSERT
- 적재 모드(load_mode): truncate(비우고 적재) / insert(이어서 추가) /
  upsert(PRIMARY KEY 기준 갱신+추가)
- 미리보기에서 컬럼별 [KEY]/[NOT NULL]을 지정하면 CREATE TABLE에 반영된다
- 시트 이름 규칙: sheet_name_keyword(기본 "TABLE") 접두어 제거 -> 테이블명
  예) 시트 "TABLE_T_SALES" -> 테이블 "T_SALES"

이 모듈은 CLI로도, Wizard GUI(wizard.py)의 엔진으로도 사용된다.

CLI 실행:
    python excel_loader.py --config config.ini
    python excel_loader.py --config config.ini --dry-run
    python excel_loader.py --config config.ini --file a.xlsx --file b.xlsx
    python excel_loader.py --only-sheet TABLE_T_SALES,TEST_TABLE
    python excel_loader.py --only-table T_SALES,T_CUSTOMER

GUI 실행:
    python wizard.py

필요 라이브러리:
    pip install pandas openpyxl xlrd hdbcli
    # DRM(DRMONE) 파일 처리(Windows + Excel 설치 환경):
    pip install xlwings pywin32
    # Wizard GUI:
    pip install PyQt5
"""

import argparse
import configparser
import datetime
import decimal
import logging
import math
import os
import re
import shutil
import sys

import pandas as pd

__all__ = [
    "DrmProtectedError",
    "TypeMismatchAbort",
    "HANA_TYPE_CHOICES",
    "CONFIG_SPEC",
    "SECTION_LABELS",
    "DRM_GUIDE_MESSAGE",
    "load_config",
    "save_config",
    "get_bool",
    "setup_logging",
    "build_options",
    "is_drmone_file",
    "open_excel_source",
    "close_excel_source",
    "read_sheet",
    "infer_hana_type",
    "normalize_hana_type",
    "parse_hana_type",
    "check_series",
    "apply_column_types",
    "collect_sheet_plans",
    "execute_plans",
    "connect_hana",
    "set_ddl_transactional",
    "check_target_tables",
    "check_constraints",
    "get_existing_column_types",
    "get_primary_key_columns",
    "planned_column_types",
    "list_excel_files",
    "LOAD_MODES",
    "LOAD_MODE_LABELS",
    "LOAD_MODE_TRUNCATE",
    "LOAD_MODE_INSERT",
    "LOAD_MODE_UPSERT",
]


# ---------------------------------------------------------------------------
# 적재 모드
# ---------------------------------------------------------------------------

LOAD_MODE_TRUNCATE = "truncate"
LOAD_MODE_INSERT = "insert"
LOAD_MODE_UPSERT = "upsert"

LOAD_MODES = [LOAD_MODE_TRUNCATE, LOAD_MODE_INSERT, LOAD_MODE_UPSERT]

LOAD_MODE_LABELS = {
    LOAD_MODE_TRUNCATE: "TRUNCATE (기존 데이터를 모두 지우고 새로 적재)",
    LOAD_MODE_INSERT: "INSERT (기존 데이터에 이어서 추가)",
    LOAD_MODE_UPSERT: "UPSERT (키가 같으면 갱신, 없으면 추가)",
}

LOAD_MODE_DESCRIPTIONS = {
    LOAD_MODE_TRUNCATE:
        "업로드 전에 TRUNCATE로 테이블을 비웁니다. 엑셀 내용이 곧 테이블 전체 "
        "내용이 되며, 삭제된 공간도 즉시 반환됩니다.",
    LOAD_MODE_INSERT:
        "기존 데이터를 지우지 않고 뒤에 붙입니다. 같은 파일을 두 번 올리면 "
        "중복 적재되므로 주의하세요.",
    LOAD_MODE_UPSERT:
        "PRIMARY KEY가 일치하는 행은 갱신하고 없는 행만 추가합니다. "
        "대상 테이블에 PRIMARY KEY가 반드시 있어야 하며, 미리보기에서 "
        "[KEY]로 지정한 컬럼이 키가 됩니다.",
}


# ---------------------------------------------------------------------------
# 예외 / 안내 문구
# ---------------------------------------------------------------------------

DRM_GUIDE_MESSAGE = (
    "'{name}' 파일은 DRM(문서보안)으로 암호화되어 있어 일반 실행으로는 "
    "읽을 수 없습니다.\n\n"
    "상단의 [미리보기] 버튼을 먼저 실행해 주세요.\n"
    "미리보기는 설치된 Excel을 통해 DRM을 우회하여 파일을 열고, 그때 읽어온 "
    "데이터를 그대로 업로드에 사용합니다.\n\n"
    "(미리보기 실행 중 Excel 창에서 보안 로그인 또는 열기 승인 창이 나타나면 "
    "완료해 주세요.)"
)


class DrmProtectedError(Exception):
    """DRM(DRMONE)으로 보호된 파일을 미리보기 없이 열려고 할 때 발생."""

    def __init__(self, path, message=None):
        self.path = path
        self.file_name = os.path.basename(path)
        super().__init__(message or DRM_GUIDE_MESSAGE.format(name=self.file_name))


class TypeMismatchAbort(Exception):
    """'업로드 중단' 정책에서 유형 불일치 발견 시 발생."""


# ---------------------------------------------------------------------------
# 설정 스펙 -- Wizard의 탭 구성과 CLI가 같은 정의를 공유한다.
#   (key, label, kind, default, help)
#   kind: text | password | int | bool | choice:a|b|c | dir
# ---------------------------------------------------------------------------

CONFIG_SPEC = [
    ("HANA", [
        ("host", "호스트", "text", "", "Datasphere DB 호스트 주소"),
        ("port", "포트", "int", "443", "일반적으로 443"),
        ("user", "사용자", "text", "", "Open SQL Schema 데이터베이스 사용자"),
        ("password", "비밀번호", "password", "", "데이터베이스 사용자 비밀번호"),
        ("schema", "스키마", "text", "", "업로드 대상 Open SQL Schema 이름"),
        ("proxy_host", "프록시 호스트", "text", "", "사내 프록시 경유가 필요할 때만 입력"),
        ("proxy_port", "프록시 포트", "int", "8080", "프록시 호스트를 입력한 경우에만 사용"),
    ]),
    ("GENERAL", [
        ("sheet_name_keyword", "시트 이름 키워드", "text", "TABLE",
         "이 문자열로 시작하는 시트만 처리 (대소문자 무시)"),
        ("load_mode", "적재 모드", "choice:truncate|insert|upsert", "truncate",
         "truncate=비우고 적재 / insert=이어서 추가 / upsert=키 기준 갱신+추가"),
        ("use_filename_prefix_for_table", "테이블명에 파일명 prefix 사용", "bool", "false",
         "예) SALES.xlsx + TABLE_T_A -> SALES_T_A"),
        ("xlwings_visible", "DRM 처리 시 Excel 창 표시", "bool", "true",
         "보안 로그인 창을 직접 확인하려면 켜 두세요"),
        ("excel_dir", "기본 엑셀 폴더", "dir", "",
         "파일 선택 대화상자가 처음 열리는 위치 (CLI에서는 처리 대상 폴더)"),
    ]),
    ("HEADER", [
        ("auto_header_detect", "헤더 행 자동 감지", "bool", "true",
         "헤더 행 번호가 비어 있을 때 문자열이 가장 많은 행을 헤더로 추정"),
        ("header_row", "헤더 행 번호", "text", "",
         "0부터 시작. 비워 두면 자동 감지 사용"),
        ("use_description_row", "헤더 다음 행을 컬럼 설명으로 사용", "bool", "true",
         "해당 행은 데이터에서 제외되고 컬럼 COMMENT로 등록됩니다"),
        ("preview_rows", "미리보기 표시 행 수", "int", "100",
         "미리보기 화면에 보여줄 최대 데이터 행 수 (업로드는 전체 행)"),
    ]),
    ("LOGGING", [
        ("log_dir", "로그 폴더", "dir", "logs", "상대경로면 config.ini 위치 기준"),
        ("log_level", "로그 레벨", "choice:DEBUG|INFO|WARNING|ERROR", "INFO", ""),
    ]),
    ("ARCHIVE", [
        ("move_processed_files", "처리 완료 파일 이동", "bool", "false",
         "업로드에 성공한 원본 파일을 archive 폴더로 이동"),
        ("archive_dir", "archive 폴더", "text", "archive",
         "상대경로면 원본 파일이 있는 폴더 하위에 생성"),
    ]),
]

SECTION_LABELS = {
    "HANA": "HANA 접속",
    "GENERAL": "일반",
    "HEADER": "헤더",
    "LOGGING": "로깅",
    "ARCHIVE": "보관",
}

BOOL_TRUE = ("1", "true", "yes", "on", "y")
BOOL_FALSE = ("0", "false", "no", "off", "n", "fale")  # 'fale' 오타 허용


# ---------------------------------------------------------------------------
# 설정 / 로깅
# ---------------------------------------------------------------------------

def ensure_config_sections(cfg: configparser.ConfigParser) -> configparser.ConfigParser:
    """CONFIG_SPEC에 정의된 섹션/키가 없으면 기본값으로 채운다."""
    for section, fields in CONFIG_SPEC:
        if not cfg.has_section(section):
            cfg.add_section(section)
        for key, _label, _kind, default, _help in fields:
            if not cfg.has_option(section, key):
                cfg.set(section, key, default)
    return cfg


def load_config(config_path: str, required: bool = True) -> configparser.ConfigParser:
    cfg = configparser.ConfigParser(inline_comment_prefixes=("#", ";"))
    if not os.path.isfile(config_path):
        if required:
            print(f"[오류] 설정 파일을 찾을 수 없습니다: {config_path}")
            sys.exit(1)
        return ensure_config_sections(cfg)
    cfg.read(config_path, encoding="utf-8")
    return ensure_config_sections(cfg)


def save_config(cfg: configparser.ConfigParser, config_path: str) -> None:
    parent = os.path.dirname(os.path.abspath(config_path))
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(config_path, "w", encoding="utf-8") as f:
        cfg.write(f)


def get_bool(cfg, section, keys, default=False):
    """설정에서 boolean 값을 읽는다. 오타 키(fallback key)도 지원."""
    if isinstance(keys, str):
        keys = [keys]
    for key in keys:
        if cfg.has_option(section, key):
            raw = cfg.get(section, key).strip().lower()
            if raw in BOOL_TRUE:
                return True
            if raw in BOOL_FALSE:
                return False
    return default


def as_bool(raw, default=False):
    if isinstance(raw, bool):
        return raw
    text = str(raw if raw is not None else "").strip().lower()
    if text in BOOL_TRUE:
        return True
    if text in BOOL_FALSE:
        return False
    return default


def setup_logging(base_dir: str, cfg: configparser.ConfigParser,
                  extra_handlers=None) -> logging.Logger:
    log_dir = cfg.get("LOGGING", "log_dir", fallback="logs").strip() or "logs"
    if not os.path.isabs(log_dir):
        log_dir = os.path.join(base_dir, log_dir)
    os.makedirs(log_dir, exist_ok=True)

    level_name = cfg.get("LOGGING", "log_level", fallback="INFO").strip().upper()
    level = getattr(logging, level_name, logging.INFO)

    # 일자별 로그 파일 (같은 날 여러 번 실행 시 append)
    log_file = os.path.join(
        log_dir, f"excel_loader_{datetime.datetime.now():%Y%m%d}.txt"
    )
    logger = logging.getLogger("excel_loader")
    logger.setLevel(level)
    logger.propagate = False
    # 같은 프로세스에서 다시 호출해도 로그가 중복 출력되지 않게 한다.
    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)
    fmt = logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_file, mode="a", encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    if sys.stdout is not None:
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(fmt)
        logger.addHandler(ch)
    for handler in (extra_handlers or []):
        handler.setFormatter(fmt)
        logger.addHandler(handler)
    logger.debug("로그 파일: %s (level=%s)", log_file, level_name)
    return logger


def build_options(cfg: configparser.ConfigParser, base_dir: str) -> dict:
    """config -> 처리 옵션 dict. GUI/CLI 공통."""
    excel_dir = cfg.get("GENERAL", "excel_dir", fallback="").strip()
    if excel_dir and not os.path.isabs(excel_dir):
        excel_dir = os.path.abspath(os.path.join(base_dir, excel_dir))
    header_cfg = cfg.get("HEADER", "header_row", fallback="").strip() or None
    try:
        preview_rows = int(cfg.get("HEADER", "preview_rows", fallback="100").strip() or 100)
    except ValueError:
        preview_rows = 100
    # 적재 모드. 예전 설정(truncate_before_insert)도 계속 읽어 준다.
    load_mode = cfg.get("GENERAL", "load_mode", fallback="").strip().lower()
    if load_mode not in LOAD_MODES:
        load_mode = (LOAD_MODE_TRUNCATE
                     if get_bool(cfg, "GENERAL", "truncate_before_insert",
                                 default=True)
                     else LOAD_MODE_INSERT)
    return {
        "base_dir": base_dir,
        "excel_dir": excel_dir,
        "keyword": cfg.get("GENERAL", "sheet_name_keyword",
                           fallback="TABLE").strip() or "TABLE",
        "load_mode": load_mode,
        "truncate": load_mode == LOAD_MODE_TRUNCATE,
        "use_prefix": get_bool(cfg, "GENERAL", "use_filename_prefix_for_table",
                               default=False),
        "xlwings_visible": get_bool(cfg, "GENERAL", "xlwings_visible", default=True),
        "header_cfg": header_cfg,
        "auto_header": get_bool(cfg, "HEADER", "auto_header_detect", default=True),
        "use_desc_row": get_bool(cfg, "HEADER", "use_description_row", default=False),
        "preview_rows": max(1, preview_rows),
        "schema": cfg.get("HANA", "schema", fallback="").strip(),
        "move_processed": get_bool(cfg, "ARCHIVE",
                                   ["move_processed_files", "move_processed_fiels"],
                                   default=False),
        "archive_dir_name": cfg.get("ARCHIVE", "archive_dir",
                                    fallback="archive").strip() or "archive",
    }


# ---------------------------------------------------------------------------
# 이름 처리
# ---------------------------------------------------------------------------

def sheet_to_table_name(sheet_name: str, keyword: str) -> str:
    """시트 이름에서 keyword 접두어를 제거해 테이블 이름을 만든다.
    예) TABLE_T_SALES -> T_SALES / MY_TABLE_X -> MY_TABLE_X(접두어 아님, 그대로)
    """
    name = sheet_name.strip()
    pattern = re.compile(rf"^{re.escape(keyword)}[_\- ]+", re.IGNORECASE)
    stripped = pattern.sub("", name)
    return sanitize_identifier(stripped if stripped else name)


def sanitize_identifier(name: str) -> str:
    """HANA 식별자로 안전하게 변환 (공백/특수문자 -> _, 대문자화)."""
    name = str(name).strip()
    name = re.sub(r"[^\w가-힣]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name.upper()


# ---------------------------------------------------------------------------
# Excel 열기 / Header 탐지
# ---------------------------------------------------------------------------

def get_excel_engine(excel_path: str) -> str:
    """확장자보다 파일 시그니처를 우선해 pandas Excel 엔진을 선택한다."""
    ext = os.path.splitext(excel_path)[1].lower()

    with open(excel_path, "rb") as f:
        signature = f.read(8)

    # XLSX/XLSM은 ZIP 컨테이너(PK) 형식이다.
    if signature.startswith(b"PK"):
        return "openpyxl"

    # 구형 XLS는 OLE Compound Document 형식이다. 확장자가 잘못되어도 읽는다.
    if signature == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xlrd"

    if b"DRMONE" in signature.upper():
        raise DrmProtectedError(excel_path)

    if signature.lstrip().startswith((b"<", b"<!DOCTYPE", b"<?xml")):
        actual = "HTML/XML"
    elif b"," in signature or b"\t" in signature:
        actual = "CSV/TSV 또는 텍스트"
    else:
        actual = f"알 수 없는 형식(시그니처={signature.hex(' ')})"

    raise ValueError(
        f"Excel 파일 내부 형식이 아닙니다: {actual}. "
        f"파일명 확장자는 {ext or '(없음)'}입니다. "
        "Excel에서 파일을 연 뒤 'Excel 통합 문서(.xlsx)'로 다시 저장하세요."
    )


def is_drmone_file(excel_path: str) -> bool:
    try:
        with open(excel_path, "rb") as f:
            return b"DRMONE" in f.read(16).upper()
    except OSError:
        return False


def open_excel_source(excel_path: str, logger, xlwings_visible: bool = True,
                      allow_drm: bool = True):
    """일반 파일은 pandas, DRM(DRMONE) 파일은 실제 Excel(xlwings)로 연다.

    allow_drm=False 이면 DRM 파일을 만났을 때 DrmProtectedError를 던진다.
    (미리보기를 거치지 않은 '그냥 실행' 경로에서 사용)
    """
    if not is_drmone_file(excel_path):
        engine = get_excel_engine(excel_path)
        xls = pd.ExcelFile(excel_path, engine=engine)
        return {
            "kind": "pandas",
            "engine": engine,
            "handle": xls,
            "sheet_names": list(xls.sheet_names),
        }

    if not allow_drm:
        raise DrmProtectedError(excel_path)

    if sys.platform != "win32":
        raise RuntimeError(
            "DRM 파일은 Excel이 설치된 Windows 환경에서만 미리보기로 열 수 있습니다."
        )

    try:
        import xlwings as xw
    except ImportError as e:
        raise ImportError(
            "DRM 파일을 열려면 xlwings가 필요합니다. "
            "'pip install xlwings pywin32'로 설치하세요."
        ) from e

    logger.info("[DRM] DRM(DRMONE) 파일 감지 -> xlwings/Excel로 열기")
    logger.info("[DRM] Excel 창에서 보안 로그인 또는 열기 승인이 나타나면 완료해 주세요.")
    app = None
    book = None
    try:
        app = xw.App(visible=xlwings_visible, add_book=False)
        app.screen_updating = xlwings_visible
        app.display_alerts = True
        book = app.books.open(
            os.path.abspath(excel_path),
            update_links=False,
            read_only=True,
            ignore_read_only_recommended=True,
        )
        return {
            "kind": "xlwings",
            "app": app,
            "handle": book,
            "sheet_names": [sheet.name for sheet in book.sheets],
        }
    except Exception:
        if book is not None:
            try:
                book.close()
            except Exception:
                pass
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
        raise


def close_excel_source(source, logger):
    if not source:
        return
    try:
        source["handle"].close()
    except Exception as e:
        logger.warning("Excel 파일 닫기 중 경고: %s", e)
    if source.get("kind") == "xlwings":
        try:
            source["app"].quit()
        except Exception as e:
            logger.warning("Excel 종료 중 경고: %s", e)


def detect_header_row(raw: pd.DataFrame) -> int:
    """상단 20행 중 컬럼명으로 보이는 문자열이 가장 많은 행을 찾는다."""
    best_idx = 0
    best_score = (-1, -1)
    for idx in range(min(len(raw), 20)):
        row = raw.iloc[idx]
        non_null = row.dropna()
        if len(non_null) < 2:
            continue
        str_cnt = sum(1 for v in non_null if isinstance(v, str) and str(v).strip())
        ratio = str_cnt / len(non_null)
        if ratio >= 0.5 and (str_cnt, len(non_null)) > best_score:
            best_idx = idx
            best_score = (str_cnt, len(non_null))
    return best_idx


def make_unique_columns(columns):
    """정규화 후 같은 이름이 된 컬럼에 _2, _3 ...을 붙인다."""
    result = []
    counts = {}
    for i, column in enumerate(columns):
        base = sanitize_identifier(column) if str(column).strip() else f"COL_{i + 1}"
        base = base or f"COL_{i + 1}"
        counts[base] = counts.get(base, 0) + 1
        result.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
    return result


def read_sheet(excel_path: str, sheet_name: str, header_cfg, auto_detect: bool,
               source=None, use_desc_row: bool = False):
    """header 설정에 따라 시트를 DataFrame으로 읽는다.
    use_desc_row=True 이면 header 바로 다음 행을 컬럼 description으로 추출하고
    데이터에서 제외한다.
    반환: (df, header_row, 원본 컬럼명 리스트, auto_detect 사용 여부, comments)"""
    auto_used = False
    if source and source.get("kind") == "xlwings":
        sheet = source["handle"].sheets[sheet_name]
        values = sheet.used_range.value
        if values is None:
            raw = pd.DataFrame()
        elif not isinstance(values, list):
            raw = pd.DataFrame([[values]])
        elif values and not isinstance(values[0], list):
            raw = pd.DataFrame([values])
        else:
            raw = pd.DataFrame(values)

        if header_cfg is not None and str(header_cfg).strip() != "":
            header_row = int(header_cfg)
        elif auto_detect:
            header_row = detect_header_row(raw.head(20))
            auto_used = True
        else:
            header_row = 0

        if raw.empty or header_row >= len(raw):
            df = pd.DataFrame()
        else:
            headers = raw.iloc[header_row].tolist()
            df = raw.iloc[header_row + 1:].copy()
            df.columns = headers
            df = df.infer_objects()
    else:
        engine = source["engine"] if source else get_excel_engine(excel_path)
        if header_cfg is not None and str(header_cfg).strip() != "":
            header_row = int(header_cfg)
        elif auto_detect:
            raw = pd.read_excel(
                excel_path, sheet_name=sheet_name, header=None, nrows=20, engine=engine
            )
            header_row = detect_header_row(raw)
            auto_used = True
        else:
            header_row = 0
        df = pd.read_excel(
            excel_path, sheet_name=sheet_name, header=header_row, engine=engine
        )
    # 컬럼명 정규화 (description 추출 전에 유일한 이름으로 매핑)
    orig_all = [str(c) for c in df.columns]
    df.columns = make_unique_columns(df.columns)

    # header 바로 다음 행 = 컬럼 description -> comment 추출 후 데이터에서 제외
    comments = {}
    if use_desc_row and len(df) > 0:
        desc = df.iloc[0]
        # description 문자열 때문에 object로 읽힌 숫자/날짜 컬럼 타입 복원
        df = df.iloc[1:].infer_objects()
        comments = {
            c: str(desc[c]).strip()
            for c in df.columns
            if pd.notna(desc[c]) and str(desc[c]).strip()
        }

    # 빈 컬럼/빈 행 제거
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df = df.reset_index(drop=True)
    name_map = dict(zip(make_unique_columns(orig_all), orig_all))
    orig_cols = [name_map.get(c, str(c)) for c in df.columns]
    comments = {c: v for c, v in comments.items() if c in df.columns}
    return df, header_row, orig_cols, auto_used, comments


# ---------------------------------------------------------------------------
# 타입 추론 / 파싱 / 검증
# ---------------------------------------------------------------------------

HANA_TYPE_CHOICES = [
    "INTEGER",
    "BIGINT",
    "SMALLINT",
    "DECIMAL(18,4)",
    "DECIMAL(38,10)",
    "DOUBLE",
    "REAL",
    "BOOLEAN",
    "DATE",
    "TIME",
    "SECONDDATE",
    "TIMESTAMP",
    "NVARCHAR(20)",
    "NVARCHAR(50)",
    "NVARCHAR(100)",
    "NVARCHAR(200)",
    "NVARCHAR(500)",
    "NVARCHAR(1000)",
    "NVARCHAR(2000)",
    "NVARCHAR(5000)",
    "NCLOB",
]

_TYPE_RE = re.compile(
    r"^\s*([A-Za-z][A-Za-z0-9 ]*?)\s*(?:\(\s*(\d+)\s*(?:,\s*(\d+)\s*)?\))?\s*$"
)

_INT_RANGES = {
    "TINYINT": (0, 255),
    "SMALLINT": (-32768, 32767),
    "INTEGER": (-2147483648, 2147483647),
    "INT": (-2147483648, 2147483647),
    "BIGINT": (-9223372036854775808, 9223372036854775807),
}

_LENGTH_TYPES = {"NVARCHAR", "VARCHAR", "CHAR", "NCHAR", "ALPHANUM"}
_TEXT_TYPES = {"NCLOB", "CLOB", "TEXT", "SHORTTEXT"}
_FLOAT_TYPES = {"DOUBLE", "REAL", "FLOAT"}
_DECIMAL_TYPES = {"DECIMAL", "SMALLDECIMAL"}
_DATE_TYPES = {"DATE", "TIME", "TIMESTAMP", "SECONDDATE"}

KNOWN_TYPE_BASES = (set(_INT_RANGES) | _LENGTH_TYPES | _TEXT_TYPES | _FLOAT_TYPES
                    | _DECIMAL_TYPES | _DATE_TYPES | {"BOOLEAN"})


def parse_hana_type(type_str: str):
    """'DECIMAL(18,4)' -> ('DECIMAL', 18, 4). 파싱 실패 시 ValueError."""
    m = _TYPE_RE.match(str(type_str or ""))
    if not m:
        raise ValueError(f"데이터 유형 표기를 해석할 수 없습니다: {type_str!r}")
    base = m.group(1).strip().upper().replace(" ", "")
    p1 = int(m.group(2)) if m.group(2) is not None else None
    p2 = int(m.group(3)) if m.group(3) is not None else None
    return base, p1, p2


def normalize_hana_type(type_str: str) -> str:
    """사용자가 입력한 타입 문자열을 검증하고 표준 표기로 되돌린다."""
    base, p1, p2 = parse_hana_type(type_str)
    if base not in KNOWN_TYPE_BASES:
        raise ValueError(f"지원하지 않는 데이터 유형입니다: {type_str}")
    if base in _LENGTH_TYPES:
        if p1 is None:
            p1 = 255
        if not 1 <= p1 <= 5000:
            raise ValueError(f"{base} 길이는 1~5000 범위여야 합니다: {type_str}")
        return f"{base}({p1})"
    if base in _DECIMAL_TYPES:
        if p1 is None:
            return base
        p2 = 0 if p2 is None else p2
        if not 1 <= p1 <= 38 or not 0 <= p2 <= p1:
            raise ValueError(
                f"DECIMAL(precision, scale)은 1<=p<=38, 0<=s<=p 이어야 합니다: {type_str}")
        return f"{base}({p1},{p2})"
    return base


def infer_hana_type(series: pd.Series) -> str:
    """pandas Series -> HANA 컬럼 타입 문자열."""
    s = series.dropna()
    dtype = series.dtype

    # 정수
    if pd.api.types.is_integer_dtype(dtype):
        if len(s) and (s.max() > 2147483647 or s.min() < -2147483648):
            return "BIGINT"
        return "INTEGER"

    # 실수 -> DECIMAL(precision, scale) 산정
    if pd.api.types.is_float_dtype(dtype):
        if len(s) == 0:
            return "DECIMAL(18,4)"
        # float지만 전부 정수인 경우
        if all(float(v).is_integer() for v in s if not math.isinf(float(v))):
            mx = max(abs(v) for v in s) if len(s) else 0
            return "BIGINT" if mx > 2147483647 else "INTEGER"
        int_digits, scale = 1, 1
        for v in s:
            txt = f"{v:.10f}".rstrip("0")
            ip, _, fp = txt.partition(".")
            int_digits = max(int_digits, len(ip.lstrip("-")))
            scale = max(scale, len(fp))
        scale = min(scale, 10)
        precision = min(int_digits + scale, 38)
        return f"DECIMAL({precision},{scale})"

    # bool
    if pd.api.types.is_bool_dtype(dtype):
        return "BOOLEAN"

    # 날짜/시간
    if pd.api.types.is_datetime64_any_dtype(dtype):
        if len(s) and all(
            (t.hour, t.minute, t.second, t.microsecond) == (0, 0, 0, 0) for t in s
        ):
            return "DATE"
        return "TIMESTAMP"

    # 문자열
    max_len = 10
    for v in s:
        max_len = max(max_len, len(str(v)))
    # 여유분 반영 후 단계값으로 반올림
    for cap in (20, 50, 100, 200, 500, 1000, 2000, 5000):
        if max_len <= cap * 0.8:
            return f"NVARCHAR({cap})"
    return "NCLOB"


def build_column_defs(df: pd.DataFrame):
    return [(col, infer_hana_type(df[col])) for col in df.columns]


_TRUE_WORDS = {"true", "t", "yes", "y", "1", "참", "예", "o"}
_FALSE_WORDS = {"false", "f", "no", "n", "0", "거짓", "아니오", "x"}


def _is_null(value):
    if value is None:
        return True
    try:
        if isinstance(value, (list, dict, tuple, set)):
            return False
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def coerce_value(value, base, p1, p2):
    """단일 값을 지정 타입으로 변환. 반환 (ok, converted)."""
    if _is_null(value):
        return True, None

    if isinstance(value, str):
        text = value.strip()
    elif isinstance(value, float) and float(value).is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()
    if text == "":
        return True, None

    try:
        if base == "BOOLEAN":
            if isinstance(value, bool):
                return True, value
            low = text.lower()
            if low in _TRUE_WORDS:
                return True, True
            if low in _FALSE_WORDS:
                return True, False
            return False, None

        if base in _INT_RANGES:
            if isinstance(value, bool):
                return True, int(value)
            num = decimal.Decimal(text.replace(",", ""))
            if num != num.to_integral_value():
                return False, None
            ival = int(num)
            lo, hi = _INT_RANGES[base]
            if not lo <= ival <= hi:
                return False, None
            return True, ival

        if base in _FLOAT_TYPES:
            fval = float(text.replace(",", ""))
            if math.isnan(fval) or math.isinf(fval):
                return False, None
            return True, fval

        if base in _DECIMAL_TYPES:
            num = decimal.Decimal(text.replace(",", ""))
            if num.is_nan() or num.is_infinite():
                return False, None
            if p1 is not None:
                scale = p2 or 0
                num = num.quantize(decimal.Decimal(1).scaleb(-scale),
                                   rounding=decimal.ROUND_HALF_UP)
                tup = num.as_tuple()
                int_digits = len(tup.digits) + tup.exponent
                if int_digits > (p1 - scale):
                    return False, None
            return True, float(num)

        if base in _DATE_TYPES:
            if isinstance(value, (datetime.datetime, datetime.date, pd.Timestamp)):
                ts = pd.Timestamp(value)
            else:
                ts = pd.to_datetime(text, errors="raise")
            if pd.isna(ts):
                return False, None
            py = ts.to_pydatetime()
            if base == "DATE":
                return True, py.date()
            if base == "TIME":
                return True, py.time()
            if base == "SECONDDATE":
                return True, py.replace(microsecond=0)
            return True, py

        if base in _LENGTH_TYPES:
            if isinstance(value, (datetime.datetime, pd.Timestamp)):
                text = pd.Timestamp(value).isoformat(sep=" ")
            if p1 is not None and len(text) > p1:
                return False, None
            return True, text

        if base in _TEXT_TYPES:
            return True, text

    except (ValueError, TypeError, ArithmeticError, decimal.InvalidOperation,
            OverflowError, pd.errors.OutOfBoundsDatetime):
        return False, None

    return False, None


def check_series(series: pd.Series, type_str: str, sample_limit: int = 5) -> dict:
    """지정 타입으로 변환되지 않는 값들을 찾는다.

    반환: {ok, total, null_count, bad_count, samples: [(행번호, 원본값)], error}
    """
    result = {
        "ok": True, "total": int(len(series)), "null_count": 0,
        "bad_count": 0, "samples": [], "error": None,
    }
    try:
        base, p1, p2 = parse_hana_type(normalize_hana_type(type_str))
    except ValueError as e:
        result["ok"] = False
        result["error"] = str(e)
        return result

    for pos, value in enumerate(series.tolist()):
        ok, converted = coerce_value(value, base, p1, p2)
        if not ok:
            result["bad_count"] += 1
            if len(result["samples"]) < sample_limit:
                result["samples"].append((pos + 1, value))
        elif converted is None:
            result["null_count"] += 1
    result["ok"] = result["bad_count"] == 0 and result["error"] is None
    return result


# 불일치 처리 정책
POLICY_NULL = "null"
POLICY_SKIP_ROW = "skip_row"
POLICY_ABORT = "abort"
POLICY_REVERT = "revert"

POLICY_ORDER = [POLICY_NULL, POLICY_SKIP_ROW, POLICY_REVERT, POLICY_ABORT]

POLICY_LABELS = {
    POLICY_NULL: "NULL로 변환 후 진행",
    POLICY_SKIP_ROW: "해당 행 제외",
    POLICY_REVERT: "원래(추론) 유형으로 되돌리기",
    POLICY_ABORT: "업로드 중단",
}

POLICY_DESCRIPTIONS = {
    POLICY_NULL: "변환에 실패한 셀만 NULL(빈 값)로 넣고 나머지는 정상 적재합니다.",
    POLICY_SKIP_ROW: "변환에 실패한 셀이 있는 행 전체를 업로드에서 제외합니다.",
    POLICY_REVERT: "자동 추론된 안전한 유형으로 되돌려 데이터 손실 없이 적재합니다.",
    POLICY_ABORT: "불일치가 하나라도 있으면 해당 시트 업로드를 중단하고 오류로 처리합니다.",
}


def apply_column_types(df: pd.DataFrame, columns: list, logger=None):
    """컬럼 계획(columns)에 따라 df를 지정 타입으로 변환한다.

    columns: [{'name', 'target_type', 'inferred_type', 'policy', 'include'}, ...]
    반환: (변환된 df, 실제 사용된 [(컬럼, 타입)] 목록, report dict)
    """
    report = {"changed": [], "nulled": {}, "dropped_rows": 0, "reverted": []}
    used = [c for c in columns if c.get("include", True) and c["name"] in df.columns]
    out = pd.DataFrame(index=df.index)
    drop_rows = set()
    col_defs = []

    for col in used:
        name = col["name"]
        inferred = normalize_hana_type(col.get("inferred_type")
                                       or col.get("target_type"))
        target = normalize_hana_type(col.get("target_type") or inferred)
        policy = col.get("policy") or POLICY_NULL

        if target == inferred:
            # 유형 변경 없음 -> 원본 유지
            out[name] = df[name]
            col_defs.append((name, target))
            continue

        base, p1, p2 = parse_hana_type(target)
        values, bad_positions = [], []
        for pos, value in enumerate(df[name].tolist()):
            ok, converted = coerce_value(value, base, p1, p2)
            values.append(converted)
            if not ok:
                bad_positions.append(pos)

        if bad_positions and policy == POLICY_ABORT:
            raise TypeMismatchAbort(
                f"컬럼 '{name}'을(를) {target}(으)로 변환할 수 없는 값이 "
                f"{len(bad_positions)}건 있습니다. (처리 방식: 업로드 중단)"
            )
        if bad_positions and policy == POLICY_REVERT:
            out[name] = df[name]
            col_defs.append((name, inferred))
            report["reverted"].append((name, target, inferred, len(bad_positions)))
            if logger:
                logger.warning("   - 컬럼 '%s': %s 변환 실패 %d건 -> %s(으)로 되돌림",
                               name, target, len(bad_positions), inferred)
            continue

        out[name] = pd.Series(values, index=df.index, dtype="object")
        col_defs.append((name, target))
        report["changed"].append((name, inferred, target))
        if bad_positions:
            if policy == POLICY_SKIP_ROW:
                drop_rows.update(df.index[p] for p in bad_positions)
                if logger:
                    logger.warning("   - 컬럼 '%s': %s 변환 실패 %d건 -> 해당 행 제외",
                                   name, target, len(bad_positions))
            else:
                report["nulled"][name] = len(bad_positions)
                if logger:
                    logger.warning("   - 컬럼 '%s': %s 변환 실패 %d건 -> NULL 처리",
                                   name, target, len(bad_positions))

    if drop_rows:
        out = out.drop(index=list(drop_rows))
        report["dropped_rows"] = len(drop_rows)

    return out, col_defs, report


# ---------------------------------------------------------------------------
# DB 작업
# ---------------------------------------------------------------------------

def connect_hana(cfg, logger):
    try:
        from hdbcli import dbapi
    except ImportError:
        raise RuntimeError("hdbcli 미설치. 'pip install hdbcli' 후 다시 실행하세요.")

    host = cfg.get("HANA", "host", fallback="").strip()
    port = cfg.getint("HANA", "port", fallback=443)
    user = cfg.get("HANA", "user", fallback="").strip()
    if not host or not user:
        raise RuntimeError(
            "HANA 접속 정보(호스트/사용자)가 비어 있습니다. [HANA 접속] 탭을 확인하세요.")
    logger.info("[HANA] 접속 시도: %s:%s, USER=%s", host, port, user)

    conn_args = dict(
        address=host,
        port=port,
        user=user,
        password=cfg.get("HANA", "password", fallback=""),
        encrypt=True,
        sslValidateCertificate=False,
    )
    # 사내 프록시 경유가 필요한 경우 [HANA] proxy_host/proxy_port 설정
    proxy_host = cfg.get("HANA", "proxy_host", fallback="").strip()
    if proxy_host:
        conn_args["proxy_host"] = proxy_host
        conn_args["proxy_port"] = cfg.getint("HANA", "proxy_port", fallback=8080)
        conn_args["proxy_http"] = True
        logger.info("   - 프록시 사용: %s:%s",
                    conn_args["proxy_host"], conn_args["proxy_port"])

    try:
        conn = dbapi.connect(**conn_args)
    except dbapi.Error as e:
        logger.error("   - 접속 실패: %s", e)
        logger.error("   확인 사항:")
        logger.error("     1) 사내 방화벽/프록시에서 %s:%s (HTTPS) 아웃바운드 허용 여부",
                     host, port)
        logger.error("        -> 프록시 필요 시 [HANA 접속] 탭에 프록시 호스트/포트 설정")
        logger.error("     2) Datasphere IP Allowlist에 현재 공인 IP 등록 여부"
                     " (시스템 > 구성 > IP 허용 목록)")
        logger.error("     3) 호스트/포트/사용자/비밀번호 값 확인")
        raise RuntimeError(f"HANA 접속 실패: {e}") from e
    logger.info("   - 접속 성공")
    return conn


def set_ddl_transactional(conn, logger) -> bool:
    """DDL(TRUNCATE/CREATE TABLE)을 트랜잭션에 포함시킨다.

    HANA에서 TRUNCATE는 DML이 아니라 DDL이다. 세션의 AUTOCOMMIT DDL이 ON이면
    TRUNCATE가 즉시 커밋되므로, 그 뒤 INSERT가 실패해 rollback을 해도
    "기존 데이터는 이미 삭제되고 새 데이터도 없는" 상태가 되어 데이터가 유실된다.
    conn.setautocommit(False)는 DML에만 적용되고 DDL에는 영향을 주지 않는다.

    설정에 실패해도 적재는 그대로 진행한다(TRUNCATE 사용). 다만 그 경우
    업로드가 중간에 실패하면 기존 데이터를 되돌릴 수 없으므로 경고를 남긴다.
    반환: True면 TRUNCATE 롤백 가능.
    """
    try:
        cursor = conn.cursor()
        try:
            cursor.execute("SET TRANSACTION AUTOCOMMIT DDL OFF")
        finally:
            cursor.close()
    except Exception as e:
        logger.warning("[HANA] AUTOCOMMIT DDL OFF 설정 실패: %s", e)
        logger.warning("        -> TRUNCATE가 즉시 커밋되므로, 업로드가 중간에 "
                       "실패하면 기존 데이터를 복구할 수 없습니다.")
        return False
    logger.info("[HANA] AUTOCOMMIT DDL OFF 설정 (TRUNCATE 롤백 가능)")
    return True


def table_exists(cursor, schema: str, table: str) -> bool:
    cursor.execute(
        "SELECT COUNT(*) FROM SYS.TABLES WHERE SCHEMA_NAME = ? AND TABLE_NAME = ?",
        (schema, table),
    )
    return cursor.fetchone()[0] > 0


def create_table(cursor, schema: str, table: str, col_defs, logger, comments=None,
                 keys=None, not_nulls=None):
    """테이블을 생성한다.

    keys      : PRIMARY KEY로 묶을 컬럼 이름 목록 (순서 유지)
    not_nulls : NOT NULL로 만들 컬럼 이름 집합
                (PRIMARY KEY 컬럼은 HANA에서 암묵적으로 NOT NULL이지만
                 DDL에 명시해 의도를 분명히 남긴다)
    """
    comments = comments or {}
    keys = list(keys or [])
    not_nulls = set(not_nulls or []) | set(keys)
    parts = []
    for c, t in col_defs:
        d = f'"{c}" {t}'
        if c in not_nulls:
            d += " NOT NULL"
        cm = comments.get(c)
        if cm:
            d += " COMMENT '{}'".format(str(cm).replace("'", "''"))
        parts.append(d)
    if keys:
        parts.append("PRIMARY KEY ({})".format(
            ", ".join(f'"{k}"' for k in keys)))
    ddl = f'CREATE COLUMN TABLE "{schema}"."{table}" ({", ".join(parts)})'
    logger.info("[DDL] CREATE TABLE 실행: %s", ddl)
    cursor.execute(ddl)


def get_existing_columns(cursor, schema: str, table: str):
    cursor.execute(
        "SELECT COLUMN_NAME FROM SYS.TABLE_COLUMNS "
        "WHERE SCHEMA_NAME = ? AND TABLE_NAME = ? ORDER BY POSITION",
        (schema, table),
    )
    return [r[0] for r in cursor.fetchall()]


def format_db_type(data_type_name, length, scale) -> str:
    """SYS.TABLE_COLUMNS 값을 'NVARCHAR(50)' 형태의 표기로 만든다."""
    base = str(data_type_name or "").strip().upper()
    if base in _LENGTH_TYPES and length:
        return f"{base}({int(length)})"
    if base in _DECIMAL_TYPES and length:
        return f"{base}({int(length)},{int(scale or 0)})"
    return base


def get_existing_column_types(cursor, schema: str, table: str):
    """[(컬럼명, 타입표기)] 를 테이블 정의 순서대로 반환."""
    cursor.execute(
        "SELECT COLUMN_NAME, DATA_TYPE_NAME, LENGTH, SCALE "
        "FROM SYS.TABLE_COLUMNS "
        "WHERE SCHEMA_NAME = ? AND TABLE_NAME = ? ORDER BY POSITION",
        (schema, table),
    )
    return [(r[0], format_db_type(r[1], r[2], r[3])) for r in cursor.fetchall()]


def get_primary_key_columns(cursor, schema: str, table: str):
    """기존 테이블의 PRIMARY KEY 컬럼 목록."""
    cursor.execute(
        "SELECT COLUMN_NAME FROM SYS.CONSTRAINTS "
        "WHERE SCHEMA_NAME = ? AND TABLE_NAME = ? AND IS_PRIMARY_KEY = 'TRUE' "
        "ORDER BY POSITION",
        (schema, table),
    )
    return [r[0] for r in cursor.fetchall()]


def check_constraints(df: pd.DataFrame, columns, sample_limit: int = 5):
    """[KEY] / [NOT NULL] 지정이 실제 데이터와 맞는지 검사한다.

    NOT NULL 컬럼에 빈 값이 있거나, KEY 컬럼 조합에 빈 값/중복이 있으면
    CREATE TABLE은 되더라도 INSERT가 실패한다. 그래서 업로드 전에 잡아낸다.

    반환: [{'kind', 'columns', 'count', 'rows', 'detail'}]
      kind: 'not_null' | 'key_null' | 'key_duplicate'
    """
    issues = []
    used = [c for c in columns
            if c.get("include", True) and c["name"] in df.columns]
    keys = [c["name"] for c in used if c.get("is_key")]

    for col in used:
        name = col["name"]
        if not (col.get("not_null") or col.get("is_key")):
            continue
        mask = df[name].isna()
        count = int(mask.sum())
        if not count:
            continue
        rows = [int(i) + 1 for i in df.index[mask][:sample_limit]]
        issues.append({
            "kind": "key_null" if col.get("is_key") else "not_null",
            "columns": [name],
            "count": count,
            "rows": rows,
            "detail": (f"{'KEY' if col.get('is_key') else 'NOT NULL'} 컬럼 "
                       f"'{name}'에 빈 값 {count}건"),
        })

    if keys:
        sub = df[keys]
        dup_mask = sub.duplicated(keep=False) & ~sub.isna().any(axis=1)
        count = int(dup_mask.sum())
        if count:
            rows = [int(i) + 1 for i in df.index[dup_mask][:sample_limit]]
            examples = []
            for i in df.index[dup_mask][:sample_limit]:
                examples.append(" / ".join(str(sub.at[i, k]) for k in keys))
            issues.append({
                "kind": "key_duplicate",
                "columns": list(keys),
                "count": count,
                "rows": rows,
                "detail": ("KEY({}) 값이 중복된 행 {}건 — 예: {}".format(
                    ", ".join(keys), count, "; ".join(examples))),
            })
    return issues


def planned_column_types(plan):
    """업로드에 사용될 (컬럼명, 타입) 목록. 실제 변환 없이 계획만 계산한다."""
    result = []
    for col in plan["columns"]:
        if not col.get("include", True):
            continue
        target = col.get("target_type") or col.get("inferred_type")
        try:
            target = normalize_hana_type(target)
        except ValueError:
            target = str(target or "").strip().upper()
        result.append((col["name"], target))
    return result


def check_target_tables(plans, cfg, options, logger, cancelled=None) -> dict:
    """업로드 전에 대상 테이블의 기존 구조를 확인한다(메타데이터만 조회).

    기존 테이블은 CREATE를 다시 하지 않으므로, 미리보기에서 바꾼 데이터 유형이
    실제로는 적용되지 않는다. 그 사실을 업로드 전에 사용자에게 알리기 위한 검사다.

    [KEY]/[NOT NULL] 체크도 CREATE TABLE 시점에만 반영되므로, 기존 테이블에
    대해서는 적용되지 않는다는 사실을 함께 보고한다.

    반환: {'new': [...], 'missing_columns': [...], 'type_mismatch': [...],
           'constraints_ignored': [...], 'no_primary_key': [...],
           'error': str|None}
    """
    report = {"new": [], "missing_columns": [], "type_mismatch": [],
              "constraints_ignored": [], "no_primary_key": [], "error": None}
    active = [p for p in plans if p.get("enabled", True)]
    if not active:
        return report

    schema = (options.get("schema") or "").strip()
    if not schema:
        report["error"] = "HANA 스키마가 비어 있습니다."
        return report

    conn = cursor = None
    try:
        conn = connect_hana(cfg, logger)
        cursor = conn.cursor()
        for plan in active:
            if cancelled and cancelled():
                break
            table = plan["table_name"]
            planned = planned_column_types(plan)
            used = [c for c in plan["columns"] if c.get("include", True)]
            wanted_keys = [c["name"] for c in used if c.get("is_key")]
            wanted_nn = [c["name"] for c in used
                         if c.get("not_null") and not c.get("is_key")]

            if not table_exists(cursor, schema, table):
                report["new"].append((plan["sheet_name"], table, len(planned)))
                if options.get("load_mode") == LOAD_MODE_UPSERT and not wanted_keys:
                    report["no_primary_key"].append(
                        (plan["sheet_name"], table,
                         "새로 만들 테이블에 [KEY] 지정이 없어 PRIMARY KEY가 "
                         "생성되지 않습니다. UPSERT를 쓸 수 없습니다."))
                continue

            actual = dict((n.upper(), t)
                          for n, t in get_existing_column_types(cursor, schema, table))
            missing = [n for n, _ in planned if n.upper() not in actual]
            if missing:
                report["missing_columns"].append(
                    (plan["sheet_name"], table, missing))
            for name, want in planned:
                have = actual.get(name.upper())
                if have and want and have != want:
                    report["type_mismatch"].append(
                        (plan["sheet_name"], table, name, want, have))

            # 기존 테이블은 CREATE를 다시 하지 않으므로 KEY/NOT NULL이 반영되지 않는다.
            existing_pk = get_primary_key_columns(cursor, schema, table)
            if wanted_keys or wanted_nn:
                pk_upper = [k.upper() for k in existing_pk]
                same_pk = pk_upper == [k.upper() for k in wanted_keys]
                details = []
                if wanted_keys and not same_pk:
                    details.append(
                        "[KEY] 지정: {} / 테이블의 실제 PRIMARY KEY: {}".format(
                            ", ".join(wanted_keys),
                            ", ".join(existing_pk) if existing_pk else "없음"))
                if wanted_nn:
                    details.append("[NOT NULL] 지정: " + ", ".join(wanted_nn))
                if details:
                    report["constraints_ignored"].append(
                        (plan["sheet_name"], table, details))

            if options.get("load_mode") == LOAD_MODE_UPSERT and not existing_pk:
                report["no_primary_key"].append(
                    (plan["sheet_name"], table,
                     "테이블에 PRIMARY KEY가 없어 UPSERT를 실행할 수 없습니다."))
    except Exception as e:
        report["error"] = str(e)
        logger.warning("테이블 사전 확인 실패: %s", e)
    finally:
        if cursor:
            try:
                cursor.close()
            except Exception:
                pass
        if conn:
            try:
                conn.close()
            except Exception:
                pass
    return report


def to_db_value(v):
    if _is_null(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, decimal.Decimal):
        return float(v)
    return v


def insert_data(cursor, schema: str, table: str, df: pd.DataFrame,
                load_mode: str, logger, batch_size: int = 1000,
                progress=None, ddl_transactional: bool = False) -> int:
    """load_mode에 따라 데이터를 적재한다.

    truncate : TRUNCATE로 비운 뒤 INSERT
    insert   : 기존 데이터 유지하고 INSERT
    upsert   : UPSERT ... WITH PRIMARY KEY (테이블에 PK가 있어야 한다)
    """
    if load_mode == LOAD_MODE_TRUNCATE:
        note = "롤백 가능" if ddl_transactional else "즉시 커밋됨"
        logger.info("[DML] TRUNCATE 실행 (%s)", note)
        cursor.execute(f'TRUNCATE TABLE "{schema}"."{table}"')

    cols = list(df.columns)
    col_sql = ", ".join(f'"{c}"' for c in cols)
    placeholders = ", ".join("?" for _ in cols)

    if load_mode == LOAD_MODE_UPSERT:
        sql = (f'UPSERT "{schema}"."{table}" ({col_sql}) '
               f'VALUES ({placeholders}) WITH PRIMARY KEY')
        verb = "UPSERT"
    else:
        sql = (f'INSERT INTO "{schema}"."{table}" ({col_sql}) '
               f'VALUES ({placeholders})')
        verb = "INSERT"

    rows = [tuple(to_db_value(v) for v in rec)
            for rec in df.itertuples(index=False, name=None)]
    total = 0
    for i in range(0, len(rows), batch_size):
        batch = rows[i:i + batch_size]
        cursor.executemany(sql, batch)
        total += len(batch)
        if progress:
            progress(total, len(rows))
    logger.info("[DML] %s 완료 (행=%d)", verb, total)
    return total


# ---------------------------------------------------------------------------
# 시트 계획 수집 (미리보기 / 실행 공통)
# ---------------------------------------------------------------------------

def collect_sheet_plans(file_paths, options, logger, allow_drm=True,
                        progress=None, cancelled=None):
    """대상 파일들을 열어 시트별 계획(DataFrame 포함)을 만든다.

    allow_drm=False 이면 DRM 파일에서 DrmProtectedError를 던진다.
    반환: (plans, errors)
      plan = {file_path, file_name, sheet_name, table_name, header_row,
              auto_header_used, row_count, columns:[...], df, enabled, is_drm}
      column = {name, orig_name, inferred_type, target_type, comment,
                policy, include, null_count}
    """
    plans, errors = [], []
    keyword = options["keyword"]

    for fpath in file_paths:
        if cancelled and cancelled():
            break
        fname = os.path.basename(fpath)
        file_stem = sanitize_identifier(os.path.splitext(fname)[0])
        if progress:
            progress(f"파일 열기: {fname}")
        source = None
        try:
            source = open_excel_source(fpath, logger, options["xlwings_visible"],
                                       allow_drm=allow_drm)
        except DrmProtectedError:
            raise
        except Exception as e:
            logger.error("파일 열기 실패(%s): %s", fname, e)
            errors.append((fname, "", str(e)))
            continue

        try:
            sheet_names = source["sheet_names"]
            logger.info("[FILE] %s | 전체 시트: %s", fname, sheet_names)
            target_sheets = [s for s in sheet_names
                             if s.strip().upper().startswith(keyword.upper())]
            logger.info("   - 대상 시트: %s", target_sheets)
            if not target_sheets:
                errors.append((fname, "", f"'{keyword}'로 시작하는 시트가 없습니다."))
                continue

            for sheet in target_sheets:
                if cancelled and cancelled():
                    break
                if progress:
                    progress(f"{fname} / {sheet} 읽는 중...")
                try:
                    df, header_row, orig_cols, auto_used, comments = read_sheet(
                        fpath, sheet, options["header_cfg"], options["auto_header"],
                        source, options["use_desc_row"])
                except Exception as e:
                    logger.error("시트 '%s' 읽기 실패: %s", sheet, e)
                    errors.append((fname, sheet, str(e)))
                    continue

                table = sheet_to_table_name(sheet, keyword)
                if options["use_prefix"]:
                    table = f"{file_stem}_{table}"

                columns = []
                for name, orig in zip(df.columns, orig_cols):
                    inferred = infer_hana_type(df[name])
                    columns.append({
                        "name": name,
                        "orig_name": orig,
                        "inferred_type": inferred,
                        "target_type": inferred,
                        "comment": comments.get(name, ""),
                        "policy": POLICY_NULL,
                        "include": True,
                        "is_key": False,
                        "not_null": False,
                        "null_count": int(df[name].isna().sum()),
                    })

                plans.append({
                    "file_path": fpath,
                    "file_name": fname,
                    "sheet_name": sheet,
                    "table_name": table,
                    "header_row": header_row,
                    "auto_header_used": auto_used,
                    "row_count": int(len(df)),
                    "columns": columns,
                    "df": df,
                    "enabled": True,
                    "is_drm": source.get("kind") == "xlwings",
                    "checked": False,
                })
                logger.info("   - '%s' -> 테이블 '%s' (행=%d, 컬럼=%d)",
                            sheet, table, len(df), len(df.columns))
        finally:
            close_excel_source(source, logger)

    return plans, errors


# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------

def execute_plans(plans, cfg, options, logger, dry_run=False,
                  progress=None, cancelled=None):
    """시트 계획들을 실제로 업로드한다. 반환: (summary, had_error)

    summary = [(파일, 시트, 테이블, 건수, 상태)]
    """
    summary = []
    had_error = False
    active = [p for p in plans if p.get("enabled", True)]
    if not active:
        logger.warning("업로드할 시트가 없습니다.")
        return summary, False

    schema = (options.get("schema") or "").strip()
    if not dry_run and not schema:
        raise RuntimeError("HANA 스키마가 비어 있습니다. [HANA 접속] 탭에서 스키마를 입력하세요.")

    load_mode = options.get("load_mode", LOAD_MODE_TRUNCATE)
    if load_mode not in LOAD_MODES:
        load_mode = LOAD_MODE_TRUNCATE
    logger.info("적재 모드: %s", LOAD_MODE_LABELS[load_mode])

    conn = cursor = None
    ddl_tx = False
    if not dry_run:
        conn = connect_hana(cfg, logger)
        conn.setautocommit(False)
        # TRUNCATE(DDL)까지 트랜잭션에 넣어 업로드 실패 시 기존 데이터를 지킨다.
        ddl_tx = set_ddl_transactional(conn, logger)
        cursor = conn.cursor()

    file_status = {}
    try:
        for idx, plan in enumerate(active, 1):
            if cancelled and cancelled():
                logger.warning("사용자 요청으로 중단되었습니다.")
                break
            fname = plan["file_name"]
            sheet = plan["sheet_name"]
            table = plan["table_name"]
            file_status.setdefault(fname, {"ok": True, "any": False,
                                           "path": plan["file_path"]})
            if progress:
                progress(idx, len(active), f"{fname} / {sheet} -> {table}")

            logger.info("-" * 50)
            logger.info("[%d/%d] 시트 '%s' -> 테이블 '%s'", idx, len(active), sheet, table)

            try:
                df = plan.get("df")
                if df is None or df.empty:
                    logger.warning("   - 데이터 없음 -> 건너뜀")
                    summary.append((fname, sheet, table, 0, "데이터없음"))
                    continue

                df2, col_defs, report = apply_column_types(df, plan["columns"], logger)
                comments = {c["name"]: c["comment"] for c in plan["columns"]
                            if c.get("include", True) and c.get("comment")}

                if report["changed"]:
                    logger.info("   - 유형 변경: %s",
                                ", ".join(f"{n}: {o} -> {t}"
                                          for n, o, t in report["changed"]))
                if report["dropped_rows"]:
                    logger.warning("   - 유형 불일치로 제외된 행: %d", report["dropped_rows"])

                if df2.empty or not len(df2.columns):
                    logger.warning("   - 유효한 데이터가 없습니다 -> 건너뜀")
                    summary.append((fname, sheet, table, 0, "데이터없음"))
                    continue

                used_cols = [c for c in plan["columns"] if c.get("include", True)]
                keys = [c["name"] for c in used_cols if c.get("is_key")]
                not_nulls = [c["name"] for c in used_cols if c.get("not_null")]

                if dry_run:
                    logger.info("[DRY-RUN] CREATE(미존재 시) + %s 예정 (%d건)",
                                load_mode.upper(), len(df2))
                    if keys:
                        logger.info("[DRY-RUN] PRIMARY KEY: %s", ", ".join(keys))
                    if not_nulls:
                        logger.info("[DRY-RUN] NOT NULL: %s", ", ".join(not_nulls))
                    for c, t in col_defs:
                        logger.debug("    %-30s %s", c, t)
                    summary.append((fname, sheet, table, len(df2), "DRY-RUN"))
                    file_status[fname]["any"] = True
                    continue

                logger.info("[DDL] 테이블 확인/생성: %s.%s", schema, table)
                if not table_exists(cursor, schema, table):
                    create_table(cursor, schema, table, col_defs, logger, comments,
                                 keys=keys, not_nulls=not_nulls)
                else:
                    actual = dict(
                        (c.upper(), t)
                        for c, t in get_existing_column_types(cursor, schema, table))
                    missing = [c for c in df2.columns if c.upper() not in actual]
                    if missing:
                        raise RuntimeError(
                            f"기존 테이블 {table}에 없는 컬럼 존재: {missing}. "
                            f"테이블 구조를 확인하세요.")
                    # 기존 테이블은 CREATE를 다시 하지 않으므로 지정 유형이 반영되지
                    # 않는다. 실제 테이블 유형과 다르면 알려 준다.
                    mismatch = [(c, t, actual[c.upper()]) for c, t in col_defs
                                if actual.get(c.upper())
                                and actual[c.upper()] != t]
                    if mismatch:
                        logger.warning(
                            "   - 기존 테이블 유형이 지정 유형과 다릅니다. "
                            "테이블의 기존 유형이 그대로 사용됩니다:")
                        for c, want, have in mismatch:
                            logger.warning("       %s : 지정 %s / 실제 %s",
                                           c, want, have)
                    if keys or not_nulls:
                        logger.warning(
                            "   - 기존 테이블이므로 [KEY]/[NOT NULL] 지정은 "
                            "반영되지 않습니다. 테이블의 기존 제약이 사용됩니다.")
                    # 기존 테이블 컬럼 순서에 맞춰 존재하는 컬럼만 적재
                    df2 = df2[[c for c in df2.columns if c.upper() in actual]]

                logger.info("[DML] 적재 (모드=%s, 행=%d, 컬럼=%d)",
                            load_mode, len(df2), len(df2.columns))
                cnt = insert_data(cursor, schema, table, df2, load_mode, logger,
                                  ddl_transactional=ddl_tx)
                conn.commit()

                status = "성공"
                if report["dropped_rows"]:
                    status += f" (행 {report['dropped_rows']}건 제외)"
                if report["nulled"]:
                    status += f" (NULL 처리 {sum(report['nulled'].values())}건)"
                summary.append((fname, sheet, table, cnt, status))
                file_status[fname]["any"] = True

            except Exception as e:
                rolled_back = False
                if conn:
                    try:
                        conn.rollback()
                        rolled_back = True
                    except Exception as re:
                        logger.error("rollback 실패: %s", re)
                logger.error("시트 '%s' 처리 실패: %s", sheet, e)
                if rolled_back and load_mode == LOAD_MODE_TRUNCATE:
                    if ddl_tx:
                        logger.info("   - rollback 완료: 테이블 '%s'의 기존 "
                                    "데이터는 그대로 유지됩니다.", table)
                    else:
                        logger.error("   - 주의: AUTOCOMMIT DDL을 끌 수 없어 "
                                     "TRUNCATE가 이미 커밋되었습니다. 테이블 "
                                     "'%s'는 비어 있을 수 있습니다.", table)
                summary.append((fname, sheet, table, 0, f"실패: {e}"))
                file_status[fname]["ok"] = False
                had_error = True
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

    # 파일 단위 archive 이동
    if options.get("move_processed") and not dry_run:
        for fname, st in file_status.items():
            if not (st["ok"] and st["any"]):
                continue
            fpath = st["path"]
            if not os.path.isfile(fpath):
                continue
            archive_name = options.get("archive_dir_name", "archive")
            archive_dir = archive_name if os.path.isabs(archive_name) \
                else os.path.join(os.path.dirname(fpath), archive_name)
            try:
                os.makedirs(archive_dir, exist_ok=True)
                dest = os.path.join(archive_dir, fname)
                if os.path.exists(dest):
                    stem, ext = os.path.splitext(fname)
                    dest = os.path.join(
                        archive_dir,
                        f"{stem}_{datetime.datetime.now():%Y%m%d_%H%M%S}{ext}")
                shutil.move(fpath, dest)
                logger.info("archive 이동: %s -> %s", fname, dest)
            except Exception as e:
                logger.warning("archive 이동 실패(%s): %s", fname, e)

    log_summary(summary, logger)
    return summary, had_error


def log_summary(summary, logger):
    logger.info("=" * 60)
    logger.info("처리 결과 요약")
    for fname, sheet, table, cnt, status in summary:
        logger.info("  %-28s | %-24s | %-24s | %8s | %s",
                    fname, sheet, table, f"{cnt}건", status)
    if not summary:
        logger.info("  처리된 시트가 없습니다.")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")


def list_excel_files(directory):
    return [
        os.path.join(directory, f)
        for f in sorted(os.listdir(directory))
        if f.lower().endswith(EXCEL_EXTENSIONS) and not f.startswith("~$")
    ]


def parse_csv_option(value):
    if not value:
        return None
    return {v.strip().upper() for v in value.split(",") if v.strip()}


def main():
    parser = argparse.ArgumentParser(description="Excel -> HANA Datasphere Loader")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parser.add_argument(
        "--config",
        default=os.path.join(script_dir, "config.ini"),
        help="설정 파일 경로 (기본값: 실행 스크립트와 같은 폴더의 config.ini)",
    )
    parser.add_argument("--file", action="append", default=None,
                        help="처리할 엑셀 파일 경로 (여러 번 지정 가능). "
                             "지정하지 않으면 [GENERAL] excel_dir 전체를 처리")
    parser.add_argument("--dry-run", action="store_true",
                        help="실제로 DB 작업을 수행하지 않고, 매핑/계획만 출력")
    parser.add_argument("--only-sheet", default=None,
                        help="처리할 시트 이름(들), 콤마로 구분 (예: TABLE_T_SALES, TEST_TABLE)")
    parser.add_argument("--only-table", default=None,
                        help="처리할 테이블 이름(들), 콤마로 구분 (예: T_SALES, T_CUSTOMER)")
    parser.add_argument("--allow-drm", action="store_true",
                        help="DRM 파일을 Excel(xlwings)로 열어 처리 (Windows + Excel 필요)")
    parser.add_argument("--mode", choices=LOAD_MODES, default=None,
                        help="적재 모드 (config의 load_mode를 덮어씀). "
                             "truncate=비우고 적재 / insert=이어서 추가 / "
                             "upsert=키 기준 갱신+추가")
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(args.config)) or os.getcwd()
    cfg = load_config(args.config)
    logger = setup_logging(base_dir, cfg)
    options = build_options(cfg, base_dir)
    if args.mode:
        options["load_mode"] = args.mode
        options["truncate"] = args.mode == LOAD_MODE_TRUNCATE

    logger.info("=" * 50)
    logger.info(" Excel → HANA Loader 시작")
    logger.info(" Config 파일: %s", os.path.abspath(args.config))
    logger.info(" Dry-run 모드: %s", args.dry_run)
    logger.info("=" * 50)

    if args.file:
        file_paths = [os.path.abspath(p) for p in args.file]
        missing = [p for p in file_paths if not os.path.isfile(p)]
        if missing:
            logger.error("파일을 찾을 수 없습니다: %s", missing)
            sys.exit(1)
    else:
        excel_dir = options["excel_dir"]
        if not excel_dir or not os.path.isdir(excel_dir):
            logger.error("excel_dir 디렉터리가 존재하지 않습니다: %s", excel_dir)
            sys.exit(1)
        file_paths = list_excel_files(excel_dir)
        logger.info("엑셀 폴더: %s", excel_dir)

    if not file_paths:
        logger.warning("처리할 Excel 파일이 없습니다.")
        return
    for p in file_paths:
        logger.info("   - %s", os.path.basename(p))

    only_sheets = parse_csv_option(args.only_sheet)
    only_tables = parse_csv_option(args.only_table)

    try:
        plans, errors = collect_sheet_plans(file_paths, options, logger,
                                            allow_drm=args.allow_drm)
    except DrmProtectedError as e:
        logger.error("%s", e)
        logger.error("CLI에서는 --allow-drm 옵션을, GUI에서는 [미리보기]를 사용하세요.")
        sys.exit(1)

    for fname, sheet, msg in errors:
        logger.warning("건너뜀: %s %s -> %s", fname, sheet or "-", msg)

    if only_sheets:
        plans = [p for p in plans if p["sheet_name"].strip().upper() in only_sheets]
    if only_tables:
        plans = [p for p in plans if p["table_name"].upper() in only_tables]

    summary, had_error = execute_plans(plans, cfg, options, logger,
                                       dry_run=args.dry_run)
    sys.exit(1 if (had_error or errors) else 0)


if __name__ == "__main__":
    main()
