# -*- coding: utf-8 -*-
"""테스트 공통 픽스처.

실제 HANA나 Excel, PyQt5 없이도 전체 로직을 검증할 수 있도록
샘플 워크북과 가짜 DB 커넥션을 제공한다.
"""

import os
import sys
import datetime

import pandas as pd
import pytest

# 저장소 루트를 import 경로에 넣는다 (excel_loader / wizard).
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

# PyQt5가 없는 환경에서도 GUI 로직을 테스트하기 위한 Qt 대역.
MOCKQT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mockqt")


# ---------------------------------------------------------------------------
# 샘플 엑셀
# ---------------------------------------------------------------------------

SHEET_MAIN = "TABLE_T_SALES"
SHEET_CODE = "TABLE_T_CODE"


@pytest.fixture(scope="session")
def sample_xlsx(tmp_path_factory):
    """제목 2줄 + 헤더 + description 행 + 데이터가 있는 실제 xlsx를 만든다.

    일부러 지저분한 값을 섞어 두었다.
      · 수량 '다섯'        -> 숫자로 변환 불가
      · 수량 9999999999    -> INTEGER 범위 초과, BIGINT는 가능
      · 고객명 매우 긴 문자열 -> NVARCHAR 길이 초과 테스트
      · 비고 빈 값          -> NOT NULL 위반 테스트
      · 주문일 문자열 날짜   -> DATE 변환 테스트
    """
    path = tmp_path_factory.mktemp("data") / "sample.xlsx"
    rows = [
        ["2026년 상반기 매출", None, None, None, None, None],
        [None, None, None, None, None, None],
        ["주문번호", "고객명", "수량", "단가", "주문일", "비고"],
        ["Order ID", "Customer", "Qty", "Unit price", "Order date", "Remark"],
        ["A001", "가나상사", 10, 1500.25, datetime.datetime(2026, 1, 5), "정상"],
        ["A002", "다라식품", 3, 99999.5, datetime.datetime(2026, 2, 11), "긴급"],
        ["A003", "테스트고객", "다섯", 120.0, datetime.datetime(2026, 3, 2), None],
        ["A004", "가" * 40, 7, 0.125, "2026-04-01", "긴 이름"],
        ["A005", "마바물산", 9999999999, 12.5, datetime.datetime(2026, 5, 20), "큰 수량"],
    ]
    code = [["코드", "이름"], ["Code", "Name"], ["C1", "알파"], ["C2", "베타"]]

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name=SHEET_MAIN,
                                    index=False, header=False)
        # keyword로 시작하지 않으므로 처리 대상이 아니어야 한다.
        pd.DataFrame({"X": [1, 2, 3]}).to_excel(writer, sheet_name="설명",
                                                index=False)
        pd.DataFrame(code).to_excel(writer, sheet_name=SHEET_CODE,
                                    index=False, header=False)
    return str(path)


@pytest.fixture
def drm_xlsx(tmp_path):
    """DRMONE 시그니처를 가진 가짜 DRM 파일."""
    path = tmp_path / "drm_secret.xlsx"
    path.write_bytes(b"DRMONE\x00\x01" + b"\x00" * 200)
    return str(path)


# ---------------------------------------------------------------------------
# 설정 / 옵션
# ---------------------------------------------------------------------------

@pytest.fixture
def cfg(tmp_path):
    import excel_loader as core
    path = tmp_path / "config.ini"
    path.write_text(
        "[HANA]\nhost = hana.example.invalid\nport = 443\nuser = TESTUSER\n"
        "password = secret\nschema = TESTSCHEMA\n"
        "[GENERAL]\nsheet_name_keyword = TABLE\nload_mode = truncate\n"
        "use_filename_prefix_for_table = false\nxlwings_visible = false\n"
        f"excel_dir = {tmp_path}\n"
        "[HEADER]\nheader_row =\nauto_header_detect = true\n"
        "use_description_row = true\npreview_rows = 100\n"
        f"[LOGGING]\nlog_dir = {tmp_path / 'logs'}\nlog_level = ERROR\n"
        "[ARCHIVE]\nmove_processed_files = false\narchive_dir = archive\n",
        encoding="utf-8")
    return core.load_config(str(path))


@pytest.fixture
def options(cfg, tmp_path):
    import excel_loader as core
    return core.build_options(cfg, str(tmp_path))


@pytest.fixture
def logger():
    import logging
    log = logging.getLogger("test_excel_loader")
    log.handlers = [logging.NullHandler()]
    log.setLevel(logging.CRITICAL)
    return log


@pytest.fixture
def plans(sample_xlsx, options, logger):
    """샘플 파일의 시트 계획. 첫 시트(T_SALES)만 활성화한다."""
    import excel_loader as core
    result, errors = core.collect_sheet_plans([sample_xlsx], options, logger,
                                              allow_drm=False)
    assert not errors, errors
    assert len(result) == 2
    result[1]["enabled"] = False
    return result


# ---------------------------------------------------------------------------
# 가짜 HANA 커넥션
# ---------------------------------------------------------------------------

class FakeCursor:
    def __init__(self, conn):
        self.conn = conn
        self._last = ""

    def execute(self, sql, params=None):
        self.conn.sql.append(sql)
        self._last = sql
        if "AUTOCOMMIT DDL" in sql and not self.conn.allow_ddl_off:
            raise RuntimeError("insufficient privilege: SET TRANSACTION")
        if sql.startswith("INSERT") and self.conn.fail_on_insert:
            raise RuntimeError("INSERT 실패(테스트)")

    def executemany(self, sql, batch):
        self.conn.sql.append(sql)
        if self.conn.fail_on_insert:
            raise RuntimeError("INSERT 실패(테스트)")
        self.conn.rows.extend(batch)

    def fetchone(self):
        return (1 if self.conn.table_exists else 0,)

    def fetchall(self):
        if "SYS.CONSTRAINTS" in self._last:
            return [(c,) for c in self.conn.primary_key]
        return self.conn.columns

    def close(self):
        pass


class FakeConnection:
    """hdbcli 커넥션을 대신해 실행된 SQL과 바인딩 값을 기록한다."""

    def __init__(self, table_exists=False, columns=(), primary_key=(),
                 allow_ddl_off=True, fail_on_insert=False):
        self.sql = []
        self.rows = []
        self.commits = 0
        self.rollbacks = 0
        self.table_exists = table_exists
        self.columns = list(columns)
        self.primary_key = list(primary_key)
        self.allow_ddl_off = allow_ddl_off
        self.fail_on_insert = fail_on_insert

    # -- hdbcli 인터페이스 --------------------------------------------------
    def setautocommit(self, value):
        pass

    def cursor(self):
        return FakeCursor(self)

    def commit(self):
        self.commits += 1

    def rollback(self):
        self.rollbacks += 1

    def close(self):
        pass

    # -- 검사 편의 ----------------------------------------------------------
    def statements(self, *verbs):
        """지정한 동사로 시작하는 SQL만 추린다."""
        return [s for s in self.sql if s.split()[0] in verbs]

    def verbs(self):
        keep = ("TRUNCATE", "INSERT", "UPSERT", "DELETE", "MERGE", "CREATE")
        return [s.split()[0] for s in self.sql if s.split()[0] in keep]


@pytest.fixture
def fake_hana(monkeypatch):
    """connect_hana를 가로채 FakeConnection을 돌려주는 팩토리."""
    import excel_loader as core

    def factory(**kwargs):
        conn = FakeConnection(**kwargs)
        monkeypatch.setattr(core, "connect_hana", lambda *a, **k: conn)
        return conn

    return factory


# 기존 테이블 컬럼 메타데이터 (SYS.TABLE_COLUMNS 형식)
SALES_COLUMNS = [
    ("주문번호", "NVARCHAR", 20, None),
    ("고객명", "NVARCHAR", 50, None),
    ("수량", "NVARCHAR", 20, None),
    ("단가", "DECIMAL", 8, 3),
    ("주문일", "NVARCHAR", 50, None),
    ("비고", "NVARCHAR", 20, None),
]
