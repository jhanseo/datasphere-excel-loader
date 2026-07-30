# -*- coding: utf-8 -*-
"""[KEY] / [NOT NULL] 제약 검사와 기존 테이블 사전 확인."""

import pandas as pd

import excel_loader as core

from conftest import SALES_COLUMNS


def _set(plan, column, **flags):
    for col in plan["columns"]:
        if col["name"] == column:
            col.update(flags)
    return plan


# ---------------------------------------------------------------------------
# check_constraints
# ---------------------------------------------------------------------------

def test_no_issues_when_key_is_unique_and_filled(plans):
    plan = _set(plans[0], "주문번호", is_key=True)
    assert core.check_constraints(plan["df"], plan["columns"]) == []


def test_not_null_violation_is_detected(plans):
    plan = _set(plans[0], "비고", not_null=True)      # 비고에 빈 값 1건
    issues = core.check_constraints(plan["df"], plan["columns"])
    assert len(issues) == 1
    assert issues[0]["kind"] == "not_null"
    assert issues[0]["count"] == 1
    assert issues[0]["rows"] == [3]
    assert "비고" in issues[0]["detail"]


def test_key_null_violation_is_detected(plans):
    plan = _set(plans[0], "비고", is_key=True)
    kinds = [i["kind"] for i in core.check_constraints(plan["df"],
                                                       plan["columns"])]
    assert "key_null" in kinds


def test_duplicate_key_is_detected(plans):
    plan = plans[0]
    plan["df"] = pd.concat([plan["df"], plan["df"].iloc[[0]]],
                           ignore_index=True)
    _set(plan, "주문번호", is_key=True)
    issues = core.check_constraints(plan["df"], plan["columns"])
    dup = [i for i in issues if i["kind"] == "key_duplicate"]
    assert len(dup) == 1
    assert dup[0]["count"] == 2
    assert "A001" in dup[0]["detail"]


def test_composite_key_uniqueness_uses_the_combination(plans):
    """단독으로는 중복이지만 조합으로는 유일한 경우 통과해야 한다."""
    plan = plans[0]
    plan["df"] = plan["df"].copy()
    plan["df"]["주문번호"] = ["A", "A", "B", "B", "C"]
    _set(plan, "주문번호", is_key=True)
    assert [i["kind"] for i in core.check_constraints(plan["df"],
                                                      plan["columns"])] \
        == ["key_duplicate"]
    _set(plan, "주문일", is_key=True)      # 조합하면 유일
    assert core.check_constraints(plan["df"], plan["columns"]) == []


def test_excluded_columns_are_not_checked(plans):
    plan = _set(plans[0], "비고", not_null=True, include=False)
    assert core.check_constraints(plan["df"], plan["columns"]) == []


# ---------------------------------------------------------------------------
# check_target_tables
# ---------------------------------------------------------------------------

def test_new_table_is_reported(plans, cfg, options, logger, fake_hana):
    fake_hana(table_exists=False)
    report = core.check_target_tables(plans, cfg, options, logger)
    assert report["new"] == [("TABLE_T_SALES", "T_SALES", 6)]
    assert not report["type_mismatch"]
    assert not report["constraints_ignored"]
    assert report["error"] is None


def test_type_mismatch_against_existing_table(plans, cfg, options, logger,
                                             fake_hana):
    fake_hana(table_exists=True, columns=SALES_COLUMNS)
    _set(plans[0], "수량", target_type="BIGINT")
    _set(plans[0], "주문일", target_type="DATE")
    report = core.check_target_tables(plans, cfg, options, logger)
    assert ("TABLE_T_SALES", "T_SALES", "수량", "BIGINT",
            "NVARCHAR(20)") in report["type_mismatch"]
    assert ("TABLE_T_SALES", "T_SALES", "주문일", "DATE",
            "NVARCHAR(50)") in report["type_mismatch"]


def test_constraints_ignored_on_existing_table(plans, cfg, options, logger,
                                              fake_hana):
    """기존 테이블에는 KEY/NOT NULL을 적용할 수 없다고 알려야 한다."""
    fake_hana(table_exists=True, columns=SALES_COLUMNS, primary_key=[])
    _set(plans[0], "주문번호", is_key=True)
    _set(plans[0], "고객명", not_null=True)
    report = core.check_target_tables(plans, cfg, options, logger)
    assert len(report["constraints_ignored"]) == 1
    sheet, table, details = report["constraints_ignored"][0]
    assert table == "T_SALES"
    assert any("PRIMARY KEY" in d for d in details)
    assert any("NOT NULL" in d for d in details)


def test_no_warning_when_existing_key_matches(plans, cfg, options, logger,
                                              fake_hana):
    fake_hana(table_exists=True, columns=SALES_COLUMNS,
              primary_key=["주문번호"])
    _set(plans[0], "주문번호", is_key=True)
    report = core.check_target_tables(plans, cfg, options, logger)
    assert report["constraints_ignored"] == []


def test_upsert_without_primary_key_is_flagged(plans, cfg, options, logger,
                                               fake_hana):
    fake_hana(table_exists=True, columns=SALES_COLUMNS, primary_key=[])
    options = dict(options, load_mode=core.LOAD_MODE_UPSERT)
    report = core.check_target_tables(plans, cfg, options, logger)
    assert report["no_primary_key"]
    assert "PRIMARY KEY" in report["no_primary_key"][0][2]


def test_upsert_with_primary_key_is_ok(plans, cfg, options, logger, fake_hana):
    fake_hana(table_exists=True, columns=SALES_COLUMNS,
              primary_key=["주문번호"])
    options = dict(options, load_mode=core.LOAD_MODE_UPSERT)
    report = core.check_target_tables(plans, cfg, options, logger)
    assert report["no_primary_key"] == []


def test_upsert_on_new_table_needs_key_selection(plans, cfg, options, logger,
                                                 fake_hana):
    """새로 만들 테이블인데 [KEY]를 안 골랐으면 UPSERT를 쓸 수 없다."""
    fake_hana(table_exists=False)
    options = dict(options, load_mode=core.LOAD_MODE_UPSERT)
    report = core.check_target_tables(plans, cfg, options, logger)
    assert report["no_primary_key"]


def test_missing_column_is_reported(plans, cfg, options, logger, fake_hana):
    fake_hana(table_exists=True, columns=SALES_COLUMNS[:3])
    report = core.check_target_tables(plans, cfg, options, logger)
    assert report["missing_columns"]
    assert "단가" in report["missing_columns"][0][2]


def test_connection_failure_is_reported_not_raised(plans, cfg, options, logger,
                                                   monkeypatch):
    def boom(*a, **k):
        raise RuntimeError("접속 실패")
    monkeypatch.setattr(core, "connect_hana", boom)
    report = core.check_target_tables(plans, cfg, options, logger)
    assert "접속 실패" in report["error"]


# ---------------------------------------------------------------------------
# format_db_type
# ---------------------------------------------------------------------------

def test_format_db_type():
    assert core.format_db_type("NVARCHAR", 50, None) == "NVARCHAR(50)"
    assert core.format_db_type("DECIMAL", 18, 4) == "DECIMAL(18,4)"
    assert core.format_db_type("DECIMAL", 10, None) == "DECIMAL(10,0)"
    assert core.format_db_type("INTEGER", 4, 0) == "INTEGER"
    assert core.format_db_type("TIMESTAMP", 8, None) == "TIMESTAMP"
