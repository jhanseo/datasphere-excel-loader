# -*- coding: utf-8 -*-
"""데이터 유형 추론 · 표기 검증 · 변환 · 불일치 처리 정책."""

import datetime

import pandas as pd
import pytest

import excel_loader as core


# ---------------------------------------------------------------------------
# 읽기 / 추론
# ---------------------------------------------------------------------------

def test_header_and_comment_extraction(plans):
    """제목 2줄을 건너뛰고 헤더를 찾고, 그 다음 행을 COMMENT로 뽑는다."""
    plan = plans[0]
    assert plan["header_row"] == 2
    assert plan["auto_header_used"] is True
    assert plan["row_count"] == 5
    comments = {c["name"]: c["comment"] for c in plan["columns"]}
    assert comments["수량"] == "Qty"
    assert comments["주문일"] == "Order date"


def test_table_name_from_sheet(plans):
    assert plans[0]["table_name"] == "T_SALES"
    assert plans[1]["table_name"] == "T_CODE"


def test_non_keyword_sheet_is_skipped(plans):
    assert "설명" not in [p["sheet_name"] for p in plans]


def test_new_columns_default_to_no_constraints(plans):
    for col in plans[0]["columns"]:
        assert col["is_key"] is False
        assert col["not_null"] is False
        assert col["include"] is True


@pytest.mark.parametrize("values,expected", [
    ([1, 2, 3], "INTEGER"),
    ([1, 2, 9999999999], "BIGINT"),
    ([1.5, 2.25], "DECIMAL(3,2)"),
    ([True, False], "BOOLEAN"),
    (["a", "b"], "NVARCHAR(20)"),
    ([datetime.datetime(2026, 1, 1), datetime.datetime(2026, 2, 1)], "DATE"),
    ([datetime.datetime(2026, 1, 1, 10, 30)], "TIMESTAMP"),
])
def test_infer_hana_type(values, expected):
    assert core.infer_hana_type(pd.Series(values)) == expected


# ---------------------------------------------------------------------------
# 유형 표기
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("raw,expected", [
    ("integer", "INTEGER"),
    (" timestamp ", "TIMESTAMP"),
    ("nvarchar(30)", "NVARCHAR(30)"),
    ("NVARCHAR", "NVARCHAR(255)"),
    ("decimal(10,2)", "DECIMAL(10,2)"),
    ("decimal", "DECIMAL"),
])
def test_normalize_hana_type(raw, expected):
    assert core.normalize_hana_type(raw) == expected


@pytest.mark.parametrize("raw", ["INTEGERR", "DECIMAL(50,60)", "NVARCHAR(99999)", ""])
def test_normalize_rejects_bad_notation(raw):
    with pytest.raises(ValueError):
        core.normalize_hana_type(raw)


# ---------------------------------------------------------------------------
# check_series (점검의 핵심)
# ---------------------------------------------------------------------------

def test_check_series_finds_unconvertible_values(plans):
    df = plans[0]["df"]
    result = core.check_series(df["수량"], "INTEGER")
    assert result["ok"] is False
    assert result["bad_count"] == 2          # '다섯', 9999999999(범위 초과)
    assert result["total"] == 5
    bad_values = [v for _, v in result["samples"]]
    assert "다섯" in bad_values


def test_bigint_accepts_large_value(plans):
    result = core.check_series(plans[0]["df"]["수량"], "BIGINT")
    assert result["bad_count"] == 1          # '다섯'만 실패
    assert result["samples"][0][1] == "다섯"


def test_check_series_string_to_date(plans):
    assert core.check_series(plans[0]["df"]["주문일"], "DATE")["bad_count"] == 0


def test_check_series_length_overflow(plans):
    result = core.check_series(plans[0]["df"]["고객명"], "NVARCHAR(20)")
    assert result["bad_count"] == 1


def test_check_series_reports_notation_error(plans):
    result = core.check_series(plans[0]["df"]["수량"], "INTEGERR")
    assert result["error"]
    assert result["ok"] is False


def test_check_series_counts_nulls_separately(plans):
    result = core.check_series(plans[0]["df"]["비고"], "NVARCHAR(50)")
    assert result["bad_count"] == 0
    assert result["null_count"] == 1


# ---------------------------------------------------------------------------
# 불일치 처리 정책
# ---------------------------------------------------------------------------

def _with_policy(plan, policy, target="INTEGER", column="수량"):
    cols = [dict(c) for c in plan["columns"]]
    for c in cols:
        if c["name"] == column:
            c["target_type"] = target
            c["policy"] = policy
    return cols


def test_policy_null_keeps_all_rows(plans, logger):
    plan = plans[0]
    out, defs, report = core.apply_column_types(
        plan["df"], _with_policy(plan, core.POLICY_NULL), logger)
    assert len(out) == 5
    assert report["nulled"]["수량"] == 2
    assert dict(defs)["수량"] == "INTEGER"
    assert out["수량"].tolist() == [10, 3, None, 7, None]


def test_policy_skip_row_drops_offending_rows(plans, logger):
    plan = plans[0]
    out, _, report = core.apply_column_types(
        plan["df"], _with_policy(plan, core.POLICY_SKIP_ROW), logger)
    assert len(out) == 3
    assert report["dropped_rows"] == 2


def test_policy_revert_falls_back_to_inferred_type(plans, logger):
    plan = plans[0]
    out, defs, report = core.apply_column_types(
        plan["df"], _with_policy(plan, core.POLICY_REVERT), logger)
    assert len(out) == 5
    assert dict(defs)["수량"] == "NVARCHAR(20)"
    assert report["reverted"]


def test_policy_abort_raises(plans, logger):
    plan = plans[0]
    with pytest.raises(core.TypeMismatchAbort):
        core.apply_column_types(
            plan["df"], _with_policy(plan, core.POLICY_ABORT), logger)


def test_successful_conversion_produces_real_dates(plans, logger):
    plan = plans[0]
    cols = [dict(c) for c in plan["columns"]]
    for c in cols:
        if c["name"] == "주문일":
            c["target_type"] = "DATE"
        if c["name"] == "단가":
            c["target_type"] = "DOUBLE"
    out, defs, _ = core.apply_column_types(plan["df"], cols, logger)
    assert dict(defs)["주문일"] == "DATE"
    assert all(isinstance(v, datetime.date) for v in out["주문일"])


def test_excluded_column_is_dropped(plans, logger):
    plan = plans[0]
    cols = [dict(c) for c in plan["columns"]]
    for c in cols:
        if c["name"] == "비고":
            c["include"] = False
    out, _, _ = core.apply_column_types(plan["df"], cols, logger)
    assert "비고" not in out.columns


# ---------------------------------------------------------------------------
# DB 값 변환
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("value", [None, float("nan"), pd.NaT])
def test_to_db_value_normalizes_nulls(value):
    assert core.to_db_value(value) is None


def test_to_db_value_converts_timestamp():
    assert core.to_db_value(pd.Timestamp("2026-01-01")).year == 2026
