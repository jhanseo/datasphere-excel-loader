# -*- coding: utf-8 -*-
"""적재 모드(truncate/insert/upsert), DDL 생성, 트랜잭션 안전장치."""

import pytest

import excel_loader as core

from conftest import SALES_COLUMNS


def _run(plans, cfg, options, logger, mode, **conn_kwargs):
    options = dict(options)
    options["load_mode"] = mode
    options["truncate"] = mode == core.LOAD_MODE_TRUNCATE
    return core.execute_plans(plans, cfg, options, logger)


# ---------------------------------------------------------------------------
# 모드별 SQL
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("mode,expected", [
    (core.LOAD_MODE_TRUNCATE, ["CREATE", "TRUNCATE", "INSERT"]),
    (core.LOAD_MODE_INSERT, ["CREATE", "INSERT"]),
    (core.LOAD_MODE_UPSERT, ["CREATE", "UPSERT"]),
])
def test_load_mode_emits_expected_statements(plans, cfg, options, logger,
                                             fake_hana, mode, expected):
    conn = fake_hana()
    _run(plans, cfg, options, logger, mode)
    assert conn.verbs() == expected


@pytest.mark.parametrize("mode", list(core.LOAD_MODES))
def test_delete_and_merge_delta_are_never_used(plans, cfg, options, logger,
                                               fake_hana, mode):
    """DELETE FROM / MERGE DELTA 경로는 제거되었다."""
    conn = fake_hana()
    _run(plans, cfg, options, logger, mode)
    assert not [s for s in conn.sql if s.startswith(("DELETE FROM", "MERGE DELTA"))]


def test_upsert_uses_with_primary_key(plans, cfg, options, logger, fake_hana):
    conn = fake_hana()
    _run(plans, cfg, options, logger, core.LOAD_MODE_UPSERT)
    upsert = conn.statements("UPSERT")[0]
    assert upsert.endswith("WITH PRIMARY KEY")
    assert '"TESTSCHEMA"."T_SALES"' in upsert


def test_insert_mode_does_not_clear_table(plans, cfg, options, logger, fake_hana):
    conn = fake_hana()
    _run(plans, cfg, options, logger, core.LOAD_MODE_INSERT)
    assert not conn.statements("TRUNCATE")


def test_rows_are_bound_with_converted_values(plans, cfg, options, logger,
                                              fake_hana):
    conn = fake_hana()
    _run(plans, cfg, options, logger, core.LOAD_MODE_TRUNCATE)
    assert len(conn.rows) == 5
    assert conn.rows[0][0] == "A001"
    assert conn.commits == 1


# ---------------------------------------------------------------------------
# CREATE TABLE : COMMENT / NOT NULL / PRIMARY KEY
# ---------------------------------------------------------------------------

def test_create_table_includes_comments(plans, cfg, options, logger, fake_hana):
    conn = fake_hana()
    _run(plans, cfg, options, logger, core.LOAD_MODE_TRUNCATE)
    ddl = conn.statements("CREATE")[0]
    assert "COMMENT 'Qty'" in ddl
    assert "COMMENT 'Order date'" in ddl


def test_key_and_not_null_appear_in_ddl(plans, cfg, options, logger, fake_hana):
    conn = fake_hana()
    for col in plans[0]["columns"]:
        if col["name"] == "주문번호":
            col["is_key"] = True
        if col["name"] == "고객명":
            col["not_null"] = True
    _run(plans, cfg, options, logger, core.LOAD_MODE_TRUNCATE)
    ddl = conn.statements("CREATE")[0]
    assert '"주문번호" NVARCHAR(20) NOT NULL' in ddl
    assert '"고객명" NVARCHAR(50) NOT NULL' in ddl
    assert 'PRIMARY KEY ("주문번호")' in ddl
    # 지정하지 않은 컬럼에는 NOT NULL이 붙지 않는다.
    assert '"비고" NVARCHAR(20) COMMENT' in ddl


def test_composite_primary_key_preserves_order(plans, cfg, options, logger,
                                               fake_hana):
    conn = fake_hana()
    for col in plans[0]["columns"]:
        if col["name"] in ("주문번호", "주문일"):
            col["is_key"] = True
    _run(plans, cfg, options, logger, core.LOAD_MODE_TRUNCATE)
    assert 'PRIMARY KEY ("주문번호", "주문일")' in conn.statements("CREATE")[0]


def test_existing_table_is_not_recreated(plans, cfg, options, logger, fake_hana):
    conn = fake_hana(table_exists=True, columns=SALES_COLUMNS)
    for col in plans[0]["columns"]:
        if col["name"] == "주문번호":
            col["is_key"] = True
    _run(plans, cfg, options, logger, core.LOAD_MODE_TRUNCATE)
    assert not conn.statements("CREATE")


# ---------------------------------------------------------------------------
# 트랜잭션 안전장치
# ---------------------------------------------------------------------------

def test_autocommit_ddl_is_turned_off(plans, cfg, options, logger, fake_hana):
    conn = fake_hana()
    _run(plans, cfg, options, logger, core.LOAD_MODE_TRUNCATE)
    assert "SET TRANSACTION AUTOCOMMIT DDL OFF" in conn.sql


def test_truncate_still_used_when_autocommit_cannot_be_disabled(
        plans, cfg, options, logger, fake_hana):
    """권한이 없어도 TRUNCATE를 쓴다(DELETE 폴백은 제거됨)."""
    conn = fake_hana(allow_ddl_off=False)
    _run(plans, cfg, options, logger, core.LOAD_MODE_TRUNCATE)
    assert conn.statements("TRUNCATE")
    assert not [s for s in conn.sql if s.startswith("DELETE FROM")]


def test_failed_insert_rolls_back(plans, cfg, options, logger, fake_hana):
    conn = fake_hana(fail_on_insert=True)
    summary, had_error = _run(plans, cfg, options, logger,
                              core.LOAD_MODE_TRUNCATE)
    assert had_error is True
    assert conn.rollbacks == 1
    assert conn.commits == 0
    assert summary[0][4].startswith("실패")


# ---------------------------------------------------------------------------
# dry-run
# ---------------------------------------------------------------------------

def test_dry_run_does_not_touch_the_database(plans, cfg, options, logger,
                                             monkeypatch):
    called = []
    monkeypatch.setattr(core, "connect_hana",
                        lambda *a, **k: called.append(1))
    summary, had_error = core.execute_plans(plans, cfg, options, logger,
                                            dry_run=True)
    assert not called
    assert had_error is False
    assert summary[0][4] == "DRY-RUN"


# ---------------------------------------------------------------------------
# 설정 호환성
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("mode", list(core.LOAD_MODES))
def test_load_mode_round_trips_through_config(cfg, tmp_path, mode):
    cfg.set("GENERAL", "load_mode", mode)
    assert core.build_options(cfg, str(tmp_path))["load_mode"] == mode


@pytest.mark.parametrize("legacy,expected", [
    ("true", core.LOAD_MODE_TRUNCATE),
    ("false", core.LOAD_MODE_INSERT),
])
def test_legacy_truncate_before_insert_is_still_honoured(cfg, tmp_path,
                                                         legacy, expected):
    """예전 설정 파일도 그대로 동작해야 한다."""
    cfg.remove_option("GENERAL", "load_mode")
    cfg.set("GENERAL", "truncate_before_insert", legacy)
    assert core.build_options(cfg, str(tmp_path))["load_mode"] == expected


def test_unknown_load_mode_falls_back_to_truncate(cfg, tmp_path):
    cfg.set("GENERAL", "load_mode", "bogus")
    assert core.build_options(cfg, str(tmp_path))["load_mode"] == \
        core.LOAD_MODE_TRUNCATE
