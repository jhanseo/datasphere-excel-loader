# -*- coding: utf-8 -*-
"""Wizard GUI 로직 테스트.

실제 PyQt5 없이 tests/mockqt의 Qt 대역으로 시그널·핸들러·대화상자 흐름을
검증한다. 자세한 내용은 tests/mockqt/README.md 참고.
"""

import sys

import pytest

from conftest import MOCKQT, SALES_COLUMNS

# 실제 PyQt5가 설치되어 있어도, 화면 없는 환경에서 돌 수 있도록 대역을 우선한다.
if MOCKQT not in sys.path:
    sys.path.insert(0, MOCKQT)

wizard = pytest.importorskip("wizard")
from PyQt5.QtCore import Qt                                    # noqa: E402
from PyQt5.QtWidgets import QDialog, QMessageBox               # noqa: E402


@pytest.fixture
def win(tmp_path, cfg, sample_xlsx, monkeypatch):
    """config.ini가 놓인 임시 폴더를 app_dir로 쓰는 MainWindow.

    cfg 픽스처가 tmp_path/config.ini를 이미 만들어 두므로, app_dir만 바꿔주면
    MainWindow가 그 설정을 읽는다.
    """
    assert (tmp_path / "config.ini").exists()
    monkeypatch.setattr(wizard, "app_dir", lambda: str(tmp_path))
    QMessageBox.CALLS.clear()
    QMessageBox.ANSWERS.clear()
    window = wizard.MainWindow()
    window._add_file_paths([sample_xlsx])
    return window


def preview(window):
    """미리보기 워커를 동기적으로 돌려 plans를 채운다."""
    worker = wizard.PreviewWorker(window.selected_files(),
                                  window.current_options(), window.logger)
    captured = {}
    worker.finished_ok.connect(lambda p, e: captured.update(plans=p, errors=e))
    worker.run()
    window._on_preview_done(captured["plans"], captured["errors"])
    return window.plans


def row_of(window, name):
    for row in range(window.col_table.rowCount()):
        if window.col_table.item(row, wizard.C_NAME).text() == name:
            return row
    raise AssertionError(f"컬럼 없음: {name}")


# ---------------------------------------------------------------------------
# 기본 구성
# ---------------------------------------------------------------------------

def test_window_has_file_config_preview_and_log_tabs(win):
    labels = [t[1] for t in win.tabs._tabs]
    assert len(labels) == 8
    assert labels[0].startswith("1.")
    assert any("일반" in l for l in labels)


def test_config_round_trips_through_widgets(win):
    win.widgets[("HANA", "schema")][0].setText("NEWSCHEMA")
    win.widgets[("GENERAL", "load_mode")][0].setCurrentText("insert")
    options = win.current_options()
    assert options["schema"] == "NEWSCHEMA"
    assert options["load_mode"] == "insert"


def test_load_mode_offers_three_choices(win):
    combo = win.widgets[("GENERAL", "load_mode")][0]
    assert [text for text, _ in combo._items] == ["truncate", "insert", "upsert"]


def test_drm_file_is_marked_in_the_list(win, drm_xlsx):
    win._add_file_paths([drm_xlsx])
    assert "DRM 1개" in win.file_count_label.text()


# ---------------------------------------------------------------------------
# 미리보기
# ---------------------------------------------------------------------------

def test_preview_populates_columns_and_comments(win):
    preview(win)
    assert win.col_table.rowCount() == 6
    comments = [win.col_table.item(r, wizard.C_COMMENT).text()
                for r in range(6)]
    assert "Qty" in comments


def test_preview_columns_include_key_and_not_null(win):
    preview(win)
    assert wizard.COL_HEADERS[wizard.C_KEY] == "KEY"
    assert wizard.COL_HEADERS[wizard.C_NOTNULL] == "NOT NULL"
    assert win.col_table.columnCount() == 10


def test_changing_file_list_invalidates_preview(win):
    preview(win)
    assert win.plans
    win.clear_files()
    assert win.plans == []
    assert win.preview_files == []


# ---------------------------------------------------------------------------
# 데이터 유형 변경
# ---------------------------------------------------------------------------

def test_type_change_updates_plan_and_marks_unchecked(win):
    preview(win)
    row = row_of(win, "수량")
    win.col_table.cellWidget(row, wizard.C_TARGET).setCurrentText("INTEGER")
    col = win.plans[0]["columns"][row]
    assert col["target_type"] == "INTEGER"
    assert not col.get("checked")
    assert "점검 (" in win.btn_check.text()


def test_check_offers_all_four_policies(win):
    preview(win)
    row = row_of(win, "수량")
    win.col_table.cellWidget(row, wizard.C_TARGET).setCurrentText("INTEGER")

    captured = {}
    original = wizard.CheckDialog.__init__

    def spy(self, issues, parent=None):
        original(self, issues, parent)
        captured["dialog"] = self

    wizard.CheckDialog.__init__ = spy
    wizard.CheckDialog.exec_ = lambda self: QDialog.Accepted
    try:
        assert win.run_check(silent=True) is True
    finally:
        wizard.CheckDialog.__init__ = original

    dialog = captured["dialog"]
    assert dialog.table.rowCount() == 1
    assert len(dialog.combos[0]._items) == 4
    assert "다섯" in dialog.table.item(0, 4).text()


def test_invalid_type_notation_blocks_check(win):
    preview(win)
    row = row_of(win, "단가")
    win.col_table.cellWidget(row, wizard.C_TARGET).setCurrentText("DECIMAL(50,60)")
    before = len(QMessageBox.CALLS)
    assert win.run_check(silent=True) is False
    assert QMessageBox.CALLS[before][0] == "crit"


# ---------------------------------------------------------------------------
# KEY / NOT NULL 체크박스
# ---------------------------------------------------------------------------

def test_key_checkbox_also_enables_not_null(win):
    preview(win)
    row = row_of(win, "주문번호")
    win.col_table.item(row, wizard.C_KEY).setCheckState(Qt.Checked)
    col = win.plans[0]["columns"][row]
    assert col["is_key"] is True
    assert col["not_null"] is True
    assert win.col_table.item(row, wizard.C_NOTNULL).checkState() == Qt.Checked


def test_not_null_cannot_be_cleared_on_a_key_column(win):
    preview(win)
    row = row_of(win, "주문번호")
    win.col_table.item(row, wizard.C_KEY).setCheckState(Qt.Checked)
    win.col_table.item(row, wizard.C_NOTNULL).setCheckState(Qt.Unchecked)
    assert win.plans[0]["columns"][row]["not_null"] is True
    assert win.col_table.item(row, wizard.C_NOTNULL).checkState() == Qt.Checked


def test_not_null_can_be_set_alone(win):
    preview(win)
    row = row_of(win, "고객명")
    win.col_table.item(row, wizard.C_NOTNULL).setCheckState(Qt.Checked)
    col = win.plans[0]["columns"][row]
    assert col["not_null"] is True
    assert col["is_key"] is False


def test_unchecking_use_excludes_the_column(win):
    preview(win)
    row = row_of(win, "비고")
    win.col_table.item(row, wizard.C_USE).setCheckState(Qt.Unchecked)
    assert win.plans[0]["columns"][row]["include"] is False


def test_not_null_violation_blocks_check(win):
    preview(win)
    row = row_of(win, "비고")            # 빈 값이 1건 있는 컬럼
    win.col_table.item(row, wizard.C_NOTNULL).setCheckState(Qt.Checked)

    captured = {}
    original = wizard.ConstraintCheckDialog.__init__

    def spy(self, violations, report, parent=None):
        original(self, violations, report, parent)
        captured["violations"] = violations

    wizard.ConstraintCheckDialog.__init__ = spy
    wizard.ConstraintCheckDialog.exec_ = lambda self: QDialog.Accepted
    try:
        assert win.run_check(silent=True) is False
    finally:
        wizard.ConstraintCheckDialog.__init__ = original
    assert captured["violations"][0][2]["kind"] == "not_null"


def test_check_warns_when_target_table_already_exists(win, fake_hana):
    """기존 테이블에는 KEY/NOT NULL을 적용할 수 없다고 경고해야 한다."""
    preview(win)
    win.widgets[("HANA", "host")][0].setText("hana.example.invalid")
    win.widgets[("HANA", "schema")][0].setText("TESTSCHEMA")
    win.plans[1]["enabled"] = False
    row = row_of(win, "주문번호")
    win.col_table.item(row, wizard.C_KEY).setCheckState(Qt.Checked)
    fake_hana(table_exists=True, columns=SALES_COLUMNS, primary_key=[])

    captured = {}
    original = wizard.ConstraintCheckDialog.__init__

    def spy(self, violations, report, parent=None):
        original(self, violations, report, parent)
        captured["report"] = report

    wizard.ConstraintCheckDialog.__init__ = spy
    wizard.ConstraintCheckDialog.exec_ = lambda self: QDialog.Accepted
    try:
        assert win.run_check(silent=False) is True
        if win.check_worker:
            win.check_worker.wait()
    finally:
        wizard.ConstraintCheckDialog.__init__ = original

    assert captured["report"]["constraints_ignored"]


def test_check_skips_db_lookup_without_connection_info(win):
    preview(win)
    row = row_of(win, "주문번호")
    win.col_table.item(row, wizard.C_KEY).setCheckState(Qt.Checked)
    win.widgets[("HANA", "host")][0].setText("")
    before = len(QMessageBox.CALLS)
    win.run_check(silent=False)
    assert any("접속 정보" in m for _, _, m in QMessageBox.CALLS[before:])


# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------

def test_run_aborts_entirely_when_a_drm_file_is_present(win, drm_xlsx):
    """미리보기를 거치지 않은 DRM 파일이 있으면 전체 실행을 멈춘다."""
    win._add_file_paths([drm_xlsx])
    before = len(QMessageBox.CALLS)
    win.run_execute(dry_run=True)
    calls = QMessageBox.CALLS[before:]
    assert len(calls) == 1
    assert calls[0][0] == "warn"
    assert "DRM" in calls[0][1]
    assert "미리보기" in calls[0][2]
    assert win.plans == []


def test_plain_files_run_without_preview(win):
    QMessageBox.ANSWERS.append(QMessageBox.Ok)
    win.run_execute(dry_run=True)
    assert len(win.plans) == 2


def test_dry_run_skips_the_preflight_table_check(win, monkeypatch):
    preview(win)
    calls = []
    monkeypatch.setattr(wizard.PreflightWorker, "run",
                        lambda self: calls.append(1))
    QMessageBox.ANSWERS.append(QMessageBox.Ok)
    win.run_execute(dry_run=True)
    assert calls == []


def test_preflight_dialog_can_cancel_the_upload(win, fake_hana, monkeypatch):
    """유형 불일치 대화상자에서 취소하면 업로드가 시작되지 않아야 한다."""
    preview(win)
    win.plans[0]["enabled"] = False          # T_CODE 시트만 사용
    win.plans[1]["enabled"] = True
    for col in win.plans[1]["columns"]:
        if col["name"] == "코드":
            col["target_type"] = "NVARCHAR(100)"
    fake_hana(table_exists=True,
              columns=[("코드", "NVARCHAR", 5, None),
                       ("이름", "NVARCHAR", 20, None)])

    started = []
    monkeypatch.setattr(win, "_start_upload",
                        lambda *a, **k: started.append(1))
    wizard.TableCheckDialog.exec_ = lambda self: QDialog.Rejected
    QMessageBox.ANSWERS.extend([QMessageBox.No, QMessageBox.Ok])
    win.run_execute(dry_run=False)
    if win.worker:
        win.worker.wait()
    assert started == []


def test_execute_reports_needs_check_for_constraints(win):
    preview(win)
    row = row_of(win, "주문번호")
    win.col_table.item(row, wizard.C_KEY).setCheckState(Qt.Checked)
    assert win._needs_check() is True
    before = len(QMessageBox.CALLS)
    QMessageBox.ANSWERS.extend([QMessageBox.No, QMessageBox.Cancel])
    win.run_execute(dry_run=True)
    prompts = [m for _, t, m in QMessageBox.CALLS[before:]
               if t == "점검이 필요합니다"]
    assert prompts and "KEY" in prompts[0]
