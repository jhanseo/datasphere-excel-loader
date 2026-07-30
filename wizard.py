# -*- coding: utf-8 -*-
"""
Excel -> HANA Datasphere Loader :: Wizard GUI
=============================================
excel_loader.py의 기능을 PyQt5 GUI로 감싼 실행 마법사.

탭 구성
  1. 파일 선택   : 업로드할 엑셀 파일을 여러 개 선택 (DRM 여부 표시)
  2~6. 설정 탭   : config.ini의 각 섹션(HANA 접속/일반/헤더/로깅/보관)
  7. 미리보기    : 시트별 컬럼 목록, 데이터 유형, Comment, 샘플 데이터
  8. 실행 로그   : 진행 상황 및 결과 요약

동작 규칙
  [미리보기] DRM 우회 기능(설치된 Excel + xlwings)을 사용해 파일을 읽는다.
             읽어온 데이터는 캐시되어 [실행] 시 그대로 사용된다.
  [점검]     미리보기에서 변경한 데이터 유형으로 실제 변환이 가능한지 검사하고,
             맞지 않는 데이터의 처리 방식을 사용자가 선택하게 한다.
  [실행]     미리보기 캐시가 있으면 그대로 업로드한다. 캐시가 없는 상태에서
             DRM 파일이 포함되어 있으면 안내 문구를 띄우고 전체 실행을 중단한다.

실행:
    python wizard.py
필요 라이브러리:
    pip install PyQt5 pandas openpyxl xlrd hdbcli
    pip install xlwings pywin32   # DRM 파일 처리 (Windows + Excel)
"""

import logging
import os
import sys
import traceback

from PyQt5.QtCore import Qt, QObject, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QIntValidator
from PyQt5.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog,
    QDialogButtonBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QSplitter,
    QTabWidget, QTableWidget, QTableWidgetItem, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget,
)

import excel_loader as core

APP_TITLE = "Excel → HANA Datasphere 업로드 마법사"
SAMPLE_LIMIT = 3

# 미리보기 컬럼 표의 열 구성
(C_USE, C_NAME, C_ORIG, C_INFER, C_TARGET,
 C_KEY, C_NOTNULL, C_COMMENT, C_NULLCNT, C_SAMPLE) = range(10)

COL_HEADERS = ["사용", "컬럼명", "원본 컬럼명", "추론 유형", "지정 유형",
               "KEY", "NOT NULL", "Comment", "NULL", "샘플 값"]


def app_dir() -> str:
    """스크립트/exe가 놓인 폴더 (config.ini 기본 위치)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def short_text(value, limit=40) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\n", " ").replace("\r", " ").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


# ---------------------------------------------------------------------------
# 로깅 -> GUI 연결
# ---------------------------------------------------------------------------

class LogBridge(QObject):
    message = pyqtSignal(str, int)


class QtLogHandler(logging.Handler):
    def __init__(self, bridge: LogBridge):
        super().__init__()
        self.bridge = bridge

    def emit(self, record):
        try:
            self.bridge.message.emit(self.format(record), record.levelno)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 백그라운드 작업 스레드
# ---------------------------------------------------------------------------

class PreviewWorker(QThread):
    progressed = pyqtSignal(str)
    finished_ok = pyqtSignal(object, object)   # plans, errors
    failed = pyqtSignal(str, str)              # 제목, 본문

    def __init__(self, file_paths, options, logger, parent=None):
        super().__init__(parent)
        self.file_paths = file_paths
        self.options = options
        self.logger = logger
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        com_ready = False
        try:
            # xlwings(COM)를 워커 스레드에서 쓰려면 COM 초기화가 필요하다.
            try:
                import pythoncom
                pythoncom.CoInitialize()
                com_ready = True
            except Exception:
                pass

            plans, errors = core.collect_sheet_plans(
                self.file_paths, self.options, self.logger, allow_drm=True,
                progress=self.progressed.emit, cancelled=lambda: self._cancel)
            self.finished_ok.emit(plans, errors)
        except core.DrmProtectedError as e:
            self.failed.emit("DRM 파일을 열 수 없습니다", str(e))
        except Exception as e:
            self.logger.error("미리보기 실패: %s", e)
            self.logger.debug("%s", traceback.format_exc())
            self.failed.emit("미리보기 실패", str(e))
        finally:
            if com_ready:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


class PreflightWorker(QThread):
    """업로드 전에 대상 테이블 구조를 확인한다(메타데이터 조회만)."""
    finished_ok = pyqtSignal(object)       # report dict
    failed = pyqtSignal(str, str)

    def __init__(self, plans, cfg, options, logger, parent=None):
        super().__init__(parent)
        self.plans = plans
        self.cfg = cfg
        self.options = options
        self.logger = logger
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            report = core.check_target_tables(
                self.plans, self.cfg, self.options, self.logger,
                cancelled=lambda: self._cancel)
            self.finished_ok.emit(report)
        except Exception as e:
            self.logger.error("테이블 확인 실패: %s", e)
            self.failed.emit("테이블 확인 실패", str(e))


class ExecuteWorker(QThread):
    progressed = pyqtSignal(int, int, str)
    finished_ok = pyqtSignal(object, bool)     # summary, had_error
    failed = pyqtSignal(str, str)

    def __init__(self, plans, cfg, options, logger, dry_run=False, parent=None):
        super().__init__(parent)
        self.plans = plans
        self.cfg = cfg
        self.options = options
        self.logger = logger
        self.dry_run = dry_run
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            summary, had_error = core.execute_plans(
                self.plans, self.cfg, self.options, self.logger,
                dry_run=self.dry_run,
                progress=lambda i, n, m: self.progressed.emit(i, n, m),
                cancelled=lambda: self._cancel)
            self.finished_ok.emit(summary, had_error)
        except Exception as e:
            self.logger.error("실행 실패: %s", e)
            self.logger.debug("%s", traceback.format_exc())
            self.failed.emit("실행 실패", str(e))


# ---------------------------------------------------------------------------
# 점검 결과 대화상자
# ---------------------------------------------------------------------------

class CheckDialog(QDialog):
    """유형 불일치 목록과 처리 방식 선택 대화상자."""

    HEADERS = ["시트", "컬럼", "지정 유형", "불일치", "예시 값", "처리 방식"]

    def __init__(self, issues, parent=None):
        super().__init__(parent)
        self.issues = issues
        self.setWindowTitle("데이터 유형 점검 결과")
        self.resize(1000, 520)

        layout = QVBoxLayout(self)

        total_bad = sum(i["bad_count"] for i in issues)
        head = QLabel(
            f"<b>{len(issues)}개 컬럼</b>에서 지정한 데이터 유형으로 변환할 수 없는 값이 "
            f"<b>{total_bad:,}건</b> 발견되었습니다.<br>"
            "각 컬럼별로 이 값들을 어떻게 처리할지 선택한 뒤 [적용]을 눌러 주세요."
        )
        head.setWordWrap(True)
        layout.addWidget(head)

        bulk_row = QHBoxLayout()
        bulk_row.addWidget(QLabel("전체 일괄 적용:"))
        self.bulk_combo = QComboBox()
        for p in core.POLICY_ORDER:
            self.bulk_combo.addItem(core.POLICY_LABELS[p], p)
        bulk_row.addWidget(self.bulk_combo, 1)
        bulk_btn = QPushButton("모든 컬럼에 적용")
        bulk_btn.clicked.connect(lambda *_: self._apply_bulk())
        bulk_row.addWidget(bulk_btn)
        layout.addLayout(bulk_row)

        self.table = QTableWidget(len(issues), len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.combos = []
        for row, issue in enumerate(issues):
            self._set(row, 0, f"{issue['file_name']} / {issue['sheet_name']}")
            self._set(row, 1, issue["column"])
            self._set(row, 2, issue["target_type"])
            self._set(row, 3, f"{issue['bad_count']:,}건 / {issue['total']:,}행")
            self._set(row, 4, issue["sample_text"])
            combo = QComboBox()
            for p in core.POLICY_ORDER:
                combo.addItem(core.POLICY_LABELS[p], p)
            idx = combo.findData(issue.get("policy", core.POLICY_NULL))
            combo.setCurrentIndex(max(0, idx))
            combo.currentIndexChanged.connect(lambda *_: self._update_hint())
            self.table.setCellWidget(row, 5, combo)
            self.combos.append(combo)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        for c in (0, 1, 2, 3, 5):
            header.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        layout.addWidget(self.table, 1)

        self.hint = QLabel()
        self.hint.setWordWrap(True)
        self.hint.setStyleSheet(
            "background:#FFF8E1; border:1px solid #E6C200; padding:8px;")
        layout.addWidget(self.hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("적용")
        buttons.button(QDialogButtonBox.Cancel).setText("취소")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._update_hint()

    def _set(self, row, col, text):
        item = QTableWidgetItem(str(text))
        item.setToolTip(str(text))
        self.table.setItem(row, col, item)

    def _apply_bulk(self):
        policy = self.bulk_combo.currentData()
        for combo in self.combos:
            combo.setCurrentIndex(combo.findData(policy))
        self._update_hint()

    def _update_hint(self):
        chosen = []
        for combo in self.combos:
            p = combo.currentData()
            if p not in chosen:
                chosen.append(p)
        lines = [f"<b>{core.POLICY_LABELS[p]}</b> — {core.POLICY_DESCRIPTIONS[p]}"
                 for p in chosen]
        self.hint.setText("선택한 처리 방식 안내<br>" + "<br>".join(lines))

    def policies(self):
        return [combo.currentData() for combo in self.combos]


class TableCheckDialog(QDialog):
    """업로드 전 대상 테이블 확인 결과. 유형 불일치를 사용자에게 알린다."""

    HEADERS = ["시트", "테이블", "컬럼", "지정 유형", "실제 테이블 유형"]

    def __init__(self, report, parent=None):
        super().__init__(parent)
        self.setWindowTitle("업로드 전 테이블 확인")
        self.resize(950, 500)
        layout = QVBoxLayout(self)

        mismatch = report["type_mismatch"]
        head = QLabel(
            f"이미 존재하는 테이블의 컬럼 <b>{len(mismatch)}개</b>가 미리보기에서 "
            "지정한 데이터 유형과 다릅니다.<br><br>"
            "기존 테이블은 다시 만들지 않기 때문에 <b>지정한 유형은 적용되지 않고, "
            "테이블에 이미 정의된 유형이 그대로 사용됩니다.</b> 값은 테이블 유형에 "
            "맞춰 변환되어 들어가므로, 자리수가 모자라면 오류가 나거나 소수점이 "
            "잘릴 수 있습니다.<br><br>"
            "지정한 유형을 반드시 적용해야 한다면 업로드를 취소하고, "
            "Datasphere에서 테이블 컬럼 유형을 변경하거나 테이블을 삭제한 뒤 "
            "다시 실행하세요."
        )
        head.setWordWrap(True)
        head.setStyleSheet(
            "background:#FFF8E1; border:1px solid #E6C200; padding:10px;")
        layout.addWidget(head)

        table = QTableWidget(len(mismatch), len(self.HEADERS))
        table.setHorizontalHeaderLabels(self.HEADERS)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        for row, (sheet, tname, col, want, have) in enumerate(mismatch):
            for c, v in enumerate((sheet, tname, col, want, have)):
                item = QTableWidgetItem(str(v))
                item.setToolTip(str(v))
                if c >= 3:
                    item.setBackground(QColor("#FFECEC"))
                table.setItem(row, c, item)
        header = table.horizontalHeader()
        for c in range(len(self.HEADERS)):
            header.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        layout.addWidget(table, 1)

        extra = []
        if report["new"]:
            names = ", ".join(f"{t}({n}컬럼)" for _, t, n in report["new"][:8])
            extra.append(f"새로 생성될 테이블 {len(report['new'])}개 — {names}")
        if report["missing_columns"]:
            for sheet, tname, cols in report["missing_columns"][:5]:
                extra.append(
                    f"'{tname}'에 없는 컬럼: {', '.join(cols)} → 이 시트는 실패합니다")
        if extra:
            note = QLabel("<br>".join(extra))
            note.setWordWrap(True)
            note.setStyleSheet("padding:6px; color:#555;")
            layout.addWidget(note)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("이대로 업로드 계속")
        buttons.button(QDialogButtonBox.Cancel).setText("업로드 취소")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class ConstraintCheckDialog(QDialog):
    """[KEY] / [NOT NULL] 지정에 대한 점검 결과."""

    HEADERS = ["시트", "테이블", "구분", "내용", "예시 행"]

    KIND_LABELS = {
        "not_null": "NOT NULL 위반",
        "key_null": "KEY 빈 값",
        "key_duplicate": "KEY 중복",
    }

    def __init__(self, violations, table_report, parent=None):
        super().__init__(parent)
        self.setWindowTitle("KEY / NOT NULL 점검 결과")
        self.resize(1000, 560)
        layout = QVBoxLayout(self)

        blocking = bool(violations)
        existing = (table_report or {}).get("constraints_ignored") or []
        no_pk = (table_report or {}).get("no_primary_key") or []

        msgs = []
        if blocking:
            msgs.append(
                f"<b>지정한 제약과 데이터가 맞지 않는 항목이 {len(violations)}건 "
                "있습니다.</b> 이 상태로는 테이블 생성 또는 적재가 실패합니다. "
                "해당 컬럼의 체크를 해제하거나 엑셀 데이터를 수정한 뒤 다시 "
                "점검해 주세요.")
        if existing:
            msgs.append(
                f"<b>이미 존재하는 테이블 {len(existing)}개</b>에는 "
                "[KEY]/[NOT NULL] 지정을 <b>적용할 수 없습니다.</b> 기존 테이블은 "
                "다시 생성하지 않으므로 제약은 CREATE TABLE 시점에만 반영됩니다. "
                "반드시 적용해야 한다면 Datasphere에서 테이블을 삭제하거나 "
                "제약을 직접 변경한 뒤 다시 실행하세요.")
        if no_pk:
            msgs.append(
                "<b>UPSERT 모드인데 PRIMARY KEY가 없습니다.</b> "
                "UPSERT는 PRIMARY KEY 기준으로 동작하므로 그대로 실행하면 "
                "실패합니다.")
        if not msgs:
            msgs.append("지정한 [KEY]/[NOT NULL] 제약에 문제가 없습니다.")

        head = QLabel("<br><br>".join(msgs))
        head.setWordWrap(True)
        head.setStyleSheet(
            "background:{}; border:1px solid {}; padding:10px;".format(
                "#FFECEC" if blocking else "#FFF8E1",
                "#E08A8A" if blocking else "#E6C200"))
        layout.addWidget(head)

        if violations:
            table = QTableWidget(len(violations), len(self.HEADERS))
            table.setHorizontalHeaderLabels(self.HEADERS)
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.verticalHeader().setVisible(False)
            for row, (sheet, tname, issue) in enumerate(violations):
                values = (
                    sheet, tname,
                    self.KIND_LABELS.get(issue["kind"], issue["kind"]),
                    issue["detail"],
                    ", ".join(str(r) for r in issue["rows"]) + " …",
                )
                for c, v in enumerate(values):
                    item = QTableWidgetItem(str(v))
                    item.setToolTip(str(v))
                    table.setItem(row, c, item)
            header = table.horizontalHeader()
            header.setSectionResizeMode(3, QHeaderView.Stretch)
            for c in (0, 1, 2, 4):
                header.setSectionResizeMode(c, QHeaderView.ResizeToContents)
            layout.addWidget(table, 1)

        details = []
        for sheet, tname, lines in existing:
            details.append(f"· <b>{tname}</b> (시트 {sheet})<br>&nbsp;&nbsp;&nbsp;"
                           + "<br>&nbsp;&nbsp;&nbsp;".join(lines))
        for sheet, tname, msg in no_pk:
            details.append(f"· <b>{tname}</b> (시트 {sheet}) — {msg}")
        if details:
            note = QLabel("<br>".join(details))
            note.setWordWrap(True)
            note.setStyleSheet(
                "padding:8px; background:#F5F5F5; border:1px solid #DDD;")
            layout.addWidget(note, 0 if violations else 1)

        if blocking:
            buttons = QDialogButtonBox(QDialogButtonBox.Close)
            buttons.button(QDialogButtonBox.Close).setText("확인")
            buttons.rejected.connect(self.reject)
            buttons.clicked.connect(lambda *_: self.reject())
        else:
            buttons = QDialogButtonBox(
                QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            buttons.button(QDialogButtonBox.Ok).setText("확인 (이대로 진행)")
            buttons.button(QDialogButtonBox.Cancel).setText("취소")
            buttons.accepted.connect(self.accept)
            buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


# ---------------------------------------------------------------------------
# 메인 윈도우
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1180, 800)

        self.config_path = os.path.join(app_dir(), "config.ini")
        self.cfg = core.load_config(self.config_path, required=False)
        self.plans = []
        self.preview_files = []      # 미리보기 캐시가 적용된 파일 목록
        self.widgets = {}
        self.worker = None
        self.check_worker = None
        self._loading = 0        # >0 이면 위젯 시그널 처리를 건너뛴다(중첩 가능)
        self._busy = False

        self.bridge = LogBridge()
        self.bridge.message.connect(self._append_log)
        self.logger = self._make_logger()

        self._build_ui()
        self._config_to_widgets()
        self._update_buttons()
        self.logger.info("설정 파일: %s", self.config_path)

    def _make_logger(self):
        """로그 폴더를 만들 수 없는 위치(예: Program Files)면 임시 폴더로 대체한다."""
        handler = QtLogHandler(self.bridge)
        try:
            return core.setup_logging(app_dir(), self.cfg, extra_handlers=[handler])
        except OSError:
            import tempfile
            fallback = os.path.join(tempfile.gettempdir(), "ExcelUploadWizard")
            os.makedirs(fallback, exist_ok=True)
            self.cfg.set("LOGGING", "log_dir", fallback)
            logger = core.setup_logging(fallback, self.cfg, extra_handlers=[handler])
            logger.warning("실행 폴더에 로그를 쓸 수 없어 임시 폴더를 사용합니다: %s", fallback)
            return logger

    # -- UI 구성 ------------------------------------------------------------

    def _build_ui(self):
        central = QWidget()
        outer = QVBoxLayout(central)
        outer.setContentsMargins(10, 10, 10, 10)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_files_tab(), "1. 파일 선택")
        for section, fields in core.CONFIG_SPEC:
            label = core.SECTION_LABELS.get(section, section)
            self.tabs.addTab(self._build_config_tab(section, fields), f"⚙ {label}")
        self.preview_tab_index = self.tabs.count()
        self.tabs.addTab(self._build_preview_tab(), "2. 미리보기")
        self.log_tab_index = self.tabs.count()
        self.tabs.addTab(self._build_log_tab(), "3. 실행 로그")
        outer.addWidget(self.tabs, 1)

        # 하단 버튼 바
        bar = QHBoxLayout()
        self.btn_save = QPushButton("설정 저장")
        self.btn_save.clicked.connect(self.save_settings)
        bar.addWidget(self.btn_save)
        bar.addStretch(1)

        self.btn_preview = QPushButton("미리보기")
        self.btn_preview.setToolTip(
            "DRM 우회(설치된 Excel 사용)로 파일을 읽어 컬럼/유형/Comment를 확인합니다.")
        self.btn_preview.clicked.connect(self.run_preview)
        bar.addWidget(self.btn_preview)

        self.btn_check = QPushButton("점검")
        self.btn_check.setToolTip("변경한 데이터 유형으로 변환 가능한지 검사합니다.")
        self.btn_check.clicked.connect(lambda: self.run_check(silent=False))
        bar.addWidget(self.btn_check)

        self.btn_dryrun = QPushButton("모의 실행 (Dry-run)")
        self.btn_dryrun.clicked.connect(lambda: self.run_execute(dry_run=True))
        bar.addWidget(self.btn_dryrun)

        self.btn_run = QPushButton("실행")
        self.btn_run.setDefault(True)
        run_font = QFont()
        run_font.setBold(True)
        self.btn_run.setFont(run_font)
        self.btn_run.clicked.connect(lambda: self.run_execute(dry_run=False))
        bar.addWidget(self.btn_run)

        self.btn_cancel = QPushButton("중단")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_worker)
        bar.addWidget(self.btn_cancel)
        outer.addLayout(bar)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        outer.addWidget(self.progress)

        self.setCentralWidget(central)
        self.statusBar().showMessage("엑셀 파일을 선택한 뒤 [미리보기]를 실행하세요.")

    # 1. 파일 선택 ----------------------------------------------------------

    def _build_files_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        info = QLabel(
            "업로드할 엑셀 파일을 선택하세요. 여러 개를 한 번에 고를 수 있고, "
            "창으로 끌어다 놓아도 됩니다.<br>"
            "🔒 표시는 <b>DRM(문서보안)</b>이 걸린 파일입니다. "
            "이 파일은 <b>[미리보기]</b>를 먼저 실행해야 업로드할 수 있습니다."
        )
        info.setWordWrap(True)
        info.setStyleSheet("background:#EEF4FF; border:1px solid #B9CBEE; padding:10px;")
        layout.addWidget(info)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        layout.addWidget(self.file_list, 1)

        row = QHBoxLayout()
        btn_add = QPushButton("파일 추가…")
        btn_add.clicked.connect(self.add_files)
        row.addWidget(btn_add)
        btn_del = QPushButton("선택 제거")
        btn_del.clicked.connect(self.remove_selected_files)
        row.addWidget(btn_del)
        btn_clear = QPushButton("전체 제거")
        btn_clear.clicked.connect(self.clear_files)
        row.addWidget(btn_clear)
        row.addStretch(1)
        self.file_count_label = QLabel("선택된 파일 0개")
        row.addWidget(self.file_count_label)
        layout.addLayout(row)

        self.setAcceptDrops(True)
        return page

    # 2~6. 설정 탭 ----------------------------------------------------------

    def _build_config_tab(self, section, fields):
        page = QWidget()
        outer = QVBoxLayout(page)
        box = QGroupBox(f"[{section}] {core.SECTION_LABELS.get(section, section)} 설정")
        form = QFormLayout(box)
        form.setLabelAlignment(Qt.AlignRight)

        for key, label, kind, default, help_text in fields:
            widget, container = self._make_field_widget(kind, default)
            self.widgets[(section, key)] = (widget, kind)
            if help_text:
                widget.setToolTip(help_text)
                hint = QLabel(f"<span style='color:#666;font-size:11px;'>{help_text}</span>")
                hint.setWordWrap(True)
                wrapper = QWidget()
                wl = QVBoxLayout(wrapper)
                wl.setContentsMargins(0, 0, 0, 6)
                wl.setSpacing(2)
                wl.addWidget(container)
                wl.addWidget(hint)
                container = wrapper
            form.addRow(f"{label} :", container)

        outer.addWidget(box)
        outer.addStretch(1)
        return page

    def _make_field_widget(self, kind, default):
        """(값 위젯, 레이아웃에 넣을 컨테이너) 반환."""
        if kind == "bool":
            widget = QCheckBox()
            widget.setChecked(core.as_bool(default))
            return widget, widget

        if kind.startswith("choice:"):
            widget = QComboBox()
            widget.addItems(kind.split(":", 1)[1].split("|"))
            return widget, widget

        if kind == "password":
            widget = QLineEdit()
            widget.setEchoMode(QLineEdit.Password)
            container = QWidget()
            hl = QHBoxLayout(container)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.addWidget(widget, 1)
            toggle = QCheckBox("표시")
            toggle.toggled.connect(
                lambda on, w=widget: w.setEchoMode(
                    QLineEdit.Normal if on else QLineEdit.Password))
            hl.addWidget(toggle)
            return widget, container

        if kind == "dir":
            widget = QLineEdit()
            container = QWidget()
            hl = QHBoxLayout(container)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.addWidget(widget, 1)
            browse = QPushButton("찾아보기…")
            browse.clicked.connect(lambda _, w=widget: self._browse_dir(w))
            hl.addWidget(browse)
            return widget, container

        widget = QLineEdit()
        if kind == "int":
            widget.setValidator(QIntValidator(0, 10_000_000, widget))
        return widget, widget

    def _browse_dir(self, line_edit):
        start = line_edit.text().strip() or app_dir()
        path = QFileDialog.getExistingDirectory(self, "폴더 선택", start)
        if path:
            line_edit.setText(os.path.normpath(path))

    # 7. 미리보기 -----------------------------------------------------------

    def _build_preview_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        banner = QLabel(
            "[미리보기]는 설치된 Excel을 이용해 DRM 파일도 읽어옵니다. "
            "여기서 확인한 데이터가 그대로 업로드에 사용됩니다.<br>"
            "<b>지정 유형</b> 칸에서 데이터 유형을 바꾼 뒤 반드시 <b>[점검]</b>을 "
            "눌러 변환 가능 여부를 확인하세요."
        )
        banner.setWordWrap(True)
        banner.setStyleSheet("background:#EEF4FF; border:1px solid #B9CBEE; padding:8px;")
        layout.addWidget(banner)

        splitter = QSplitter(Qt.Horizontal)

        self.sheet_tree = QTreeWidget()
        self.sheet_tree.setHeaderLabels(["파일 / 시트", "테이블", "행"])
        self.sheet_tree.setColumnWidth(0, 240)
        self.sheet_tree.currentItemChanged.connect(self._on_sheet_selected)
        self.sheet_tree.itemChanged.connect(self._on_sheet_toggled)
        splitter.addWidget(self.sheet_tree)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)

        self.sheet_header = QLabel("미리보기를 실행하면 이곳에 시트 정보가 표시됩니다.")
        self.sheet_header.setWordWrap(True)
        self.sheet_header.setStyleSheet("font-weight:bold; padding:4px;")
        rl.addWidget(self.sheet_header)

        self.detail_tabs = QTabWidget()

        self.col_table = QTableWidget(0, len(COL_HEADERS))
        self.col_table.setHorizontalHeaderLabels(COL_HEADERS)
        self.col_table.verticalHeader().setVisible(False)
        self.col_table.itemChanged.connect(self._on_col_item_changed)
        ch = self.col_table.horizontalHeader()
        ch.setSectionResizeMode(C_SAMPLE, QHeaderView.Stretch)
        ch.setSectionResizeMode(C_COMMENT, QHeaderView.Interactive)
        self.col_table.setColumnWidth(C_COMMENT, 180)
        self.col_table.setColumnWidth(C_TARGET, 160)
        self.detail_tabs.addTab(self.col_table, "컬럼 / 데이터 유형")

        self.data_table = QTableWidget(0, 0)
        self.data_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.detail_tabs.addTab(self.data_table, "데이터 미리보기")

        rl.addWidget(self.detail_tabs, 1)

        self.check_status = QLabel("점검 상태: 미실행")
        self.check_status.setWordWrap(True)
        self.check_status.setStyleSheet("padding:4px; color:#555;")
        rl.addWidget(self.check_status)

        splitter.addWidget(right)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([300, 860])
        layout.addWidget(splitter, 1)
        return page

    # 8. 로그 ---------------------------------------------------------------

    def _build_log_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(20000)
        self.log_view.setStyleSheet(
            "font-family:Consolas,'D2Coding','맑은 고딕',monospace; font-size:12px;")
        layout.addWidget(self.log_view, 1)

        row = QHBoxLayout()
        row.addStretch(1)
        btn_open = QPushButton("로그 폴더 열기")
        btn_open.clicked.connect(self.open_log_dir)
        row.addWidget(btn_open)
        btn_clear = QPushButton("화면 지우기")
        btn_clear.clicked.connect(self.log_view.clear)
        row.addWidget(btn_clear)
        layout.addLayout(row)
        return page

    # -- 드래그 앤 드롭 ------------------------------------------------------

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [u.toLocalFile() for u in event.mimeData().urls()]
        excel = [p for p in paths
                 if p.lower().endswith(core.EXCEL_EXTENSIONS) and os.path.isfile(p)]
        if excel:
            self._add_file_paths(excel)
            event.acceptProposedAction()

    # -- 설정 <-> 위젯 -------------------------------------------------------

    def _config_to_widgets(self):
        self._loading += 1
        try:
            for (section, key), (widget, kind) in self.widgets.items():
                value = self.cfg.get(section, key, fallback="")
                if kind == "bool":
                    widget.setChecked(core.as_bool(value))
                elif kind.startswith("choice:"):
                    idx = widget.findText(str(value).strip(), Qt.MatchFixedString)
                    widget.setCurrentIndex(idx if idx >= 0 else 0)
                else:
                    widget.setText(str(value).strip())
        finally:
            self._loading -= 1

    def _widgets_to_config(self):
        for (section, key), (widget, kind) in self.widgets.items():
            if not self.cfg.has_section(section):
                self.cfg.add_section(section)
            if kind == "bool":
                value = "true" if widget.isChecked() else "false"
            elif kind.startswith("choice:"):
                value = widget.currentText()
            else:
                value = widget.text().strip()
            self.cfg.set(section, key, value)

    def current_options(self):
        self._widgets_to_config()
        return core.build_options(self.cfg, app_dir())

    def save_settings(self):
        self._widgets_to_config()
        try:
            core.save_config(self.cfg, self.config_path)
        except OSError as e:
            QMessageBox.critical(self, "설정 저장 실패", str(e))
            return
        self.logger = core.setup_logging(
            app_dir(), self.cfg, extra_handlers=[QtLogHandler(self.bridge)])
        self.statusBar().showMessage(f"설정을 저장했습니다: {self.config_path}", 5000)
        self.logger.info("설정 저장: %s", self.config_path)

    # -- 파일 목록 -----------------------------------------------------------

    def add_files(self):
        start = self.current_options().get("excel_dir") or app_dir()
        if not os.path.isdir(start):
            start = app_dir()
        paths, _ = QFileDialog.getOpenFileNames(
            self, "업로드할 엑셀 파일 선택", start,
            "Excel 파일 (*.xlsx *.xlsm *.xls);;모든 파일 (*.*)")
        if paths:
            self._add_file_paths(paths)

    def _add_file_paths(self, paths):
        existing = set(self.selected_files())
        added = 0
        for path in paths:
            path = os.path.abspath(path)
            if path in existing or not os.path.isfile(path):
                continue
            drm = core.is_drmone_file(path)
            prefix = "🔒 " if drm else "     "
            item = QListWidgetItem(f"{prefix}{os.path.basename(path)}")
            item.setData(Qt.UserRole, path)
            item.setToolTip(path + ("\n[DRM 보호 파일] 미리보기 필요" if drm else ""))
            if drm:
                item.setForeground(QColor("#B34700"))
            self.file_list.addItem(item)
            existing.add(path)
            added += 1
        if added:
            self._invalidate_preview("파일 목록이 변경되었습니다.")
        self._refresh_file_count()

    def remove_selected_files(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
        self._invalidate_preview("파일 목록이 변경되었습니다.")
        self._refresh_file_count()

    def clear_files(self):
        self.file_list.clear()
        self._invalidate_preview("파일 목록이 비워졌습니다.")
        self._refresh_file_count()

    def selected_files(self):
        return [self.file_list.item(i).data(Qt.UserRole)
                for i in range(self.file_list.count())]

    def _refresh_file_count(self):
        files = self.selected_files()
        drm = sum(1 for f in files if core.is_drmone_file(f))
        text = f"선택된 파일 {len(files)}개"
        if drm:
            text += f"  (DRM {drm}개 — 미리보기 필요)"
        self.file_count_label.setText(text)
        self._update_buttons()

    # -- 미리보기 -----------------------------------------------------------

    def _invalidate_preview(self, reason=""):
        if self.plans:
            self.logger.info("미리보기 캐시를 초기화합니다. %s", reason)
        self.plans = []
        self.preview_files = []
        self.sheet_tree.clear()
        self.col_table.setRowCount(0)
        self.data_table.setRowCount(0)
        self.data_table.setColumnCount(0)
        self.sheet_header.setText("미리보기를 실행하면 이곳에 시트 정보가 표시됩니다.")
        self._refresh_check_status()
        self._update_buttons()

    def run_preview(self):
        files = self.selected_files()
        if not files:
            QMessageBox.information(self, "파일 없음",
                                    "먼저 [1. 파일 선택] 탭에서 엑셀 파일을 선택하세요.")
            return
        options = self.current_options()
        self._invalidate_preview("미리보기를 다시 실행합니다.")
        self.tabs.setCurrentIndex(self.log_tab_index)
        self._set_busy(True, "미리보기 준비 중…")
        self.logger.info("=" * 60)
        self.logger.info("미리보기 시작 (파일 %d개)", len(files))

        self.worker = PreviewWorker(files, options, self.logger, self)
        self.worker.progressed.connect(
            lambda msg: self.statusBar().showMessage(msg))
        self.worker.finished_ok.connect(self._on_preview_done)
        self.worker.failed.connect(self._on_worker_failed)
        self.worker.finished.connect(lambda: self._set_busy(False))
        self.worker.start()

    def _on_preview_done(self, plans, errors):
        self.plans = plans
        self.preview_files = self.selected_files()
        self._populate_sheet_tree()
        for fname, sheet, msg in errors:
            self.logger.warning("건너뜀: %s %s -> %s", fname, sheet or "-", msg)
        drm_count = sum(1 for p in plans if p.get("is_drm"))
        msg = f"미리보기 완료: 시트 {len(plans)}개"
        if drm_count:
            msg += f" (DRM 우회로 읽은 시트 {drm_count}개)"
        self.logger.info(msg)
        self.tabs.setCurrentIndex(self.preview_tab_index)
        self.statusBar().showMessage(msg, 8000)
        if errors:
            detail = "\n".join(f"· {f} {s or ''} : {m}" for f, s, m in errors)
            QMessageBox.warning(self, "일부 항목을 건너뛰었습니다", detail)
        if not plans:
            QMessageBox.information(
                self, "대상 시트 없음",
                "선택한 파일에서 처리 대상 시트를 찾지 못했습니다.\n"
                "[⚙ 일반] 탭의 '시트 이름 키워드' 설정을 확인하세요.")
        self._update_buttons()

    def _populate_sheet_tree(self):
        self._loading += 1
        try:
            self.sheet_tree.clear()
            by_file = {}
            for idx, plan in enumerate(self.plans):
                by_file.setdefault(plan["file_name"], []).append((idx, plan))
            for fname, entries in by_file.items():
                drm = any(p.get("is_drm") for _, p in entries)
                root = QTreeWidgetItem([("🔒 " if drm else "") + fname, "", ""])
                root.setFlags(root.flags() & ~Qt.ItemIsUserCheckable)
                self.sheet_tree.addTopLevelItem(root)
                for idx, plan in entries:
                    node = QTreeWidgetItem([
                        plan["sheet_name"], plan["table_name"],
                        f"{plan['row_count']:,}"])
                    node.setFlags(node.flags() | Qt.ItemIsUserCheckable)
                    node.setCheckState(0, Qt.Checked if plan["enabled"] else Qt.Unchecked)
                    node.setData(0, Qt.UserRole, idx)
                    root.addChild(node)
                root.setExpanded(True)
            first = self._first_sheet_item()
            if first:
                self.sheet_tree.setCurrentItem(first)
        finally:
            self._loading -= 1
        if self.sheet_tree.topLevelItemCount():
            self._show_plan(self._current_plan_index())

    def _first_sheet_item(self):
        for i in range(self.sheet_tree.topLevelItemCount()):
            root = self.sheet_tree.topLevelItem(i)
            if root.childCount():
                return root.child(0)
        return None

    def _current_plan_index(self):
        item = self.sheet_tree.currentItem()
        if item is None:
            return None
        value = item.data(0, Qt.UserRole)
        return value if isinstance(value, int) else None

    def _on_sheet_selected(self, current, _previous):
        self._show_plan(self._current_plan_index())

    def _on_sheet_toggled(self, item, column):
        if self._loading or column != 0:
            return
        idx = item.data(0, Qt.UserRole)
        if isinstance(idx, int) and 0 <= idx < len(self.plans):
            self.plans[idx]["enabled"] = item.checkState(0) == Qt.Checked
            self._update_buttons()

    def _show_plan(self, index):
        if index is None or not (0 <= index < len(self.plans)):
            self.col_table.setRowCount(0)
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            return
        plan = self.plans[index]
        header_note = ("자동 감지" if plan["auto_header_used"] else "설정값")
        self.sheet_header.setText(
            f"{plan['file_name']}  ›  시트 '{plan['sheet_name']}'  →  "
            f"테이블 <span style='color:#0B5FFF'>{plan['table_name']}</span>"
            f"   |   행 {plan['row_count']:,} · 컬럼 {len(plan['columns'])} · "
            f"헤더 행 {plan['header_row']}({header_note})"
            + ("   |   <span style='color:#B34700'>DRM 우회로 읽음</span>"
               if plan.get("is_drm") else ""))
        self._fill_column_table(plan)
        self._fill_data_table(plan)
        self._refresh_check_status()

    def _fill_column_table(self, plan):
        self._loading += 1
        try:
            df = plan["df"]
            cols = plan["columns"]
            self.col_table.setRowCount(len(cols))
            for row, col in enumerate(cols):
                self._check_item(row, C_USE, col.get("include", True))
                self._ro_item(row, C_NAME, col["name"])
                self._ro_item(row, C_ORIG, col["orig_name"])
                self._ro_item(row, C_INFER, col["inferred_type"])

                combo = QComboBox()
                combo.setEditable(True)
                choices = list(core.HANA_TYPE_CHOICES)
                for t in (col["inferred_type"], col["target_type"]):
                    if t and t not in choices:
                        choices.insert(0, t)
                combo.addItems(choices)
                combo.setCurrentText(col["target_type"])
                combo.currentTextChanged.connect(
                    lambda text, r=row: self._on_type_changed(r, text))
                self.col_table.setCellWidget(row, C_TARGET, combo)

                key_item = self._check_item(row, C_KEY, col.get("is_key", False))
                key_item.setToolTip(
                    "체크한 컬럼들이 PRIMARY KEY가 됩니다.\n"
                    "테이블을 새로 만들 때만 적용되며, KEY 컬럼은 자동으로 "
                    "NOT NULL이 됩니다.\nUPSERT 모드에는 KEY 지정이 필요합니다.")
                nn_item = self._check_item(row, C_NOTNULL, col.get("not_null", False))
                nn_item.setToolTip(
                    "체크한 컬럼에 NOT NULL 제약을 붙입니다.\n"
                    "테이블을 새로 만들 때만 적용됩니다.")

                comment = QTableWidgetItem(col.get("comment", ""))
                comment.setToolTip("셀을 더블클릭해 컬럼 COMMENT를 수정할 수 있습니다.")
                self.col_table.setItem(row, C_COMMENT, comment)

                self._ro_item(row, C_NULLCNT, f"{col.get('null_count', 0):,}")

                series = df[col["name"]].dropna()
                sample = ", ".join(short_text(v, 18)
                                   for v in series.head(SAMPLE_LIMIT).tolist())
                self._ro_item(row, C_SAMPLE, sample)
                self._paint_type_row(row, col)
            self.col_table.resizeColumnsToContents()
            self.col_table.setColumnWidth(
                C_TARGET, max(170, self.col_table.columnWidth(C_TARGET)))
            self.col_table.setColumnWidth(
                C_COMMENT, max(180, self.col_table.columnWidth(C_COMMENT)))
        finally:
            self._loading -= 1

    def _ro_item(self, row, col, text):
        item = QTableWidgetItem(str(text))
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item.setToolTip(str(text))
        self.col_table.setItem(row, col, item)
        return item

    def _check_item(self, row, col, checked):
        item = QTableWidgetItem()
        item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        self.col_table.setItem(row, col, item)
        return item

    def _paint_type_row(self, row, col):
        changed = (col["target_type"] or "").strip().upper() != \
                  (col["inferred_type"] or "").strip().upper()
        result = col.get("check_result")
        if result and result.get("bad_count"):
            color = QColor("#FFE0E0")
        elif changed and not col.get("checked"):
            color = QColor("#FFF6D6")
        elif changed:
            color = QColor("#E4F6E4")
        else:
            color = QColor(Qt.white)
        for c in range(self.col_table.columnCount()):
            item = self.col_table.item(row, c)
            if item is not None:
                item.setBackground(color)

    def _on_type_changed(self, row, text):
        if self._loading:
            return
        index = self._current_plan_index()
        if index is None:
            return
        plan = self.plans[index]
        if not (0 <= row < len(plan["columns"])):
            return
        col = plan["columns"][row]
        col["target_type"] = text.strip().upper()
        col["checked"] = False
        col["check_result"] = None
        plan["checked"] = False
        self._paint_type_row(row, col)
        self._refresh_check_status()
        self._update_buttons()

    def _on_col_item_changed(self, item):
        if self._loading or item is None:
            return
        index = self._current_plan_index()
        if index is None:
            return
        plan = self.plans[index]
        row = item.row()
        if not (0 <= row < len(plan["columns"])):
            return
        col = plan["columns"][row]
        column = item.column()
        checked = item.checkState() == Qt.Checked
        if column == C_USE:
            col["include"] = checked
        elif column == C_COMMENT:
            col["comment"] = item.text().strip()
        elif column == C_KEY:
            col["is_key"] = checked
            # PRIMARY KEY는 HANA에서 암묵적으로 NOT NULL이다. 화면에도 반영한다.
            if checked:
                col["not_null"] = True
                nn = self.col_table.item(row, C_NOTNULL)
                if nn is not None and nn.checkState() != Qt.Checked:
                    self._loading += 1
                    try:
                        nn.setCheckState(Qt.Checked)
                    finally:
                        self._loading -= 1
            plan["checked"] = False
            self._refresh_check_status()
            self._update_buttons()
        elif column == C_NOTNULL:
            if not checked and col.get("is_key"):
                # KEY 컬럼은 NOT NULL을 해제할 수 없다. 되돌린다.
                self._loading += 1
                try:
                    item.setCheckState(Qt.Checked)
                finally:
                    self._loading -= 1
                self.statusBar().showMessage(
                    "KEY로 지정한 컬럼은 NOT NULL을 해제할 수 없습니다.", 5000)
                return
            col["not_null"] = checked
            plan["checked"] = False
            self._refresh_check_status()
            self._update_buttons()

    def _fill_data_table(self, plan):
        df = plan["df"]
        limit = self.current_options()["preview_rows"]
        head = df.head(limit)
        self.data_table.clear()
        self.data_table.setColumnCount(len(head.columns))
        self.data_table.setRowCount(len(head))
        self.data_table.setHorizontalHeaderLabels([str(c) for c in head.columns])
        for r in range(len(head)):
            for c, name in enumerate(head.columns):
                value = head.iloc[r, c]
                text = "" if core._is_null(value) else short_text(value, 60)
                self.data_table.setItem(r, c, QTableWidgetItem(text))
        self.data_table.resizeColumnsToContents()
        self.detail_tabs.setTabText(
            1, f"데이터 미리보기 ({len(head):,} / {len(df):,}행)")

    # -- 점검 ---------------------------------------------------------------

    def _changed_columns(self):
        """유형이 변경된(=검증이 필요한) 컬럼 목록."""
        result = []
        for plan in self.plans:
            if not plan.get("enabled", True):
                continue
            for col in plan["columns"]:
                if not col.get("include", True):
                    continue
                target = (col.get("target_type") or "").strip().upper()
                inferred = (col.get("inferred_type") or "").strip().upper()
                if target and target != inferred:
                    result.append((plan, col))
        return result

    def _pending_check(self):
        return [(p, c) for p, c in self._changed_columns() if not c.get("checked")]

    def _plans_with_constraints(self):
        """[KEY] 또는 [NOT NULL]을 지정한 시트 목록."""
        out = []
        for plan in self.plans:
            if not plan.get("enabled", True):
                continue
            cols = [c for c in plan["columns"] if c.get("include", True)]
            if any(c.get("is_key") or c.get("not_null") for c in cols):
                out.append(plan)
        return out

    def _needs_check(self):
        if self._pending_check():
            return True
        return any(not p.get("checked") for p in self._plans_with_constraints())

    def _collect_constraint_violations(self):
        """지정한 KEY/NOT NULL이 실제 데이터와 맞는지 검사한다."""
        violations = []
        for plan in self.plans:
            if not plan.get("enabled", True):
                continue
            for issue in core.check_constraints(plan["df"], plan["columns"]):
                violations.append(
                    (plan["sheet_name"], plan["table_name"], issue))
        return violations

    def _constraint_summary(self):
        index = self._current_plan_index()
        if index is None or not (0 <= index < len(self.plans)):
            return ""
        cols = [c for c in self.plans[index]["columns"] if c.get("include", True)]
        keys = [c["name"] for c in cols if c.get("is_key")]
        nn = [c["name"] for c in cols if c.get("not_null") and not c.get("is_key")]
        parts = []
        if keys:
            parts.append("PRIMARY KEY: " + ", ".join(keys))
        if nn:
            parts.append("NOT NULL: " + ", ".join(nn))
        return "  |  ".join(parts)

    def _refresh_check_status(self):
        changed = self._changed_columns()
        constraints = self._constraint_summary()
        suffix = f"<br>{constraints}" if constraints else ""
        if not changed:
            base = ("점검 상태: 변경된 데이터 유형이 없습니다. "
                    "자동 추론 유형 그대로 업로드합니다.")
            if constraints:
                pending_c = [p for p in self._plans_with_constraints()
                             if not p.get("checked")]
                if pending_c:
                    base = ("점검 상태: [KEY]/[NOT NULL]을 지정했습니다. "
                            "[점검]으로 데이터와 기존 테이블을 확인하세요.")
                    self.check_status.setText(base + suffix)
                    self.check_status.setStyleSheet(
                        "padding:4px; color:#8A6D00; background:#FFF8E1;")
                    return
            self.check_status.setText(base + suffix)
            self.check_status.setStyleSheet("padding:4px; color:#555;")
            return
        pending = self._pending_check()
        bad = [(p, c) for p, c in changed
               if (c.get("check_result") or {}).get("bad_count")]
        if pending:
            self.check_status.setText(
                f"점검 상태: 유형을 변경한 컬럼 {len(changed)}개 중 "
                f"{len(pending)}개가 아직 점검되지 않았습니다. [점검]을 눌러 주세요."
                + suffix)
            self.check_status.setStyleSheet(
                "padding:4px; color:#8A6D00; background:#FFF8E1;")
        elif bad:
            detail = ", ".join(
                f"{c['name']}({core.POLICY_LABELS.get(c.get('policy'), '')})"
                for _, c in bad)
            self.check_status.setText(
                f"점검 완료: 불일치가 있는 컬럼 {len(bad)}개 — {detail}" + suffix)
            self.check_status.setStyleSheet(
                "padding:4px; color:#8A2A00; background:#FFECEC;")
        else:
            self.check_status.setText(
                f"점검 완료: 변경한 유형 {len(changed)}개 모두 정상 변환 가능합니다."
                + suffix)
            self.check_status.setStyleSheet(
                "padding:4px; color:#1B5E20; background:#E8F5E9;")

    def run_check(self, silent=False):
        """유형 변경 컬럼을 검사하고, 불일치가 있으면 처리 방식을 묻는다.

        반환: True(계속 진행 가능) / False(사용자가 취소했거나 유형 표기 오류)
        """
        if not self.plans:
            if not silent:
                QMessageBox.information(self, "미리보기 필요",
                                        "먼저 [미리보기]를 실행하세요.")
            return False

        # --- KEY / NOT NULL 제약 검사 (지정한 데이터와 맞는지) --------------
        constrained = self._plans_with_constraints()
        violations = self._collect_constraint_violations()
        if violations:
            for sheet, tname, issue in violations:
                self.logger.error("[점검] %s -> %s : %s",
                                  sheet, tname, issue["detail"])
            ConstraintCheckDialog(violations, None, self).exec_()
            self._refresh_preview_paint()
            return False

        changed = self._changed_columns()
        if not changed and not constrained:
            if not silent:
                QMessageBox.information(
                    self, "점검 결과",
                    "변경된 데이터 유형이 없고 지정한 제약도 없습니다.\n"
                    "자동 추론된 유형 그대로 업로드합니다.")
            return True

        if not changed:
            # 유형 변경은 없고 제약만 지정한 경우
            for plan in self.plans:
                plan["checked"] = True
            self._refresh_preview_paint()
            if not silent:
                self._check_existing_tables_async(constrained)
            return True

        self.statusBar().showMessage("데이터 유형 점검 중…")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        issues, invalid = [], []
        try:
            for plan, col in changed:
                result = core.check_series(plan["df"][col["name"]],
                                           col["target_type"])
                col["check_result"] = result
                col["checked"] = True
                if result["error"]:
                    invalid.append((plan, col, result["error"]))
                elif result["bad_count"]:
                    sample_text = ", ".join(
                        f"{row}행: {short_text(v, 20)}" for row, v in result["samples"])
                    issues.append({
                        "plan": plan, "column_ref": col,
                        "file_name": plan["file_name"],
                        "sheet_name": plan["sheet_name"],
                        "column": col["name"],
                        "target_type": col["target_type"],
                        "bad_count": result["bad_count"],
                        "total": result["total"],
                        "sample_text": sample_text,
                        "policy": col.get("policy", core.POLICY_NULL),
                    })
        finally:
            QApplication.restoreOverrideCursor()
            self.statusBar().clearMessage()

        if invalid:
            detail = "\n".join(f"· {p['sheet_name']}.{c['name']} : {msg}"
                               for p, c, msg in invalid)
            QMessageBox.critical(
                self, "데이터 유형 표기 오류",
                "아래 컬럼의 데이터 유형 표기를 해석할 수 없습니다.\n"
                "예: INTEGER, BIGINT, DECIMAL(18,4), NVARCHAR(100), DATE, TIMESTAMP\n\n"
                + detail)
            self._refresh_preview_paint()
            return False

        if not issues:
            for plan in self.plans:
                plan["checked"] = True
            self._refresh_preview_paint()
            if not silent:
                if constrained:
                    self._check_existing_tables_async(constrained)
                else:
                    QMessageBox.information(
                        self, "점검 결과",
                        f"변경한 데이터 유형 {len(changed)}개 모두 정상 변환 "
                        "가능합니다.\n그대로 업로드할 수 있습니다.")
            return True

        dialog = CheckDialog(issues, self)
        if dialog.exec_() != QDialog.Accepted:
            self._refresh_preview_paint()
            self.statusBar().showMessage("점검 결과 적용을 취소했습니다.", 5000)
            return False

        for issue, policy in zip(issues, dialog.policies()):
            issue["column_ref"]["policy"] = policy
            self.logger.info(
                "[점검] %s.%s -> %s : 불일치 %d건, 처리 방식 = %s",
                issue["sheet_name"], issue["column"], issue["target_type"],
                issue["bad_count"], core.POLICY_LABELS[policy])
        for plan in self.plans:
            plan["checked"] = True
        self._refresh_preview_paint()
        if not silent and constrained:
            self._check_existing_tables_async(constrained)
        return True

    def _check_existing_tables_async(self, constrained):
        """[KEY]/[NOT NULL]을 지정했을 때, 대상 테이블이 이미 있는지 확인한다.

        기존 테이블에는 제약이 적용되지 않으므로 그 사실을 경고해야 한다.
        메타데이터 조회만 하지만 접속이 필요하므로 워커 스레드에서 돌린다.
        """
        options = self.current_options()
        if not (self.cfg.get("HANA", "host", fallback="").strip()
                and options.get("schema")):
            QMessageBox.information(
                self, "점검 결과",
                "데이터 기준으로는 문제가 없습니다.\n\n"
                "다만 [KEY]/[NOT NULL]은 테이블을 새로 만들 때만 적용되므로, "
                "대상 테이블이 이미 있는지 확인해야 합니다. HANA 접속 정보와 "
                "스키마를 입력하면 점검 시 자동으로 확인해 드립니다.")
            return

        self._set_busy(True, "대상 테이블 확인 중…")
        self.check_worker = PreflightWorker(constrained, self.cfg, options,
                                           self.logger, parent=self)
        self.check_worker.finished_ok.connect(self._on_constraint_check_done)
        self.check_worker.failed.connect(self._on_worker_failed)
        self.check_worker.finished.connect(lambda: self._set_busy(False))
        self.check_worker.start()

    def _on_constraint_check_done(self, report):
        if report.get("error"):
            QMessageBox.warning(
                self, "테이블 확인 실패",
                "대상 테이블이 이미 있는지 확인하지 못했습니다.\n\n"
                f"{report['error']}\n\n"
                "[KEY]/[NOT NULL]은 테이블을 새로 만들 때만 적용된다는 점을 "
                "유의하세요.")
            return
        for sheet, table, lines in report["constraints_ignored"]:
            self.logger.warning("[점검] '%s'는 이미 존재하는 테이블입니다. "
                                "KEY/NOT NULL 지정이 적용되지 않습니다.", table)
            for line in lines:
                self.logger.warning("        %s", line)
        for sheet, table, msg in report["no_primary_key"]:
            self.logger.error("[점검] %s : %s", table, msg)
        ConstraintCheckDialog([], report, self).exec_()

    def _refresh_preview_paint(self):
        index = self._current_plan_index()
        if index is not None and 0 <= index < len(self.plans):
            plan = self.plans[index]
            self._loading += 1
            try:
                for row, col in enumerate(plan["columns"]):
                    if row < self.col_table.rowCount():
                        self._paint_type_row(row, col)
            finally:
                self._loading -= 1
        self._refresh_check_status()
        self._update_buttons()

    # -- 실행 ---------------------------------------------------------------

    def run_execute(self, dry_run=False):
        files = self.selected_files()
        if not files:
            QMessageBox.information(self, "파일 없음",
                                    "먼저 [1. 파일 선택] 탭에서 엑셀 파일을 선택하세요.")
            return

        options = self.current_options()
        preview_valid = bool(self.plans) and self.preview_files == files

        # --- DRM 검사: 미리보기 캐시가 없으면 전체 실행 중단 ---------------
        if not preview_valid:
            drm_files = [f for f in files if core.is_drmone_file(f)]
            if drm_files:
                names = "\n".join(f"  · {os.path.basename(f)}" for f in drm_files)
                QMessageBox.warning(
                    self, "DRM 파일이 포함되어 있습니다",
                    "선택한 파일 중 DRM(문서보안)이 걸린 파일이 있어 "
                    "실행을 중단했습니다.\n\n"
                    f"{names}\n\n"
                    + core.DRM_GUIDE_MESSAGE.format(
                        name=os.path.basename(drm_files[0])))
                self.logger.error(
                    "DRM 파일이 포함되어 실행을 중단했습니다: %s",
                    ", ".join(os.path.basename(f) for f in drm_files))
                self.statusBar().showMessage(
                    "DRM 파일이 포함되어 실행을 중단했습니다. [미리보기]를 먼저 실행하세요.",
                    10000)
                return

        # --- 미리보기 캐시가 없으면 지금 읽는다 (DRM 없음이 확인된 상태) ---
        if not preview_valid:
            self._set_busy(True, "파일 읽는 중…")
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                plans, errors = core.collect_sheet_plans(
                    files, options, self.logger, allow_drm=False)
            except core.DrmProtectedError as e:
                QMessageBox.warning(self, "DRM 파일이 포함되어 있습니다", str(e))
                return
            except Exception as e:
                QMessageBox.critical(self, "파일 읽기 실패", str(e))
                return
            finally:
                QApplication.restoreOverrideCursor()
                self._set_busy(False)
            self.plans = plans
            self.preview_files = files
            self._populate_sheet_tree()
            for fname, sheet, msg in errors:
                self.logger.warning("건너뜀: %s %s -> %s", fname, sheet or "-", msg)

        active = [p for p in self.plans if p.get("enabled", True)]
        if not active:
            QMessageBox.information(self, "대상 없음", "업로드할 시트가 없습니다.")
            return

        # --- 유형을 바꿨는데 점검하지 않았으면 지금 점검 --------------------
        if self._needs_check():
            answer = QMessageBox.question(
                self, "점검이 필요합니다",
                "데이터 유형을 변경했거나 [KEY]/[NOT NULL]을 지정한 항목 중 "
                "아직 점검하지 않은 것이 있습니다.\n"
                "지금 점검을 실행할까요?\n\n"
                "(점검하지 않으면 유형에 맞지 않는 값은 기본 처리 방식인 "
                "'NULL로 변환 후 진행'이 적용됩니다.)",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                QMessageBox.Yes)
            if answer == QMessageBox.Cancel:
                return
            if answer == QMessageBox.Yes and not self.run_check(silent=True):
                return

        # --- 실행 전 확인 ---------------------------------------------------
        rows = sum(p["row_count"] for p in active)
        mode = "모의 실행(Dry-run)" if dry_run else "실제 업로드"
        lines = "\n".join(
            f"  · {p['file_name']} / {p['sheet_name']}  →  {p['table_name']}"
            f"  ({p['row_count']:,}행)" for p in active[:15])
        if len(active) > 15:
            lines += f"\n  … 외 {len(active) - 15}개"
        confirm = QMessageBox.question(
            self, f"{mode} 확인",
            f"{mode}을(를) 시작합니다.\n\n"
            f"스키마: {options['schema'] or '(미지정)'}\n"
            f"적재 모드: "
            f"{core.LOAD_MODE_LABELS.get(options['load_mode'], options['load_mode'])}\n"
            f"대상: 시트 {len(active)}개 / 총 {rows:,}행\n\n{lines}",
            QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Ok)
        if confirm != QMessageBox.Ok:
            return

        self.tabs.setCurrentIndex(self.log_tab_index)

        # 모의 실행은 DB에 접속하지 않으므로 사전 확인을 건너뛴다.
        if dry_run:
            self._start_upload(options, active, mode, dry_run=True)
            return

        # 실제 업로드 전에 대상 테이블 구조를 먼저 확인한다.
        self._set_busy(True, "대상 테이블 확인 중…")
        self.logger.info("=" * 60)
        self.logger.info("업로드 전 대상 테이블 확인")
        self.worker = PreflightWorker(self.plans, self.cfg, options,
                                      self.logger, parent=self)
        self.worker.finished_ok.connect(
            lambda rep: self._on_preflight_done(rep, options, active, mode))
        self.worker.failed.connect(self._on_worker_failed)
        self.worker.finished.connect(lambda: self._set_busy(False))
        self.worker.start()

    def _on_preflight_done(self, report, options, active, mode):
        if report.get("error"):
            answer = QMessageBox.question(
                self, "테이블 확인 실패",
                "업로드 전 테이블 구조를 확인할 수 없었습니다.\n\n"
                f"{report['error']}\n\n"
                "그래도 업로드를 계속할까요?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes:
                self.statusBar().showMessage("업로드를 취소했습니다.", 5000)
                return
        else:
            for sheet, table, n in report["new"]:
                self.logger.info("   - 신규 생성 예정: %s (컬럼 %d개)", table, n)
            for sheet, table, cols in report["missing_columns"]:
                self.logger.warning("   - '%s'에 없는 컬럼: %s", table, ", ".join(cols))
            for sheet, table, lines in report["constraints_ignored"]:
                self.logger.warning("   - '%s'는 기존 테이블이므로 "
                                    "[KEY]/[NOT NULL] 지정이 적용되지 않습니다.",
                                    table)
                for line in lines:
                    self.logger.warning("       %s", line)
            for sheet, table, msg in report["no_primary_key"]:
                self.logger.error("   - %s : %s", table, msg)
            if report["type_mismatch"]:
                self.logger.warning(
                    "   - 기존 테이블 유형 불일치 %d건 (테이블 유형이 사용됩니다)",
                    len(report["type_mismatch"]))
                for sheet, table, col, want, have in report["type_mismatch"]:
                    self.logger.warning("       %s.%s : 지정 %s / 실제 %s",
                                        table, col, want, have)

            # 유형 불일치가 있으면 비교 표, 없고 제약 문제만 있으면 제약 경고
            if report["type_mismatch"]:
                if TableCheckDialog(report, self).exec_() != QDialog.Accepted:
                    self.logger.info("사용자가 업로드를 취소했습니다.")
                    self.statusBar().showMessage("업로드를 취소했습니다.", 5000)
                    return
            elif report["constraints_ignored"] or report["no_primary_key"]:
                if ConstraintCheckDialog([], report, self).exec_() != QDialog.Accepted:
                    self.logger.info("사용자가 업로드를 취소했습니다.")
                    self.statusBar().showMessage("업로드를 취소했습니다.", 5000)
                    return

        self._start_upload(options, active, mode, dry_run=False)

    def _start_upload(self, options, active, mode, dry_run):
        rows = sum(p["row_count"] for p in active)
        self._set_busy(True, f"{mode} 진행 중…")
        self.progress.setRange(0, len(active))
        self.progress.setValue(0)
        self.logger.info("=" * 60)
        self.logger.info("%s 시작 (시트 %d개, %d행)", mode, len(active), rows)

        self.worker = ExecuteWorker(self.plans, self.cfg, options, self.logger,
                                    dry_run=dry_run, parent=self)
        self.worker.progressed.connect(self._on_execute_progress)
        self.worker.finished_ok.connect(
            lambda s, e: self._on_execute_done(s, e, dry_run))
        self.worker.failed.connect(self._on_worker_failed)
        self.worker.finished.connect(lambda: self._set_busy(False))
        self.worker.start()

    def _on_execute_progress(self, index, total, message):
        self.progress.setRange(0, total)
        self.progress.setValue(index)
        self.statusBar().showMessage(f"[{index}/{total}] {message}")

    def _on_execute_done(self, summary, had_error, dry_run):
        ok = [s for s in summary if not str(s[4]).startswith("실패")]
        fail = [s for s in summary if str(s[4]).startswith("실패")]
        total_rows = sum(s[3] for s in ok)
        mode = "모의 실행" if dry_run else "업로드"
        text = (f"{mode} 완료\n\n"
                f"성공 시트: {len(ok)}개 / 총 {total_rows:,}행\n"
                f"실패 시트: {len(fail)}개")
        if fail:
            text += "\n\n실패 내역:\n" + "\n".join(
                f"  · {s[1]} → {s[2]} : {s[4]}" for s in fail[:10])
            QMessageBox.warning(self, f"{mode} 완료 (오류 있음)", text)
        else:
            QMessageBox.information(self, f"{mode} 완료", text)
        self.statusBar().showMessage(
            f"{mode} 완료 — 성공 {len(ok)} / 실패 {len(fail)}", 10000)

    def _on_worker_failed(self, title, message):
        QMessageBox.critical(self, title, message)
        self.statusBar().showMessage(title, 8000)

    def cancel_worker(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.statusBar().showMessage("중단 요청됨 — 현재 작업이 끝나면 멈춥니다.")

    # -- 공통 ---------------------------------------------------------------

    def _set_busy(self, busy, message=""):
        self._busy = busy
        for btn in (self.btn_preview, self.btn_check, self.btn_run,
                    self.btn_dryrun, self.btn_save):
            btn.setEnabled(not busy)
        self.btn_cancel.setEnabled(busy)
        self.progress.setVisible(busy)
        if busy:
            self.progress.setRange(0, 0)
            if message:
                self.statusBar().showMessage(message)
        else:
            self.progress.setRange(0, 1)
            self.progress.setValue(1)
            self._update_buttons()

    def _update_buttons(self):
        if self._busy:
            return
        has_files = self.file_list.count() > 0
        self.btn_preview.setEnabled(has_files)
        self.btn_run.setEnabled(has_files)
        self.btn_dryrun.setEnabled(has_files)
        pending = self._pending_check()
        pending_c = [p for p in self._plans_with_constraints()
                     if not p.get("checked")]
        self.btn_check.setEnabled(bool(self.plans))
        if pending or pending_c:
            count = len(pending) + len(pending_c)
            self.btn_check.setText(f"점검 ({count})")
            self.btn_check.setStyleSheet(
                "background:#FFC107; font-weight:bold; padding:4px 12px;")
        else:
            self.btn_check.setText("점검")
            self.btn_check.setStyleSheet("")

    def _append_log(self, text, levelno):
        if levelno >= logging.ERROR:
            html = f"<span style='color:#C62828'>{_escape(text)}</span>"
        elif levelno >= logging.WARNING:
            html = f"<span style='color:#B36A00'>{_escape(text)}</span>"
        else:
            html = _escape(text)
        self.log_view.appendHtml(html)

    def open_log_dir(self):
        self._widgets_to_config()
        log_dir = self.cfg.get("LOGGING", "log_dir", fallback="logs").strip() or "logs"
        if not os.path.isabs(log_dir):
            log_dir = os.path.join(app_dir(), log_dir)
        os.makedirs(log_dir, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(log_dir)  # noqa: S606
            elif sys.platform == "darwin":
                os.system(f'open "{log_dir}"')
            else:
                os.system(f'xdg-open "{log_dir}"')
        except Exception as e:
            QMessageBox.information(self, "로그 폴더", f"{log_dir}\n\n({e})")

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            answer = QMessageBox.question(
                self, "작업 진행 중",
                "작업이 진행 중입니다. 정말 종료할까요?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes:
                event.ignore()
                return
            self.worker.cancel()
            self.worker.wait(3000)
        event.accept()


def _escape(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


# ---------------------------------------------------------------------------

def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    if sys.platform == "win32":
        app.setFont(QFont("맑은 고딕", 9))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
