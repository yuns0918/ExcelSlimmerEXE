import sys
import threading
from pathlib import Path

from PySide6.QtCore import Qt, QObject, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QPlainTextEdit,
    QSizePolicy,
    QTabWidget,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)


def _ensure_module_paths() -> None:
    base = Path(__file__).resolve().parent
    root = base.parent
    for name in ("ExcelCleaner", "ExcelImageOptimization", "ExcelByteReduce"):
        p = root / name
        if p.is_dir():
            sys.path.insert(0, str(p))


_ensure_module_paths()

from gui_clean_defined_names_desktop_date import process_file_gui
from excel_image_slimmer_gui_v3 import slim_xlsx, human_size, open_in_explorer_select
from excel_slimmer_precision_plus import process_file as precision_process, Progress


def run_image_slim(input_path: Path, max_edge: int, jpeg_quality: int, progressive: bool):
    base_out = input_path.with_stem(input_path.stem + "_slim")
    out_path = base_out
    idx = 1
    while out_path.exists():
        out_path = input_path.with_stem(input_path.stem + f"_slim({idx})")
        idx += 1
    log_path = input_path.with_name(input_path.stem + "_image_slim.log")
    before, after, count = slim_xlsx(
        input_path,
        out_path,
        max_edge,
        jpeg_quality,
        progressive,
        log_path,
        ui=None,
    )
    return out_path, before, after, count, log_path


def run_precision_step(
    input_path: Path,
    aggressive: bool,
    no_backup: bool,
    do_xml_cleanup: bool,
    force_custom: bool,
    logger,
):
    overall = Progress(None, None)
    file_prog = Progress(None, None)
    summary = {"files": [], "saved_bytes": 0, "original_bytes": 0}
    precision_process(
        input_path,
        aggressive,
        no_backup,
        do_xml_cleanup,
        force_custom,
        logger,
        overall,
        file_prog,
        summary,
    )
    if summary["files"]:
        _, outname, old_b, new_b, saved_mb, pct = summary["files"][-1]
        out_path = input_path.with_name(outname)
        return out_path, saved_mb, pct, old_b, new_b
    size = input_path.stat().st_size
    return input_path, 0.0, 0.0, size, size


class PipelineWorker(QObject):
    log = Signal(str)
    status = Signal(str, float)
    finished = Signal(str)
    failed = Signal(str)

    def __init__(
        self,
        path: Path,
        use_clean: bool,
        use_image: bool,
        use_precision: bool,
        aggressive: bool,
        do_xml_cleanup: bool,
        force_custom: bool,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self.path = path
        self.use_clean = use_clean
        self.use_image = use_image
        self.use_precision = use_precision
        self.aggressive = aggressive
        self.do_xml_cleanup = do_xml_cleanup
        self.force_custom = force_custom

    def run(self) -> None:
        try:
            self._run_pipeline()
        except Exception as e:  # noqa: BLE001
            self.failed.emit(f"예기치 못한 오류: {e}")

    def _run_pipeline(self) -> None:
        current = self.path
        intermediate_files: list[Path] = []
        log_files: list[Path] = []
        steps: list[str] = []
        if self.use_clean:
            steps.append("clean")
        if self.use_image:
            steps.append("image")
        if self.use_precision:
            steps.append("precision")

        if not steps:
            self.failed.emit("실행할 기능이 없습니다.")
            return

        total = len(steps)
        self.log.emit(f"[INFO] 파이프라인 시작: {current.name}, 단계 {total}개")

        for index, step in enumerate(steps, start=1):
            base = (index - 1) * 100.0 / total
            next_p = index * 100.0 / total
            try:
                if step == "clean":
                    self.status.emit("이름 정의 정리 중...", base)
                    self.log.emit(f"[{index}/{total}] 이름 정의 정리: {current.name}")
                    (
                        backup_path,
                        cleaned_path,
                        stats,
                        ts_dir,
                        top_dir,
                    ) = process_file_gui(str(current))
                    current = Path(cleaned_path)
                    if step != steps[-1]:
                        intermediate_files.append(current)
                    self.log.emit(f" - 백업: {backup_path}")
                    self.log.emit(f" - 정리본: {cleaned_path}")
                    self.log.emit(
                        " - 통계: total="
                        + str(stats["total"])
                        + ", kept="
                        + str(stats["kept"])
                        + ", removed="
                        + str(stats["removed"])
                    )
                elif step == "image":
                    self.status.emit("이미지 최적화 중...", base)
                    self.log.emit(f"[{index}/{total}] 이미지 최적화: {current.name}")
                    (
                        out_path,
                        before,
                        after,
                        count,
                        log_path,
                    ) = run_image_slim(
                        current,
                        max_edge=1400,
                        jpeg_quality=80,
                        progressive=True,
                    )
                    current = out_path
                    if step != steps[-1]:
                        intermediate_files.append(current)
                    saved = before - after
                    pct = (saved / before * 100.0) if before > 0 else 0.0
                    self.log.emit(f" - 이미지 개수: {count}")
                    self.log.emit(
                        " - Before: "
                        + human_size(before)
                        + ", After: "
                        + human_size(after)
                        + ", Saved: "
                        + human_size(saved)
                        + f" ({pct:.1f}%)"
                    )
                    self.log.emit(f" - 로그: {log_path}")
                    log_files.append(log_path)
                elif step == "precision":
                    self.status.emit("정밀 슬리머 실행 중...", base)
                    self.log.emit(f"[{index}/{total}] 정밀 슬리머: {current.name}")
                    has_clean_step = "clean" in steps
                    no_backup = has_clean_step

                    def logger(msg: str) -> None:
                        self.log.emit("[Precision] " + msg)

                    (
                        out_path,
                        saved_mb,
                        pct,
                        old_b,
                        new_b,
                    ) = run_precision_step(
                        current,
                        aggressive=self.aggressive,
                        no_backup=no_backup,
                        do_xml_cleanup=self.do_xml_cleanup,
                        force_custom=self.force_custom,
                        logger=logger,
                    )
                    current = out_path
                    self.log.emit(f" - 결과: {current.name}")
                    self.log.emit(
                        " - Before: "
                        + human_size(old_b)
                        + ", After: "
                        + human_size(new_b)
                        + f", Saved: {saved_mb:.2f} MB ({pct:.1f}%)"
                    )

                self.status.emit("진행 중...", next_p)
            except Exception as e:  # noqa: BLE001
                self.failed.emit(f"{step} 단계에서 오류: {e}")
                return

        # 모든 단계가 성공적으로 끝난 경우에만 중간 산출물 및 로그 정리
        for tmp in intermediate_files:
            try:
                if tmp.exists() and tmp != current:
                    tmp.unlink()
                    self.log.emit(f"[INFO] 중간 결과 삭제: {tmp}")
            except Exception as e:  # noqa: BLE001
                self.log.emit(f"[WARN] 중간 결과 삭제 실패: {tmp} ({e})")

        for log_path in log_files:
            try:
                if log_path.exists():
                    log_path.unlink()
                    self.log.emit(f"[INFO] 로그 파일 삭제: {log_path}")
            except Exception as e:  # noqa: BLE001
                self.log.emit(f"[WARN] 로그 파일 삭제 실패: {log_path} ({e})")

        self.status.emit("모든 작업 완료", 100.0)
        self.log.emit(f"[INFO] 파이프라인 완료. 최종 파일: {current}")
        self.finished.emit(str(current))


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("ExcelSlimmer")
        self.resize(1120, 720)

        self._worker_thread: QThread | None = None
        self._worker: PipelineWorker | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(18, 14, 18, 18)
        root_layout.setSpacing(12)

        header_layout = QHBoxLayout()
        header_layout.setSpacing(8)
        root_layout.addLayout(header_layout)

        title = QLabel("ExcelSlimmer")
        title.setStyleSheet("font-size: 18px; font-weight: 700;")
        header_layout.addWidget(title, 0, Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addStretch(1)

        tabs = QTabWidget()
        root_layout.addWidget(tabs, 1)

        self.pipeline_tab = QWidget()
        self.settings_tab = QWidget()
        tabs.addTab(self.pipeline_tab, "슬리머 실행")
        tabs.addTab(self.settings_tab, "환경 설정")

        pipe_layout = QGridLayout(self.pipeline_tab)
        # 좌우는 12px, 상단은 약간 내려서 슬리머 실행 탭 상단과 라벨 사이 간격을 확보
        pipe_layout.setContentsMargins(12, 8, 12, 0)
        pipe_layout.setHorizontalSpacing(12)

        left_col = QWidget()
        left_layout = QVBoxLayout(left_col)
        left_layout.setContentsMargins(0, 0, 0, 0)
        # 라벨과 카드 사이 간격을 줄이기 위해 spacing을 약간 낮게 설정
        left_layout.setSpacing(6)

        right_col = QWidget()
        right_layout = QVBoxLayout(right_col)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        pipe_layout.addWidget(left_col, 0, 0)
        pipe_layout.addWidget(right_col, 0, 1)
        pipe_layout.setColumnStretch(0, 0)
        pipe_layout.setColumnStretch(1, 1)

        # 대상 파일 카드
        file_label = QLabel("대상 파일")
        file_label.setStyleSheet("font-weight: 600;")
        left_layout.addWidget(file_label)

        file_group = QGroupBox()
        file_group.setStyleSheet(self._card_style())
        fg_layout = QVBoxLayout(file_group)
        fg_layout.setSpacing(6)

        fg_layout.addWidget(QLabel("파일 경로:"))

        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        fg_layout.addWidget(self.file_edit)

        browse_btn = QPushButton("찾기...")
        browse_btn.clicked.connect(self._on_browse)
        fg_layout.addWidget(browse_btn, 0, Qt.AlignRight)

        left_layout.addWidget(file_group)

        # 실행할 기능 카드
        func_label = QLabel("실행할 기능")
        func_label.setStyleSheet("font-weight: 600;")
        left_layout.addWidget(func_label)

        func_group = QGroupBox()
        func_group.setStyleSheet(self._card_style())
        func_layout = QVBoxLayout(func_group)
        func_layout.setSpacing(4)

        self.clean_check = QCheckBox("이름 정의 정리 (definedNames 클린)")
        self.clean_check.setChecked(True)
        self.image_check = QCheckBox("이미지 최적화 (이미지 리사이즈/압축)")
        self.image_check.setChecked(True)
        self.precision_check = QCheckBox("정밀 슬리머 (Precision Plus)")

        func_layout.addWidget(self.clean_check)
        func_layout.addWidget(self.image_check)
        func_layout.addWidget(self.precision_check)

        warn = QLabel("주의: 정밀 슬리머 사용 시 엑셀에서 복구 여부를 물어볼 수 있습니다.")
        warn.setStyleSheet("color: #aa0000; font-size: 9pt;")
        func_layout.addWidget(warn)

        left_layout.addWidget(func_group)

        # 정밀 슬리머 옵션 카드
        opt_label = QLabel("정밀 슬리머 옵션")
        opt_label.setStyleSheet("font-weight: 600;")
        left_layout.addWidget(opt_label)

        opt_group = QGroupBox()
        opt_group.setStyleSheet(self._card_style())
        opt_layout = QVBoxLayout(opt_group)
        opt_layout.setSpacing(4)

        self.aggressive_check = QCheckBox("공격 모드 (이미지 리사이즈 + PNG→JPG)")
        self.xmlcleanup_check = QCheckBox("XML 정리 (calcChain, printerSettings 등)")
        self.force_custom_check = QCheckBox("숨은 XML 데이터 삭제 (customXml, 주의)")

        opt_layout.addWidget(self.aggressive_check)
        opt_layout.addWidget(self.xmlcleanup_check)
        opt_layout.addWidget(self.force_custom_check)

        opt_warn = QLabel("주의: 숨은 XML 데이터 삭제는 일반적인 경우 사용하지 마세요.")
        opt_warn.setStyleSheet("color: #aa0000; font-size: 9pt;")
        opt_layout.addWidget(opt_warn)

        left_layout.addWidget(opt_group)

        # 실행/상태 카드
        run_group = QGroupBox()
        run_group.setStyleSheet(self._card_style())
        run_layout = QVBoxLayout(run_group)
        run_layout.setSpacing(8)

        self.run_button = QPushButton("선택한 기능 실행")
        self.run_button.clicked.connect(self._on_run_clicked)
        run_layout.addWidget(self.run_button, 0, Qt.AlignLeft)

        status_row = QHBoxLayout()
        status_row.setSpacing(8)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        status_row.addWidget(self.progress_bar, 1)

        self.status_label = QLabel("준비됨")
        status_row.addWidget(self.status_label)

        run_layout.addLayout(status_row)

        left_layout.addWidget(run_group)
        left_layout.addStretch(1)

        # 로그 카드
        log_label = QLabel("로그")
        log_label.setStyleSheet("font-weight: 600;")
        right_layout.addWidget(log_label)

        log_group = QGroupBox()
        log_group.setStyleSheet(self._card_style())
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(8, 6, 8, 8)
        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setLineWrapMode(QPlainTextEdit.NoWrap)
        log_layout.addWidget(self.log_edit)

        right_layout.addWidget(log_group)

        # 환경 설정 탭은 간단 안내만
        s_layout = QVBoxLayout(self.settings_tab)
        s_layout.addWidget(QLabel("환경 설정은 추후 확장을 위해 예약된 영역입니다."))
        s_layout.addStretch(1)

        self._update_precision_options_state()
        self.precision_check.toggled.connect(self._update_precision_options_state)

        self._apply_global_widget_style()

    def _card_style(self) -> str:
        return (
            "QGroupBox {"
            "  background: #ffffff;"
            "  border: 1px solid #e0e0e0;"
            "  border-radius: 4px;"
            "  margin-top: 0px;"
            "}"
        )

    def _apply_global_widget_style(self) -> None:
        """Apply a light, uniform border to inputs and buttons.

        This removes the 상대적으로 진한 하단 테두리 느낌 and aligns with the
        카드 테두리 색상.
        """
        self.setStyleSheet(
            "QLineEdit {"
            "  border: 1px solid #d0d0d0;"
            "  border-radius: 3px;"
            "  padding: 3px 6px;"
            "}"
            "QLineEdit:focus {"
            "  border-color: #5b8cff;"
            "}"
            "QPushButton {"
            "  border: 1px solid #d0d0d0;"
            "  border-radius: 3px;"
            "  padding: 4px 10px;"
            "  background: #ffffff;"
            "}"
            "QPushButton:hover {"
            "  background: #f5f5f5;"
            "}"
            "QPushButton:pressed {"
            "  background: #eaeaea;"
            "}"
        )

    def _update_precision_options_state(self) -> None:
        enabled = self.precision_check.isChecked()
        for cb in (self.aggressive_check, self.xmlcleanup_check, self.force_custom_check):
            cb.setEnabled(enabled)

    def _on_browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "대상 Excel 파일 선택",
            "",
            "Excel Files (*.xlsx *.xlsm)",
        )
        if path:
            self.file_edit.setText(path)

    def _append_log(self, text: str) -> None:
        self.log_edit.appendPlainText(text)
        self.log_edit.verticalScrollBar().setValue(self.log_edit.verticalScrollBar().maximum())

    def _set_status(self, text: str, progress: float | None = None) -> None:
        self.status_label.setText(text)
        if progress is not None:
            self.progress_bar.setValue(int(progress))

    def _on_run_clicked(self) -> None:
        path_str = self.file_edit.text().strip()
        if not path_str:
            QMessageBox.warning(self, "안내", "대상 파일을 먼저 선택하세요.")
            return
        path = Path(path_str)
        if not path.exists():
            QMessageBox.critical(self, "오류", f"파일을 찾을 수 없습니다:\n{path}")
            return
        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            QMessageBox.critical(self, "오류", "지원 형식은 .xlsx / .xlsm 입니다.")
            return
        if not (
            self.clean_check.isChecked()
            or self.image_check.isChecked()
            or self.precision_check.isChecked()
        ):
            QMessageBox.information(self, "안내", "실행할 기능을 하나 이상 선택하세요.")
            return

        self.log_edit.clear()
        self.progress_bar.setValue(0)
        self.status_label.setText("작업 시작...")
        self.run_button.setEnabled(False)

        worker = PipelineWorker(
            path=path,
            use_clean=self.clean_check.isChecked(),
            use_image=self.image_check.isChecked(),
            use_precision=self.precision_check.isChecked(),
            aggressive=self.aggressive_check.isChecked(),
            do_xml_cleanup=self.xmlcleanup_check.isChecked(),
            force_custom=self.force_custom_check.isChecked(),
        )
        thread = QThread(self)
        worker.moveToThread(thread)

        worker.log.connect(self._append_log)
        worker.status.connect(self._set_status)

        def on_finished(final_path: str) -> None:
            self._set_status("모든 작업 완료", 100.0)
            self.run_button.setEnabled(True)
            try:
                open_in_explorer_select(Path(final_path))
            except Exception:  # noqa: BLE001
                pass
            QMessageBox.information(
                self,
                "완료",
                f"모든 작업이 완료되었습니다.\n\n최종 결과 파일:\n{final_path}",
            )

        def on_failed(msg: str) -> None:
            self.run_button.setEnabled(True)
            self._set_status("오류 발생", None)
            QMessageBox.critical(self, "오류", msg)

        worker.finished.connect(on_finished)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.failed.connect(on_failed)
        worker.failed.connect(thread.quit)
        worker.failed.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)

        thread.started.connect(worker.run)
        thread.start()

        self._worker_thread = thread
        self._worker = worker


def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
