import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


def _ensure_module_paths() -> None:
    base = Path(__file__).resolve().parent
    root = base.parent
    for name in ("ExcelCleaner", "ExcelImageOptimization", "ExcelByteReduce"):
        p = root / name
        if p.is_dir():
            sys.path.insert(0, str(p))


_ensure_module_paths()

from gui_clean_defined_names_desktop_date import process_file_gui
from excel_image_slimmer_gui_v3 import (
    slim_xlsx,
    human_size,
    open_in_explorer_select,
)
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


class ExcelSuiteApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("ExcelSlimmer")
        self.root.geometry("1120x720")
        self.root.minsize(960, 640)

        self.file_var = tk.StringVar()
        self.clean_var = tk.IntVar(value=1)
        self.image_var = tk.IntVar(value=1)
        self.precision_var = tk.IntVar(value=0)

        self.prec_aggressive_var = tk.IntVar(value=0)
        self.prec_xmlcleanup_var = tk.IntVar(value=0)
        self.prec_force_custom_var = tk.IntVar(value=0)

        self.status_var = tk.StringVar(value="준비됨")
        self.progress_var = tk.DoubleVar(value=0.0)

        self._build_ui()

    def _build_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass

        base_bg = "#f4f5fb"
        card_bg = "#ffffff"
        self.root.configure(bg=base_bg)

        style.configure("App.TFrame", background=base_bg)
        style.configure("Card.TFrame", background=card_bg)
        style.configure("Card.TLabelframe", background=card_bg)
        style.configure("Card.TLabelframe.Label", background=card_bg, font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), background=base_bg)
        style.configure("SubHeader.TLabel", foreground="#666666", background=base_bg)

        outer = ttk.Frame(self.root, style="App.TFrame", padding=(18, 14, 18, 18))
        outer.pack(fill="both", expand=True)

        header_frame = ttk.Frame(outer, style="App.TFrame")
        header_frame.pack(fill="x", pady=(0, 12))

        title_label = ttk.Label(header_frame, text="ExcelSlimmer", style="Header.TLabel")
        title_label.pack(side="left", anchor="w")

        subtitle = ttk.Label(
            header_frame,
            text="Excel 파일을 한 번에 정리하고 슬림하게 만드는 통합 도구",
            style="SubHeader.TLabel",
        )
        subtitle.pack(side="left", padx=(12, 0), anchor="s")

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)

        pipeline_page = ttk.Frame(notebook, style="App.TFrame")
        settings_page = ttk.Frame(notebook, style="App.TFrame")
        notebook.add(pipeline_page, text="파이프라인")
        notebook.add(settings_page, text="환경 설정")

        ttk.Label(
            settings_page,
            text="환경 설정은 추후 확장을 위해 예약된 영역입니다.",
            style="SubHeader.TLabel",
        ).pack(pady=20, padx=20, anchor="w")

        pipeline_page.columnconfigure(0, weight=0, minsize=380)
        pipeline_page.columnconfigure(1, weight=1)
        pipeline_page.rowconfigure(0, weight=1)

        left_col = ttk.Frame(pipeline_page, style="App.TFrame")
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        right_col = ttk.Frame(pipeline_page, style="App.TFrame")
        right_col.grid(row=0, column=1, sticky="nsew")

        file_card = ttk.Labelframe(
            left_col,
            text="대상 파일",
            style="Card.TLabelframe",
            padding=(12, 10, 12, 12),
        )
        file_card.pack(fill="x", pady=(0, 10))
        ttk.Label(file_card, text="파일 경로:").pack(anchor="w")
        entry = ttk.Entry(file_card, textvariable=self.file_var)
        entry.pack(fill="x", expand=True, pady=(4, 6))
        ttk.Button(file_card, text="찾기...", command=self._select_file).pack(anchor="e")

        pipeline_card = ttk.Labelframe(
            left_col,
            text="실행할 기능",
            style="Card.TLabelframe",
            padding=(12, 8, 12, 10),
        )
        pipeline_card.pack(fill="x", pady=(0, 10))

        ttk.Checkbutton(
            pipeline_card,
            text="1) 이름 정의 정리 (definedNames 클린)",
            variable=self.clean_var,
        ).pack(anchor="w", pady=(2, 2))
        ttk.Checkbutton(
            pipeline_card,
            text="2) 이미지 최적화 (이미지 리사이즈/압축)",
            variable=self.image_var,
        ).pack(anchor="w", pady=(2, 2))
        self.precision_check = ttk.Checkbutton(
            pipeline_card,
            text="3) 정밀 슬리머 (Precision Plus)",
            variable=self.precision_var,
            command=self._on_precision_toggle,
        )
        self.precision_check.pack(anchor="w", pady=(2, 0))
        self.precision_warning = ttk.Label(
            pipeline_card,
            text="주의: 정밀 슬리머 사용 시 엑셀에서 복구 여부를 물어볼 수 있습니다.",
            foreground="#aa0000",
            background=card_bg,
        )
        self.precision_warning.pack(anchor="w", padx=18, pady=(0, 4))

        precision_card = ttk.Labelframe(
            left_col,
            text="정밀 슬리머 옵션",
            style="Card.TLabelframe",
            padding=(12, 8, 12, 10),
        )
        precision_card.pack(fill="x", pady=(0, 10))

        self.prec_aggressive_cb = ttk.Checkbutton(
            precision_card,
            text="공격 모드 (이미지 리사이즈 + PNG→JPG)",
            variable=self.prec_aggressive_var,
        )
        self.prec_aggressive_cb.pack(anchor="w", pady=(2, 2))
        self.prec_xmlcleanup_cb = ttk.Checkbutton(
            precision_card,
            text="XML 정리 (calcChain, printerSettings 등)",
            variable=self.prec_xmlcleanup_var,
        )
        self.prec_xmlcleanup_cb.pack(anchor="w", pady=(2, 2))
        self.prec_force_custom_cb = ttk.Checkbutton(
            precision_card,
            text="숨은 XML 데이터 삭제 (customXml, 주의)",
            variable=self.prec_force_custom_var,
        )
        self.prec_force_custom_cb.pack(anchor="w", pady=(2, 2))
        self.prec_force_custom_hint = ttk.Label(
            precision_card,
            text="권장: 일반적인 경우 사용하지 마세요",
            foreground="#aa0000",
            background=card_bg,
        )
        self.prec_force_custom_hint.pack(anchor="w", padx=18, pady=(0, 4))

        run_card = ttk.Frame(left_col, style="Card.TFrame", padding=(12, 10, 12, 12))
        run_card.pack(fill="x")

        self.run_button = ttk.Button(
            run_card,
            text="선택한 기능 실행",
            command=self._on_run_clicked,
        )
        self.run_button.pack(anchor="w")

        status_row = ttk.Frame(run_card, style="Card.TFrame")
        status_row.pack(fill="x", pady=(8, 0))
        status_label = ttk.Label(status_row, textvariable=self.status_var)
        status_label.pack(side="right")
        self.progress = ttk.Progressbar(
            status_row,
            maximum=100.0,
            variable=self.progress_var,
        )
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 8))

        log_card = ttk.Labelframe(
            right_col,
            text="로그",
            style="Card.TLabelframe",
            padding=(12, 8, 12, 12),
        )
        log_card.pack(fill="both", expand=True)

        self.log_box = scrolledtext.ScrolledText(
            log_card,
            height=10,
            state="disabled",
            bg=card_bg,
            relief="flat",
        )
        self.log_box.pack(fill="both", expand=True)

        self._update_precision_options_state()

    def _on_precision_toggle(self) -> None:
        self._update_precision_options_state()

    def _update_precision_options_state(self) -> None:
        enabled = bool(self.precision_var.get())
        state = "normal" if enabled else "disabled"
        for cb in (
            self.prec_aggressive_cb,
            self.prec_xmlcleanup_cb,
            self.prec_force_custom_cb,
        ):
            cb.configure(state=state)

    def _select_file(self) -> None:
        filetypes = [("Excel 파일", "*.xlsx;*.xlsm"), ("모든 파일", "*.*")]
        path = filedialog.askopenfilename(
            title="대상 Excel 파일 선택",
            filetypes=filetypes,
        )
        if path:
            self.file_var.set(path)

    def _append_log(self, text: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def log(self, text: str) -> None:
        self.root.after(0, lambda: self._append_log(text))

    def set_status(self, text: str, progress: float = None) -> None:
        def _update() -> None:
            self.status_var.set(text)
            if progress is not None:
                self.progress_var.set(progress)

        self.root.after(0, _update)

    def show_info(self, title: str, text: str) -> None:
        self.root.after(0, lambda: messagebox.showinfo(title, text))

    def show_error(self, title: str, text: str) -> None:
        self.root.after(0, lambda: messagebox.showerror(title, text))

    def _on_run_clicked(self) -> None:
        path_str = self.file_var.get().strip()
        if not path_str:
            messagebox.showwarning("안내", "대상 파일을 먼저 선택하세요.")
            return
        path = Path(path_str)
        if not path.exists():
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{path}")
            return
        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            messagebox.showerror("오류", "지원 형식은 .xlsx / .xlsm 입니다.")
            return
        if not (
            self.clean_var.get()
            or self.image_var.get()
            or self.precision_var.get()
        ):
            messagebox.showinfo("안내", "실행할 기능을 하나 이상 선택하세요.")
            return

        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.progress_var.set(0.0)
        self.status_var.set("작업 시작...")
        self.run_button.configure(state="disabled")

        t = threading.Thread(
            target=self._run_pipeline_worker,
            args=(path,),
            daemon=True,
        )
        t.start()

    def _run_pipeline_worker(self, start_path: Path) -> None:
        try:
            self._run_pipeline(start_path)
        except Exception as e:
            self.log(f"[ERROR] 예기치 못한 오류: {e}")
            traceback.print_exc()
            self.set_status("오류 발생", None)
            self.show_error("오류", f"예기치 못한 오류가 발생했습니다.\n\n{e}")
        finally:
            self.root.after(0, lambda: self.run_button.configure(state="normal"))

    def _reset_ui_after_finish(self) -> None:
        """파이프라인 완료 후 기본 상태로 되돌립니다 (로그는 유지)."""
        self.file_var.set("")
        self.clean_var.set(1)
        self.image_var.set(1)
        self.precision_var.set(0)
        self.prec_aggressive_var.set(0)
        self.prec_xmlcleanup_var.set(0)
        self.prec_force_custom_var.set(0)
        self._update_precision_options_state()
        self.progress_var.set(0.0)
        self.status_var.set("준비됨")

    def _run_pipeline(self, start_path: Path) -> None:
        current = start_path
        intermediate_files = []
        log_files = []
        steps = []
        if self.clean_var.get():
            steps.append("clean")
        if self.image_var.get():
            steps.append("image")
        if self.precision_var.get():
            steps.append("precision")

        total = len(steps)
        self.log(f"[INFO] 파이프라인 시작: {start_path.name}, 단계 {total}개")

        for index, step in enumerate(steps, start=1):
            base = (index - 1) * 100.0 / total
            next_p = index * 100.0 / total
            try:
                if step == "clean":
                    self.set_status("이름 정의 정리 중...", base)
                    self.log(f"[{index}/{total}] 이름 정의 정리: {current.name}")
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
                    self.log(f" - 백업: {backup_path}")
                    self.log(f" - 정리본: {cleaned_path}")
                    self.log(
                        " - 통계: total="
                        + str(stats["total"])
                        + ", kept="
                        + str(stats["kept"])
                        + ", removed="
                        + str(stats["removed"])
                    )
                elif step == "image":
                    self.set_status("이미지 최적화 중...", base)
                    self.log(f"[{index}/{total}] 이미지 최적화: {current.name}")
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
                    self.log(f" - 이미지 개수: {count}")
                    self.log(
                        " - Before: "
                        + human_size(before)
                        + ", After: "
                        + human_size(after)
                        + ", Saved: "
                        + human_size(saved)
                        + f" ({pct:.1f}%)"
                    )
                    self.log(f" - 로그: {log_path}")
                    log_files.append(log_path)
                elif step == "precision":
                    self.set_status("정밀 슬리머 실행 중...", base)
                    self.log(f"[{index}/{total}] 정밀 슬리머: {current.name}")
                    aggressive = bool(self.prec_aggressive_var.get())
                    has_clean_step = "clean" in steps
                    no_backup = has_clean_step
                    do_xml_cleanup = bool(self.prec_xmlcleanup_var.get())
                    force_custom = bool(self.prec_force_custom_var.get())

                    def logger(msg: str) -> None:
                        self.log("[Precision] " + msg)

                    (
                        out_path,
                        saved_mb,
                        pct,
                        old_b,
                        new_b,
                    ) = run_precision_step(
                        current,
                        aggressive,
                        no_backup,
                        do_xml_cleanup,
                        force_custom,
                        logger,
                    )
                    current = out_path
                    self.log(f" - 결과: {current.name}")
                    self.log(
                        " - Before: "
                        + human_size(old_b)
                        + ", After: "
                        + human_size(new_b)
                        + f", Saved: {saved_mb:.2f} MB ({pct:.1f}%)"
                    )

                self.set_status("진행 중...", next_p)
            except Exception as e:
                self.log(f"[ERROR] {step} 단계에서 오류: {e}")
                self.set_status("오류 발생", None)
                self.show_error(
                    "오류",
                    f"{step} 단계에서 오류가 발생했습니다.\n\n{e}",
                )
                return

        # 모든 단계가 성공적으로 끝난 경우에만 중간 산출물 및 로그 정리
        for tmp in intermediate_files:
            try:
                if tmp.exists() and tmp != current:
                    tmp.unlink()
                    self.log(f"[INFO] 중간 결과 삭제: {tmp}")
            except Exception as e:
                self.log(f"[WARN] 중간 결과 삭제 실패: {tmp} ({e})")

        for log_path in log_files:
            try:
                if log_path.exists():
                    log_path.unlink()
                    self.log(f"[INFO] 로그 파일 삭제: {log_path}")
            except Exception as e:
                self.log(f"[WARN] 로그 파일 삭제 실패: {log_path} ({e})")

        self.set_status("모든 작업 완료", 100.0)
        self.log(f"[INFO] 파이프라인 완료. 최종 파일: {current}")

        def _after_msg() -> None:
            try:
                open_in_explorer_select(current)
            except Exception:
                pass
            # 로그는 유지하고 나머지 UI 상태만 초기화
            self._reset_ui_after_finish()

        self.root.after(
            0,
            lambda: (
                messagebox.showinfo(
                    "완료",
                    f"모든 작업이 완료되었습니다.\n\n최종 결과 파일:\n{current}",
                ),
                _after_msg(),
            ),
        )

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = ExcelSuiteApp()
    app.run()


if __name__ == "__main__":
    main()
