# -*- coding: utf-8 -*-
from __future__ import annotations

import math
import queue
import re
import sys
import threading
import time
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pythoncom
import win32com.client


WORD_EXTENSIONS = {".doc", ".docx", ".docm"}
OUTPUT_BASENAME = "Word파일_통합"
OUTPUT_EXTENSION = ".xlsx"
ERROR_LOG_NAME = "WordToExcel_error.log"
DEFAULT_MAX_SHEETS = 30
MIN_MAX_SHEETS = 1
MAX_MAX_SHEETS = 9999
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def center_window(root: tk.Tk, width: int, height: int) -> None:
    root.update_idletasks()
    x = max(0, (root.winfo_screenwidth() - width) // 2)
    y = max(0, (root.winfo_screenheight() - height) // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")


def find_word_files(folder: Path) -> list[Path]:
    return sorted(
        [
            p
            for p in folder.iterdir()
            if p.is_file()
            and p.suffix.lower() in WORD_EXTENSIONS
            and not p.name.startswith("~$")
        ],
        key=lambda p: p.name.casefold(),
    )


def choose_settings() -> tuple[Path, int, list[Path]] | None:
    result: dict[str, object] = {}

    root = tk.Tk()
    root.title("Word → Excel 변환 설정")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    center_window(root, 680, 300)

    folder_var = tk.StringVar()
    max_sheets_var = tk.StringVar(value=str(DEFAULT_MAX_SHEETS))
    file_count_var = tk.StringVar(value="Word 파일이 있는 폴더를 선택하세요.")

    frame = ttk.Frame(root, padding=20)
    frame.pack(fill="both", expand=True)

    ttk.Label(frame, text="Word 파일 폴더").grid(row=0, column=0, sticky="w")

    folder_entry = ttk.Entry(frame, textvariable=folder_var, width=67, state="readonly")
    folder_entry.grid(row=1, column=0, sticky="ew", pady=(5, 0))

    def browse_folder() -> None:
        selected = filedialog.askdirectory(
            parent=root,
            title="Word 파일이 있는 폴더 선택",
            mustexist=True,
        )
        if not selected:
            return

        folder_var.set(selected)
        try:
            files = find_word_files(Path(selected))
            file_count_var.set(f"감지된 Word 파일: {len(files)}개")
        except Exception:
            file_count_var.set("폴더 내용을 확인할 수 없습니다.")

    ttk.Button(frame, text="폴더 선택", command=browse_folder).grid(
        row=1, column=1, padx=(8, 0), pady=(5, 0)
    )
    ttk.Label(frame, textvariable=file_count_var).grid(
        row=2, column=0, columnspan=2, sticky="w", pady=(6, 16)
    )

    ttk.Separator(frame).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 15))

    ttk.Label(frame, text="Excel 파일당 최대 Sheet 수").grid(row=4, column=0, sticky="w")

    spin = ttk.Spinbox(
        frame,
        from_=MIN_MAX_SHEETS,
        to=MAX_MAX_SHEETS,
        textvariable=max_sheets_var,
        width=12,
    )
    spin.grid(row=5, column=0, sticky="w", pady=(5, 0))

    ttk.Label(
        frame,
        text=(
            f"{MIN_MAX_SHEETS}~{MAX_MAX_SHEETS} 사이에서 지정할 수 있습니다. "
            "입력 파일 수보다 크게 지정하면 Excel 파일이 분할되지 않습니다."
        ),
    ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(6, 18))

    button_frame = ttk.Frame(frame)
    button_frame.grid(row=7, column=0, columnspan=2, sticky="e")

    def start_conversion() -> None:
        folder_text = folder_var.get().strip()
        if not folder_text:
            messagebox.showwarning("폴더 선택", "먼저 Word 파일이 있는 폴더를 선택해주세요.", parent=root)
            return

        folder = Path(folder_text)
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror("폴더 오류", "선택한 폴더를 찾을 수 없습니다.", parent=root)
            return

        try:
            max_sheets = int(max_sheets_var.get().strip())
        except ValueError:
            messagebox.showwarning(
                "Sheet 수 확인",
                "Excel 파일당 최대 Sheet 수를 숫자로 입력해주세요.",
                parent=root,
            )
            return

        if not (MIN_MAX_SHEETS <= max_sheets <= MAX_MAX_SHEETS):
            messagebox.showwarning(
                "Sheet 수 확인",
                f"Sheet 수는 {MIN_MAX_SHEETS}~{MAX_MAX_SHEETS} 사이로 지정해주세요.",
                parent=root,
            )
            return

        try:
            word_files = find_word_files(folder)
        except Exception as exc:
            messagebox.showerror("폴더 오류", f"폴더를 읽을 수 없습니다.\n\n{exc}", parent=root)
            return

        if not word_files:
            messagebox.showwarning(
                "Word 파일 없음",
                "선택한 폴더에 .doc / .docx / .docm 파일이 없습니다.",
                parent=root,
            )
            return

        result["folder"] = folder
        result["max_sheets"] = max_sheets
        result["word_files"] = word_files
        root.destroy()

    ttk.Button(button_frame, text="취소", command=root.destroy).pack(side="left", padx=(0, 8))
    ttk.Button(button_frame, text="변환 시작", command=start_conversion).pack(side="left")

    frame.columnconfigure(0, weight=1)
    root.mainloop()

    if "folder" not in result:
        return None

    return (
        result["folder"],  # type: ignore[return-value]
        result["max_sheets"],  # type: ignore[return-value]
        result["word_files"],  # type: ignore[return-value]
    )


def normalize_sheet_base(filename_stem: str) -> str:
    name = INVALID_SHEET_CHARS.sub("_", filename_stem)
    name = "".join(ch for ch in name if ord(ch) >= 32)
    name = name.strip().strip("'").strip()
    return name or "Sheet"


def make_unique_sheet_name(filename_stem: str, used: set[str]) -> str:
    base = normalize_sheet_base(filename_stem)
    candidate = base[:31]
    if candidate.casefold() not in used:
        return candidate

    index = 2
    while True:
        suffix = f"_{index}"
        candidate = f"{base[:31-len(suffix)]}{suffix}"
        if candidate.casefold() not in used:
            return candidate
        index += 1


def choose_run_number(folder: Path, multi: bool, expected_parts: int) -> int:
    run_number = 1
    while True:
        run_suffix = "" if run_number == 1 else f"_{run_number}"
        if multi:
            stems = [
                f"{OUTPUT_BASENAME}{run_suffix}_part{part:02d}"
                for part in range(1, expected_parts + 1)
            ]
        else:
            stems = [f"{OUTPUT_BASENAME}{run_suffix}"]

        if all(not (folder / f"{stem}{OUTPUT_EXTENSION}").exists() for stem in stems):
            return run_number
        run_number += 1


def output_path_for_part(
    folder: Path,
    multi: bool,
    run_number: int,
    part_index: int,
) -> Path:
    run_suffix = "" if run_number == 1 else f"_{run_number}"
    if multi:
        stem = f"{OUTPUT_BASENAME}{run_suffix}_part{part_index:02d}"
    else:
        stem = f"{OUTPUT_BASENAME}{run_suffix}"
    return folder / f"{stem}{OUTPUT_EXTENSION}"


def write_error_log(folder: Path, lines: list[str]) -> None:
    if not lines:
        return
    (folder / ERROR_LOG_NAME).write_text(
        "\n".join(lines).rstrip() + "\n",
        encoding="utf-8-sig",
    )


def paste_word_document_to_sheet(word_doc, excel_app, worksheet) -> None:
    """
    Word COM Copy -> Excel COM Paste.

    이 경로는 Windows 시스템 Clipboard를 사용하므로 병렬화하지 않는다.
    사용자가 처리 중 Clipboard를 변경하는 작업을 하면 붙여넣기 결과가 달라질 수 있다.
    """
    word_doc.Content.Copy()
    pythoncom.PumpWaitingMessages()
    time.sleep(0.12)

    worksheet.Activate()
    target = worksheet.Range("A1")
    target.Select()

    try:
        worksheet.Paste(Destination=target)
    except Exception:
        excel_app.ActiveSheet.Paste()

    pythoncom.PumpWaitingMessages()
    time.sleep(0.04)

    try:
        excel_app.CutCopyMode = False
    except Exception:
        pass


def prepare_workbook(excel_app):
    workbook = excel_app.Workbooks.Add()
    while workbook.Worksheets.Count > 1:
        workbook.Worksheets(workbook.Worksheets.Count).Delete()
    return workbook


def save_and_close_workbook(workbook, output_path: Path) -> None:
    workbook.SaveAs(str(output_path.resolve()), FileFormat=51)
    workbook.Close(SaveChanges=False)


def conversion_worker(
    folder: Path,
    word_files: list[Path],
    max_sheets_per_book: int,
    events: queue.Queue,
) -> None:
    total = len(word_files)
    multi = total > max_sheets_per_book
    expected_parts = max(1, math.ceil(total / max_sheets_per_book))
    run_number = choose_run_number(folder, multi, expected_parts)

    errors: list[str] = []
    outputs: list[Path] = []
    success_count = 0
    fail_count = 0
    current_part = 1
    sheets_in_current_book = 0
    used_sheet_names: set[str] = set()

    word = None
    excel = None
    workbook = None

    pythoncom.CoInitialize()
    try:
        events.put(("status", "Microsoft Word / Excel 연결 중...", 0, total, current_part))

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        for index, word_path in enumerate(word_files, start=1):
            processing_part = current_part
            events.put(("status", word_path.name, index - 1, total, processing_part))

            doc = None
            worksheet = None
            created_new_sheet = False

            try:
                doc = word.Documents.Open(
                    FileName=str(word_path.resolve()),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                    ConfirmConversions=False,
                )

                if workbook is None:
                    workbook = prepare_workbook(excel)
                    sheets_in_current_book = 0
                    used_sheet_names.clear()

                if sheets_in_current_book == 0:
                    worksheet = workbook.Worksheets(1)
                else:
                    worksheet = workbook.Worksheets.Add(
                        After=workbook.Worksheets(workbook.Worksheets.Count)
                    )
                    created_new_sheet = True

                sheet_name = make_unique_sheet_name(word_path.stem, used_sheet_names)
                worksheet.Name = sheet_name

                paste_word_document_to_sheet(doc, excel, worksheet)

                used_sheet_names.add(sheet_name.casefold())
                sheets_in_current_book += 1
                success_count += 1

                if sheets_in_current_book >= max_sheets_per_book:
                    output_path = output_path_for_part(
                        folder,
                        multi,
                        run_number,
                        current_part,
                    )
                    save_and_close_workbook(workbook, output_path)
                    outputs.append(output_path)
                    workbook = None
                    sheets_in_current_book = 0
                    used_sheet_names.clear()
                    current_part += 1

            except Exception as exc:
                fail_count += 1
                errors.append(
                    "\n".join(
                        [
                            f"[실패] {word_path.name}",
                            f"{type(exc).__name__}: {exc}",
                            traceback.format_exc(),
                        ]
                    )
                )

                try:
                    if worksheet is not None and workbook is not None:
                        if created_new_sheet and workbook.Worksheets.Count > 1:
                            worksheet.Delete()
                        elif sheets_in_current_book == 0:
                            worksheet.Cells.Clear()
                            worksheet.Name = "Sheet1"
                except Exception:
                    pass

            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass

            events.put(("progress", word_path.name, index, total, processing_part))

        if workbook is not None:
            if sheets_in_current_book > 0:
                output_path = output_path_for_part(
                    folder,
                    multi,
                    run_number,
                    current_part,
                )
                save_and_close_workbook(workbook, output_path)
                outputs.append(output_path)
                workbook = None
            else:
                workbook.Close(SaveChanges=False)
                workbook = None

        if success_count == 0:
            errors.insert(0, "모든 Word 파일 처리에 실패했습니다.")

        if errors:
            errors.insert(
                0,
                f"총 {total}개 중 {success_count}개 성공, {fail_count}개 실패했습니다.",
            )
            write_error_log(folder, errors)
        else:
            old_log = folder / ERROR_LOG_NAME
            if old_log.exists():
                try:
                    old_log.unlink()
                except Exception:
                    pass

        events.put(
            (
                "done",
                success_count,
                fail_count,
                [p.name for p in outputs],
                total,
            )
        )

    except Exception as exc:
        fatal = "\n".join(
            [
                "[프로그램 오류]",
                f"{type(exc).__name__}: {exc}",
                traceback.format_exc(),
            ]
        )
        errors.insert(0, fatal)
        try:
            write_error_log(folder, errors)
        except Exception:
            pass
        events.put(("fatal", str(exc), success_count, fail_count, total))

    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.ScreenUpdating = True
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass

        if word is not None:
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                pass

        pythoncom.CoUninitialize()


def show_progress(
    folder: Path,
    word_files: list[Path],
    max_sheets_per_book: int,
) -> None:
    total = len(word_files)
    events: queue.Queue = queue.Queue()

    root = tk.Tk()
    root.title("Word → Excel 변환 진행")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    center_window(root, 680, 345)

    frame = ttk.Frame(root, padding=18)
    frame.pack(fill="both", expand=True)

    title_label = ttk.Label(frame, text="Word 파일을 Excel로 변환 중입니다.")
    title_label.pack(anchor="w")

    current_label = ttk.Label(frame, text="준비 중...", width=82)
    current_label.pack(anchor="w", pady=(10, 6))

    progress = ttk.Progressbar(
        frame,
        orient="horizontal",
        mode="determinate",
        maximum=max(1, total),
        length=640,
    )
    progress.pack(fill="x")

    info_frame = ttk.Frame(frame)
    info_frame.pack(fill="x", pady=(7, 0))

    count_label = ttk.Label(info_frame, text=f"0 / {total}  (0%)")
    count_label.pack(side="left")

    part_label = ttk.Label(info_frame, text="Excel part 01")
    part_label.pack(side="right")

    setting_label = ttk.Label(
        frame,
        text=f"현재 설정: Excel 파일당 최대 {max_sheets_per_book}개 Sheet",
    )
    setting_label.pack(anchor="w", pady=(10, 10))

    alert_frame = tk.Frame(
        frame,
        background="#FFF4CE",
        highlightbackground="#D6B656",
        highlightthickness=1,
        padx=12,
        pady=10,
    )
    alert_frame.pack(fill="x")

    tk.Label(
        alert_frame,
        text="⚠ 복붙(Clipboard) 사용 주의",
        background="#FFF4CE",
        foreground="#5C4300",
        font=("Segoe UI", 10, "bold"),
    ).pack(anchor="w")

    tk.Label(
        alert_frame,
        text=(
            "이 프로그램은 Word에서 Copy한 뒤 Excel에 Paste하는 방식입니다. 처리 중에는 "
            "Ctrl+C / Ctrl+X / 파일 복사 / 캡처·클립보드 도구 등 Clipboard를 사용하는 작업을 "
            "가급적 하지 마세요. Clipboard 내용이 바뀌면 Sheet에 잘못 붙여넣어질 수 있습니다."
        ),
        background="#FFF4CE",
        foreground="#5C4300",
        justify="left",
        wraplength=620,
    ).pack(anchor="w", pady=(5, 0))

    button_frame = ttk.Frame(frame)
    button_frame.pack(fill="x", pady=(12, 0))

    close_button = ttk.Button(button_frame, text="닫기", command=root.destroy, state="disabled")
    close_button.pack(side="right")

    root.protocol("WM_DELETE_WINDOW", lambda: None)

    def finish_ui() -> None:
        root.protocol("WM_DELETE_WINDOW", root.destroy)
        close_button.config(state="normal")
        root.attributes("-topmost", False)

    def poll_events() -> None:
        try:
            while True:
                event = events.get_nowait()
                kind = event[0]

                if kind in {"status", "progress"}:
                    _, filename, done_count, event_total, part_index = event
                    current_label.config(text=filename)
                    progress["maximum"] = max(1, event_total)
                    progress["value"] = done_count
                    percent = int((done_count / event_total) * 100) if event_total else 0
                    count_label.config(text=f"{done_count} / {event_total}  ({percent}%)")
                    part_label.config(text=f"Excel part {max(1, part_index):02d}")

                elif kind == "done":
                    _, success_count, fail_count, outputs, event_total = event
                    progress["value"] = event_total
                    count_label.config(text=f"{event_total} / {event_total}  (100%)")
                    title_label.config(
                        text=(
                            f"완료 - 성공 {success_count}개 / 실패 {fail_count}개 / "
                            f"Excel {len(outputs)}개"
                        )
                    )
                    if outputs:
                        current_label.config(text="생성: " + ", ".join(outputs))
                    else:
                        current_label.config(text="생성된 Excel 파일이 없습니다.")
                    finish_ui()
                    return

                elif kind == "fatal":
                    _, message, success_count, fail_count, event_total = event
                    title_label.config(text="변환 중 오류가 발생했습니다.")
                    current_label.config(text=message[:82])
                    count_label.config(
                        text=f"성공 {success_count} / 실패 {fail_count} / 전체 {event_total}"
                    )
                    finish_ui()
                    return

        except queue.Empty:
            pass

        root.after(80, poll_events)

    worker = threading.Thread(
        target=conversion_worker,
        args=(folder, word_files, max_sheets_per_book, events),
        daemon=True,
    )
    worker.start()
    root.after(50, poll_events)
    root.mainloop()


def main() -> int:
    settings = choose_settings()
    if settings is None:
        return 0

    folder, max_sheets_per_book, word_files = settings
    show_progress(folder, word_files, max_sheets_per_book)
    return 0


if __name__ == "__main__":
    sys.exit(main())
