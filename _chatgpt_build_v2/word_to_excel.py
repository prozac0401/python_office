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
from tkinter import filedialog, ttk

import pythoncom
import win32com.client

WORD_EXTENSIONS = {".doc", ".docx", ".docm"}
OUTPUT_BASENAME = "Word파일_통합"
OUTPUT_EXTENSION = ".xlsx"
ERROR_LOG_NAME = "WordToExcel_error.log"
MAX_SHEETS_PER_BOOK = 30
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def choose_folder() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
        selected = filedialog.askdirectory(
            parent=root,
            title="Word 파일이 있는 폴더 선택",
            mustexist=True,
        )
    finally:
        root.destroy()
    return Path(selected) if selected else None


def find_word_files(folder: Path) -> list[Path]:
    return sorted(
        [
            p for p in folder.iterdir()
            if p.is_file()
            and p.suffix.lower() in WORD_EXTENSIONS
            and not p.name.startswith("~$")
        ],
        key=lambda p: p.name.casefold(),
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
    """같은 실행에서 생성되는 모든 part가 동일한 충돌 회피 번호를 사용하도록 한다."""
    run_number = 1
    while True:
        if multi:
            stems = []
            run_suffix = "" if run_number == 1 else f"_{run_number}"
            for part in range(1, expected_parts + 1):
                stems.append(f"{OUTPUT_BASENAME}{run_suffix}_part{part:02d}")
        else:
            run_suffix = "" if run_number == 1 else f"_{run_number}"
            stems = [f"{OUTPUT_BASENAME}{run_suffix}"]

        if all(not (folder / f"{stem}{OUTPUT_EXTENSION}").exists() for stem in stems):
            return run_number
        run_number += 1


def output_path_for_part(folder: Path, multi: bool, run_number: int, part_index: int) -> Path:
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
    시스템 Clipboard를 사용하는 Word COM Copy -> Excel COM Paste.
    Clipboard가 프로세스 전역 자원이므로 이 구간은 의도적으로 직렬 실행한다.
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


def conversion_worker(folder: Path, word_files: list[Path], events: queue.Queue) -> None:
    total = len(word_files)
    multi = total > MAX_SHEETS_PER_BOOK
    expected_parts = max(1, math.ceil(total / MAX_SHEETS_PER_BOOK))
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
            events.put(("status", word_path.name, index - 1, total, current_part))

            doc = None
            worksheet = None
            created_new_sheet = False

            try:
                # 문서 열기는 먼저 수행한다. 실패 문서는 빈 Excel 시트를 만들지 않는다.
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

                if sheets_in_current_book >= MAX_SHEETS_PER_BOOK:
                    output_path = output_path_for_part(
                        folder, multi, run_number, current_part
                    )
                    save_and_close_workbook(workbook, output_path)
                    outputs.append(output_path)
                    workbook = None
                    sheets_in_current_book = 0
                    used_sheet_names.clear()
                    current_part += 1

            except Exception as exc:
                fail_count += 1
                errors.append("\n".join([
                    f"[실패] {word_path.name}",
                    f"{type(exc).__name__}: {exc}",
                    traceback.format_exc(),
                ]))

                # 실패하면서 추가한 빈 시트 정리
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

            events.put(("progress", word_path.name, index, total, current_part))

        # 마지막 1~29개 시트 저장
        if workbook is not None:
            if sheets_in_current_book > 0:
                output_path = output_path_for_part(
                    folder, multi, run_number, current_part
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

        events.put((
            "done",
            success_count,
            fail_count,
            [p.name for p in outputs],
            total,
        ))

    except Exception as exc:
        fatal = "\n".join([
            "[프로그램 오류]",
            f"{type(exc).__name__}: {exc}",
            traceback.format_exc(),
        ])
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


def show_progress(folder: Path, word_files: list[Path]) -> None:
    total = len(word_files)
    events: queue.Queue = queue.Queue()

    root = tk.Tk()
    root.title("Word → Excel 변환")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    width, height = 560, 185
    root.update_idletasks()
    x = max(0, (root.winfo_screenwidth() - width) // 2)
    y = max(0, (root.winfo_screenheight() - height) // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

    frame = ttk.Frame(root, padding=18)
    frame.pack(fill="both", expand=True)

    title_label = ttk.Label(frame, text="Word 파일을 Excel로 변환 중입니다.")
    title_label.pack(anchor="w")

    current_label = ttk.Label(frame, text="준비 중...", width=72)
    current_label.pack(anchor="w", pady=(10, 6))

    progress = ttk.Progressbar(
        frame,
        orient="horizontal",
        mode="determinate",
        maximum=max(1, total),
        length=520,
    )
    progress.pack(fill="x")

    info_frame = ttk.Frame(frame)
    info_frame.pack(fill="x", pady=(7, 0))

    count_label = ttk.Label(info_frame, text=f"0 / {total}  (0%)")
    count_label.pack(side="left")

    part_label = ttk.Label(info_frame, text="Excel part 01")
    part_label.pack(side="right")

    detail_label = ttk.Label(
        frame,
        text="※ Clipboard 충돌 방지를 위해 Copy → Paste 구간은 순차 처리합니다.",
    )
    detail_label.pack(anchor="w", pady=(12, 0))

    # 처리 중 실수로 닫아 COM 프로세스만 남는 상황 방지
    root.protocol("WM_DELETE_WINDOW", lambda: None)

    def poll_events():
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
                    if outputs:
                        current_label.config(text="완료: " + ", ".join(outputs))
                    else:
                        current_label.config(text="완료되었지만 생성된 Excel 파일이 없습니다.")
                    title_label.config(
                        text=f"완료 - 성공 {success_count}개 / 실패 {fail_count}개 / Excel {len(outputs)}개"
                    )
                    detail_label.config(
                        text="오류가 있었다면 선택한 폴더의 WordToExcel_error.log를 확인하세요."
                    )
                    root.protocol("WM_DELETE_WINDOW", root.destroy)
                    root.after(3000, root.destroy)
                    return

                elif kind == "fatal":
                    _, message, success_count, fail_count, event_total = event
                    title_label.config(text="변환 중 오류가 발생했습니다.")
                    current_label.config(text=message[:72])
                    detail_label.config(text="선택한 폴더의 WordToExcel_error.log를 확인하세요.")
                    root.protocol("WM_DELETE_WINDOW", root.destroy)
                    root.after(4500, root.destroy)
                    return

        except queue.Empty:
            pass

        root.after(80, poll_events)

    worker = threading.Thread(
        target=conversion_worker,
        args=(folder, word_files, events),
        daemon=True,
    )
    worker.start()
    root.after(50, poll_events)
    root.mainloop()


def main() -> int:
    folder = choose_folder()
    if folder is None:
        return 0

    word_files = find_word_files(folder)
    if not word_files:
        write_error_log(folder, [
            "처리할 Word 파일을 찾지 못했습니다.",
            "대상 확장자: .doc, .docx, .docm",
            f"선택 폴더: {folder}",
        ])

        root = tk.Tk()
        root.title("Word → Excel 변환")
        root.resizable(False, False)
        ttk.Label(
            root,
            text="선택한 폴더에 처리할 Word 파일이 없습니다.",
            padding=24,
        ).pack()
        root.after(2500, root.destroy)
        root.mainloop()
        return 1

    show_progress(folder, word_files)
    return 0


if __name__ == "__main__":
    sys.exit(main())
