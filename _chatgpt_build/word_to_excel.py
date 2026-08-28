# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import sys
import time
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

import pythoncom
import win32com.client

WORD_EXTENSIONS = {".doc", ".docx", ".docm"}
OUTPUT_BASENAME = "Word파일_통합"
OUTPUT_EXTENSION = ".xlsx"
ERROR_LOG_NAME = "WordToExcel_error.log"
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def choose_folder() -> Path | None:
    """화면에는 폴더 선택 UI 하나만 표시한다."""
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


def unique_output_path(folder: Path) -> Path:
    candidate = folder / f"{OUTPUT_BASENAME}{OUTPUT_EXTENSION}"
    if not candidate.exists():
        return candidate

    index = 2
    while True:
        candidate = folder / f"{OUTPUT_BASENAME}_{index}{OUTPUT_EXTENSION}"
        if not candidate.exists():
            return candidate
        index += 1


def normalize_sheet_base(filename_stem: str) -> str:
    name = INVALID_SHEET_CHARS.sub("_", filename_stem)
    name = "".join(ch for ch in name if ord(ch) >= 32)
    name = name.strip().strip("'").strip()
    return name or "Sheet"


def unique_sheet_name(filename_stem: str, used: set[str]) -> str:
    base = normalize_sheet_base(filename_stem)
    candidate = base[:31]

    if candidate.casefold() not in used:
        used.add(candidate.casefold())
        return candidate

    index = 2
    while True:
        suffix = f"_{index}"
        candidate = f"{base[:31-len(suffix)]}{suffix}"
        if candidate.casefold() not in used:
            used.add(candidate.casefold())
            return candidate
        index += 1


def write_error_log(folder: Path, lines: list[str]) -> None:
    if not lines:
        return
    (folder / ERROR_LOG_NAME).write_text(
        "\n".join(lines).rstrip() + "\n",
        encoding="utf-8-sig",
    )


def paste_word_document_to_sheet(word_doc, excel_app, worksheet) -> None:
    """요구사항대로 Word COM Copy -> Excel COM Paste를 수행한다."""
    word_doc.Content.Copy()
    pythoncom.PumpWaitingMessages()
    time.sleep(0.15)

    worksheet.Activate()
    target = worksheet.Range("A1")
    target.Select()

    try:
        worksheet.Paste(Destination=target)
    except Exception:
        # Office 버전에 따라 Worksheet.Paste의 Destination 인자를 거부하는 경우 대비
        excel_app.ActiveSheet.Paste()

    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

    try:
        excel_app.CutCopyMode = False
    except Exception:
        pass


def convert_folder(folder: Path) -> Path | None:
    word_files = find_word_files(folder)
    if not word_files:
        write_error_log(folder, [
            "처리할 Word 파일을 찾지 못했습니다.",
            "대상 확장자: .doc, .docx, .docm",
            f"선택 폴더: {folder}",
        ])
        return None

    output_path = unique_output_path(folder)
    errors: list[str] = []
    success_count = 0

    word = None
    excel = None
    workbook = None

    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        workbook = excel.Workbooks.Add()
        while workbook.Worksheets.Count > 1:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()

        used_sheet_names: set[str] = set()
        first_sheet_available = True

        for word_path in word_files:
            doc = None
            worksheet = None
            created_new_sheet = False

            try:
                if first_sheet_available:
                    worksheet = workbook.Worksheets(1)
                    first_sheet_available = False
                else:
                    worksheet = workbook.Worksheets.Add(
                        After=workbook.Worksheets(workbook.Worksheets.Count)
                    )
                    created_new_sheet = True

                worksheet.Name = unique_sheet_name(word_path.stem, used_sheet_names)

                doc = word.Documents.Open(
                    FileName=str(word_path.resolve()),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                    ConfirmConversions=False,
                )

                paste_word_document_to_sheet(doc, excel, worksheet)
                success_count += 1

            except Exception as exc:
                errors.append("\n".join([
                    f"[실패] {word_path.name}",
                    f"{type(exc).__name__}: {exc}",
                    traceback.format_exc(),
                ]))

                try:
                    # 실패 시 만들어진 빈 시트는 제거하되, 워크북에 최소 1개는 유지
                    if created_new_sheet and worksheet is not None and workbook.Worksheets.Count > 1:
                        worksheet.Delete()
                except Exception:
                    pass

            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass

        if success_count == 0:
            errors.insert(0, "모든 Word 파일 처리에 실패했습니다.")
            write_error_log(folder, errors)
            return None

        workbook.SaveAs(str(output_path.resolve()), FileFormat=51)  # xlOpenXMLWorkbook

        if errors:
            errors.insert(
                0,
                f"총 {len(word_files)}개 중 {success_count}개 성공, "
                f"{len(word_files)-success_count}개 실패했습니다.",
            )
            write_error_log(folder, errors)
        else:
            old_log = folder / ERROR_LOG_NAME
            if old_log.exists():
                try:
                    old_log.unlink()
                except Exception:
                    pass

        return output_path

    except Exception as exc:
        errors.insert(0, "\n".join([
            "[프로그램 오류]",
            f"{type(exc).__name__}: {exc}",
            traceback.format_exc(),
        ]))
        write_error_log(folder, errors)
        return None

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


def main() -> int:
    folder = choose_folder()
    if folder is None:
        return 0
    return 0 if convert_folder(folder) is not None else 1


if __name__ == "__main__":
    sys.exit(main())
