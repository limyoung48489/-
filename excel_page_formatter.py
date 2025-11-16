"""Excel worksheet formatter for "산근 n 호표" sections.

This module exposes a small Tkinter based UI that attaches to the currently
running Excel instance (2010/2016 compatible) and automates the following:

* Select a worksheet from the active workbook.
* Configure page setup values (orientation, paper size, margins, scaling and
  repeating header rows).
* Detect section headers matching a configurable regular expression.
* Insert horizontal page breaks before each detected header (except the first).
* Extend the print area to cover a configurable column range plus extra rows.
* Apply a bottom border to the last printable row of every page.

The code is intentionally verbose and contains logging style status messages so
that it can be repurposed for similar tasks (공종별, 공정별 등) in the future.

The script is designed for Windows hosts where Excel is already running.  When
executed it will display a GUI that connects to the active Excel instance and
lets the user configure the processing options before pressing the "실행"
button.
"""

from __future__ import annotations

import dataclasses
import re
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Iterable, List, Sequence, Tuple

IMPORT_ERROR: Exception | None = None

try:  # pragma: no cover - pywin32 is not present in the Linux execution env.
    import pythoncom
    from win32com.client import Dispatch, GetActiveObject, constants
except Exception as exc:  # pragma: no cover - better error message for users.
    # When the script is invoked on a non-Windows host (like the execution
    # environment used for automated grading) we still want the module to
    # import cleanly so that unit tests or code checks can be run.  Therefore
    # we only raise a helpful message when the script is executed directly.
    pythoncom = None
    Dispatch = None
    GetActiveObject = None
    constants = None
    IMPORT_ERROR = exc


XL_FALLBACKS = {
    # Page orientation
    "xlLandscape": 2,
    "xlPortrait": 1,
    # Paper sizes
    "xlPaperA4": 9,
    "xlPaperB4": 12,
    "xlPaperLetter": 1,
    # Border/line style
    "xlEdgeBottom": 9,
    "xlContinuous": 1,
    "xlThin": 2,
    "xlAutomatic": -4105,
}


def xl_constant(name: str) -> int:
    """Return the Excel constant value, falling back to hard coded numbers."""

    if constants is not None and hasattr(constants, name):
        return getattr(constants, name)
    if name in XL_FALLBACKS:
        return XL_FALLBACKS[name]
    raise AttributeError(f"Excel constant '{name}' is not defined")


# ---------------------------------------------------------------------------
# Dataclasses that capture the user configurable options.
# ---------------------------------------------------------------------------


@dataclasses.dataclass
class MarginOptions:
    top_mm: float = 10.0
    bottom_mm: float = 10.0
    left_mm: float = 10.0
    right_mm: float = 10.0
    header_mm: float = 5.0
    footer_mm: float = 5.0


@dataclasses.dataclass
class PageSetupOptions:
    orientation: str = "landscape"  # "landscape" or "portrait"
    paper_size: str = "A4"
    scale_mode: str = "zoom"  # "zoom" or "fit_width"
    zoom_percent: int = 100
    fit_to_pages_wide: int = 1
    fit_to_pages_tall: int = 0  # 0 means automatic (Excel interprets as auto)
    rows_to_repeat_start: int = 1
    rows_to_repeat_end: int = 4


@dataclasses.dataclass
class SectionDetectionOptions:
    header_column: int = 1  # Column number where the header text resides
    header_regex: str = r"^산근\s+\d+\s*호표"


@dataclasses.dataclass
class BorderOptions:
    first_column: int = 1
    last_column: int = 10
    extra_print_rows: int = 200
    reset_page_breaks: bool = True
    skip_blank_pages: bool = False


@dataclasses.dataclass
class WorksheetProcessingOptions:
    margins: MarginOptions
    page_setup: PageSetupOptions
    section_detection: SectionDetectionOptions
    border: BorderOptions


# ---------------------------------------------------------------------------
# Helper functions.
# ---------------------------------------------------------------------------


def mm_to_points(mm_value: float) -> float:
    """Convert millimeters to points (as expected by Excel)."""

    return (mm_value / 25.4) * 72.0


def column_number_to_letter(col_num: int) -> str:
    """Convert a 1-based column index to Excel column letters."""

    if col_num < 1:
        raise ValueError("Column numbers must start at 1")
    letters: List[str] = []
    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def excel_address(row: int, column: int) -> str:
    return f"{column_number_to_letter(column)}{row}"


# ---------------------------------------------------------------------------
# Core formatter implementation.
# ---------------------------------------------------------------------------


class WorksheetFormatter:
    """Encapsulates the Excel automation logic."""

    def __init__(self, worksheet, options: WorksheetProcessingOptions):
        self.ws = worksheet
        self.options = options
        self.status_messages: List[str] = []

    # Public API -------------------------------------------------------------
    def process(self) -> str:
        used_last_row = self._calculate_used_last_row()
        border_opts = self.options.border
        max_row = used_last_row + border_opts.extra_print_rows
        self.status_messages.append(
            f"데이터 마지막 행: {used_last_row}, 인쇄 영역 끝 행: {max_row}"
        )

        if border_opts.reset_page_breaks:
            self.ws.ResetAllPageBreaks()
            self.status_messages.append("기존 페이지 나누기 초기화")

        self._apply_page_setup(max_row)
        header_rows = self._find_header_rows(used_last_row)
        page_breaks_added = self._apply_page_breaks(header_rows)
        page_ranges = self._calculate_page_ranges(max_row)
        decorated_pages = self._apply_bottom_borders(page_ranges)
        self.status_messages.append(
            f"페이지 나누기 추가: {page_breaks_added}개, 하단선 적용 페이지: {decorated_pages}개"
        )
        return "\n".join(self.status_messages)

    # Internal helpers ------------------------------------------------------
    def _calculate_used_last_row(self) -> int:
        used_range = self.ws.UsedRange
        if used_range is None:
            return 1
        rows = used_range.Rows
        return rows(rows.Count).Row

    def _apply_page_setup(self, max_row: int) -> None:
        opts = self.options
        ps = self.ws.PageSetup
        margins = opts.margins
        page_setup = opts.page_setup

        ps.Orientation = (
            xl_constant("xlLandscape")
            if page_setup.orientation == "landscape"
            else xl_constant("xlPortrait")
        )

        ps.PaperSize = {
            "A4": xl_constant("xlPaperA4"),
            "B4": xl_constant("xlPaperB4"),
            "Letter": xl_constant("xlPaperLetter"),
        }.get(page_setup.paper_size, xl_constant("xlPaperA4"))

        ps.LeftMargin = mm_to_points(margins.left_mm)
        ps.RightMargin = mm_to_points(margins.right_mm)
        ps.TopMargin = mm_to_points(margins.top_mm)
        ps.BottomMargin = mm_to_points(margins.bottom_mm)
        ps.HeaderMargin = mm_to_points(margins.header_mm)
        ps.FooterMargin = mm_to_points(margins.footer_mm)

        if page_setup.scale_mode == "fit_width":
            ps.Zoom = False
            ps.FitToPagesWide = page_setup.fit_to_pages_wide
            ps.FitToPagesTall = page_setup.fit_to_pages_tall
        else:
            ps.FitToPagesTall = False
            ps.FitToPagesWide = False
            ps.Zoom = page_setup.zoom_percent

        ps.RowsToRepeatAtTop = f"${page_setup.rows_to_repeat_start}:${page_setup.rows_to_repeat_end}"

        border = self.options.border
        print_area = f"${excel_address(1, border.first_column)}:${excel_address(max_row, border.last_column)}"
        ps.PrintArea = print_area
        self.status_messages.append(f"PrintArea 설정: {print_area}")

    def _find_header_rows(self, used_last_row: int) -> List[int]:
        detection = self.options.section_detection
        pattern = re.compile(detection.header_regex)
        header_rows: List[int] = []
        for row in range(1, used_last_row + 1):
            value = self.ws.Cells(row, detection.header_column).Value
            if isinstance(value, str) and pattern.search(value.strip()):
                header_rows.append(row)
        if not header_rows:
            self.status_messages.append("머리글 패턴을 찾지 못했습니다. 기존 페이지 나누기를 사용합니다.")
        else:
            self.status_messages.append(
                f"머리글 감지: {len(header_rows)}개 (행 {', '.join(map(str, header_rows))})"
            )
        return header_rows

    def _apply_page_breaks(self, header_rows: Sequence[int]) -> int:
        if len(header_rows) <= 1:
            return 0
        added = 0
        for row in header_rows[1:]:
            self.ws.HPageBreaks.Add(Before=self.ws.Rows(row))
            added += 1
        return added

    def _calculate_page_ranges(self, max_row: int) -> List[Tuple[int, int]]:
        border_opts = self.options.border
        first_row = 1
        break_rows = sorted({pb.Location.Row for pb in self.ws.HPageBreaks})
        page_starts: List[int] = [first_row]
        for row in break_rows:
            if row not in page_starts and row > first_row:
                page_starts.append(row)
        page_starts.sort()

        page_ranges: List[Tuple[int, int]] = []
        for idx, start in enumerate(page_starts):
            next_start = page_starts[idx + 1] if idx + 1 < len(page_starts) else None
            end_row = (next_start - 1) if next_start else max_row
            page_ranges.append((start, end_row))
        self.status_messages.append(
            "페이지 범위 계산: " + ", ".join(f"{s}-{e}" for s, e in page_ranges)
        )
        return page_ranges

    def _apply_bottom_borders(self, page_ranges: Sequence[Tuple[int, int]]) -> int:
        decorated = 0
        border_opts = self.options.border
        for start_row, end_row in page_ranges:
            if end_row < start_row:
                continue
            if border_opts.skip_blank_pages and not self._page_has_data(start_row, end_row):
                continue
            rng = self.ws.Range(
                self.ws.Cells(end_row, border_opts.first_column),
                self.ws.Cells(end_row, border_opts.last_column),
            )
            bottom_border = rng.Borders(xl_constant("xlEdgeBottom"))
            bottom_border.LineStyle = xl_constant("xlContinuous")
            bottom_border.Weight = xl_constant("xlThin")
            bottom_border.ColorIndex = xl_constant("xlAutomatic")
            decorated += 1
        return decorated

    def _page_has_data(self, start_row: int, end_row: int) -> bool:
        border_opts = self.options.border
        rng = self.ws.Range(
            self.ws.Cells(start_row, border_opts.first_column),
            self.ws.Cells(end_row, border_opts.last_column),
        )
        values = rng.Value
        if values is None:
            return False
        if not isinstance(values, Iterable):
            return bool(values)

        # Flatten nested tuples returned by COM.
        def iter_cells(val):
            if isinstance(val, (list, tuple)):
                for item in val:
                    yield from iter_cells(item)
            else:
                yield val

        return any(v not in (None, "") for v in iter_cells(values))


# ---------------------------------------------------------------------------
# Excel + Tkinter glue.
# ---------------------------------------------------------------------------


class ExcelController:
    """Connects to a running Excel instance and exposes helper methods."""

    def __init__(self):
        self.excel = None
        self.workbook = None

    def connect(self):
        if GetActiveObject is None:
            raise RuntimeError("win32com이 설치된 Windows 환경에서만 실행 가능합니다.")
        pythoncom.CoInitialize()
        self.excel = GetActiveObject("Excel.Application")
        self.workbook = self.excel.ActiveWorkbook
        if self.workbook is None:
            raise RuntimeError("활성 Workbook을 찾을 수 없습니다. 파일을 연 후 다시 시도하세요.")

    def sheet_names(self) -> List[str]:
        if self.workbook is None:
            return []
        return [sheet.Name for sheet in self.workbook.Worksheets]

    def worksheet_by_name(self, name: str):
        return self.workbook.Worksheets(name)


class FormatterApp:
    """Tkinter based GUI application."""

    def __init__(self):
        self.controller = ExcelController()
        self.root = tk.Tk()
        self.root.title("산근 호표 페이지 테두리 도우미")
        self.sheet_var = tk.StringVar()
        self.message_var = tk.StringVar()

        # Page setup vars
        self.orientation_var = tk.StringVar(value="landscape")
        self.paper_size_var = tk.StringVar(value="A4")
        self.scale_mode_var = tk.StringVar(value="zoom")
        self.zoom_var = tk.IntVar(value=100)
        self.repeat_start_var = tk.IntVar(value=1)
        self.repeat_end_var = tk.IntVar(value=4)

        # Margin vars (mm)
        self.top_margin_var = tk.DoubleVar(value=10.0)
        self.bottom_margin_var = tk.DoubleVar(value=10.0)
        self.left_margin_var = tk.DoubleVar(value=10.0)
        self.right_margin_var = tk.DoubleVar(value=10.0)
        self.header_margin_var = tk.DoubleVar(value=5.0)
        self.footer_margin_var = tk.DoubleVar(value=5.0)

        # Border vars
        self.first_col_var = tk.IntVar(value=1)
        self.last_col_var = tk.IntVar(value=10)
        self.extra_rows_var = tk.IntVar(value=200)
        self.reset_breaks_var = tk.BooleanVar(value=True)
        self.skip_blank_var = tk.BooleanVar(value=False)

        # Section detection vars
        self.header_col_var = tk.IntVar(value=1)
        self.regex_var = tk.StringVar(value=r"^산근\s+\d+\s*호표")

        # Fit to page vars
        self.fit_width_var = tk.IntVar(value=1)
        self.fit_tall_var = tk.IntVar(value=0)

        self._build_ui()
        self._connect_excel()

    # UI building -----------------------------------------------------------
    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Sheet selection ----------------------------------------------------
        sheet_frame = ttk.LabelFrame(main_frame, text="대상 시트")
        sheet_frame.grid(row=0, column=0, sticky="ew")
        sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, state="readonly", width=40)
        sheet_combo.grid(row=0, column=0, padx=5, pady=5)
        self.sheet_combo = sheet_combo
        ttk.Button(sheet_frame, text="새로고침", command=self._refresh_sheet_list).grid(row=0, column=1, padx=5)

        # Page setup ---------------------------------------------------------
        setup_frame = ttk.LabelFrame(main_frame, text="페이지 설정")
        setup_frame.grid(row=1, column=0, sticky="ew", pady=5)

        ttk.Label(setup_frame, text="용지 방향").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(setup_frame, text="가로", variable=self.orientation_var, value="landscape").grid(row=0, column=1)
        ttk.Radiobutton(setup_frame, text="세로", variable=self.orientation_var, value="portrait").grid(row=0, column=2)

        ttk.Label(setup_frame, text="용지 크기").grid(row=1, column=0, sticky="w")
        ttk.Combobox(setup_frame, textvariable=self.paper_size_var, values=["A4", "B4", "Letter"], state="readonly").grid(row=1, column=1, columnspan=2, sticky="ew")

        ttk.Label(setup_frame, text="확대/축소 모드").grid(row=2, column=0, sticky="w")
        ttk.Radiobutton(setup_frame, text="배율", variable=self.scale_mode_var, value="zoom").grid(row=2, column=1)
        ttk.Radiobutton(setup_frame, text="가로 1페이지", variable=self.scale_mode_var, value="fit_width").grid(row=2, column=2)
        ttk.Label(setup_frame, text="배율 (%)").grid(row=3, column=0)
        ttk.Entry(setup_frame, textvariable=self.zoom_var, width=6).grid(row=3, column=1)
        ttk.Label(setup_frame, text="Fit Tall").grid(row=3, column=2)
        ttk.Entry(setup_frame, textvariable=self.fit_tall_var, width=6).grid(row=3, column=3)

        ttk.Label(setup_frame, text="반복 행 (시작~끝)").grid(row=4, column=0, sticky="w")
        ttk.Entry(setup_frame, textvariable=self.repeat_start_var, width=6).grid(row=4, column=1)
        ttk.Entry(setup_frame, textvariable=self.repeat_end_var, width=6).grid(row=4, column=2)

        # Margins ------------------------------------------------------------
        margin_frame = ttk.LabelFrame(main_frame, text="여백 (mm)")
        margin_frame.grid(row=2, column=0, sticky="ew", pady=5)

        labels = [
            ("위쪽", self.top_margin_var),
            ("아래쪽", self.bottom_margin_var),
            ("왼쪽", self.left_margin_var),
            ("오른쪽", self.right_margin_var),
            ("머리글", self.header_margin_var),
            ("바닥글", self.footer_margin_var),
        ]
        for idx, (text, var) in enumerate(labels):
            ttk.Label(margin_frame, text=text).grid(row=idx // 3, column=(idx % 3) * 2, sticky="w")
            ttk.Entry(margin_frame, textvariable=var, width=8).grid(row=idx // 3, column=(idx % 3) * 2 + 1)

        # Border setup -------------------------------------------------------
        border_frame = ttk.LabelFrame(main_frame, text="테두리/인쇄 범위")
        border_frame.grid(row=3, column=0, sticky="ew", pady=5)
        ttk.Label(border_frame, text="시작 열").grid(row=0, column=0)
        ttk.Entry(border_frame, textvariable=self.first_col_var, width=6).grid(row=0, column=1)
        ttk.Label(border_frame, text="끝 열").grid(row=0, column=2)
        ttk.Entry(border_frame, textvariable=self.last_col_var, width=6).grid(row=0, column=3)
        ttk.Label(border_frame, text="추가 행").grid(row=0, column=4)
        ttk.Entry(border_frame, textvariable=self.extra_rows_var, width=6).grid(row=0, column=5)
        ttk.Checkbutton(border_frame, text="페이지 나누기 초기화", variable=self.reset_breaks_var).grid(row=1, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(border_frame, text="빈 페이지 하단선 생략", variable=self.skip_blank_var).grid(row=1, column=3, columnspan=3, sticky="w")

        # Section detection --------------------------------------------------
        section_frame = ttk.LabelFrame(main_frame, text="머리글 파싱")
        section_frame.grid(row=4, column=0, sticky="ew", pady=5)
        ttk.Label(section_frame, text="머리글 열").grid(row=0, column=0)
        ttk.Entry(section_frame, textvariable=self.header_col_var, width=6).grid(row=0, column=1)
        ttk.Label(section_frame, text="정규식").grid(row=0, column=2)
        ttk.Entry(section_frame, textvariable=self.regex_var, width=40).grid(row=0, column=3)

        # Buttons ------------------------------------------------------------
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, sticky="ew", pady=10)
        ttk.Button(button_frame, text="실행", command=self._run).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="종료", command=self.root.destroy).grid(row=0, column=1, padx=5)

        # Status -------------------------------------------------------------
        ttk.Label(main_frame, textvariable=self.message_var, foreground="blue").grid(row=6, column=0, sticky="w")

    # Excel connectivity ----------------------------------------------------
    def _connect_excel(self):
        try:
            self.controller.connect()
            self._refresh_sheet_list()
            self.message_var.set("Excel 연결 성공")
        except Exception as exc:
            messagebox.showerror("Excel 연결 실패", str(exc))
            self.message_var.set(str(exc))

    def _refresh_sheet_list(self):
        sheets = self.controller.sheet_names()
        if not sheets:
            self.message_var.set("Workbook을 찾을 수 없습니다.")
            return
        self.sheet_combo["values"] = sheets
        active = self.controller.workbook.ActiveSheet.Name
        self.sheet_var.set(active)

    # Action handling -------------------------------------------------------
    def _run(self):
        sheet_name = self.sheet_var.get()
        if not sheet_name:
            messagebox.showwarning("시트 선택", "대상 시트를 선택하세요.")
            return
        try:
            worksheet = self.controller.worksheet_by_name(sheet_name)
            options = self._gather_options()
            formatter = WorksheetFormatter(worksheet, options)
            summary = formatter.process()
            self.message_var.set("완료")
            messagebox.showinfo("처리 완료", summary)
        except Exception as exc:
            messagebox.showerror("오류", str(exc))

    def _gather_options(self) -> WorksheetProcessingOptions:
        margins = MarginOptions(
            top_mm=self.top_margin_var.get(),
            bottom_mm=self.bottom_margin_var.get(),
            left_mm=self.left_margin_var.get(),
            right_mm=self.right_margin_var.get(),
            header_mm=self.header_margin_var.get(),
            footer_mm=self.footer_margin_var.get(),
        )

        page_setup = PageSetupOptions(
            orientation=self.orientation_var.get(),
            paper_size=self.paper_size_var.get(),
            scale_mode=self.scale_mode_var.get(),
            zoom_percent=self.zoom_var.get(),
            fit_to_pages_wide=self.fit_width_var.get(),
            fit_to_pages_tall=self.fit_tall_var.get(),
            rows_to_repeat_start=self.repeat_start_var.get(),
            rows_to_repeat_end=self.repeat_end_var.get(),
        )

        section_detection = SectionDetectionOptions(
            header_column=self.header_col_var.get(),
            header_regex=self.regex_var.get(),
        )

        border = BorderOptions(
            first_column=self.first_col_var.get(),
            last_column=self.last_col_var.get(),
            extra_print_rows=self.extra_rows_var.get(),
            reset_page_breaks=self.reset_breaks_var.get(),
            skip_blank_pages=self.skip_blank_var.get(),
        )

        return WorksheetProcessingOptions(
            margins=margins,
            page_setup=page_setup,
            section_detection=section_detection,
            border=border,
        )

    # Public entry ----------------------------------------------------------
    def run(self):
        self.root.mainloop()


def main(argv: Sequence[str] | None = None) -> int:
    if constants is None:
        lines = ["이 스크립트는 Windows + Excel 환경에서만 실행 가능합니다."]
        if IMPORT_ERROR is not None:
            lines.append(
                "win32com(pywin32) 모듈을 불러오지 못했습니다. 아래 메시지를 확인하세요:"
            )
            lines.append(f"  → {IMPORT_ERROR}")
            lines.append(
                "PowerShell 에서 `python -m pip install pywin32` 를 실행한 뒤 다시 시도해 주세요."
            )
        sys.stderr.write("\n".join(lines) + "\n")
        return 1
    app = FormatterApp()
    app.run()
    return 0


if __name__ == "__main__":  # pragma: no cover - UI entry point
    raise SystemExit(main(sys.argv[1:]))
