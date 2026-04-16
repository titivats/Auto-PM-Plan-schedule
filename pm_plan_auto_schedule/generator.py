from __future__ import annotations

import calendar
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable

import pythoncom
import win32com.client


MONTH_ABBRS = [
    "JAN",
    "FEB",
    "MAR",
    "APR",
    "MAY",
    "JUN",
    "JUL",
    "AUG",
    "SEP",
    "OCT",
    "NOV",
    "DEC",
]

DAY_START_COL = 6
DAY_END_COL = 36
MONTH_LABEL_CELL = (16, 5)
YEAR_CELL = (17, 5)
DATE_CELL = (9, 2)
SHEET_INDEX = 1
PLAN_START_ROW = 18
PLAN_END_ROW = 52

DE_DROSS_TEXT = "DE-DROSS\n30MIN"
XL_PASTE_FORMATS = -4122


class GenerationError(Exception):
    pass


@dataclass(frozen=True)
class GeneratedFile:
    month: int
    path: Path


@dataclass(frozen=True)
class CellRef:
    row: int
    col: int


@dataclass(frozen=True)
class RowScheduleRule:
    row: int
    machine_name: str
    blank_source_col: int
    de_dross_source_col: int | None
    pm_plan_source_col: int | None
    de_dross_start_day: int | None
    pm_plan_start_day: int | None
    pm_plan_text: str | None


LogFn = Callable[[str], None]


def default_year() -> int:
    return datetime.now().year


def default_output_dir(template_path: Path, year: int) -> Path:
    return template_path.parent / f"generated-{year}"


def build_output_filename(template_path: Path, month_abbr: str, year: int) -> str:
    stem = template_path.stem
    suffix = template_path.suffix or ".xls"
    month_pattern = re.compile(
        r"\b(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b",
        re.IGNORECASE,
    )
    year_pattern = re.compile(r"\b20\d{2}\b")

    has_month = month_pattern.search(stem) is not None
    has_year = year_pattern.search(stem) is not None

    updated = month_pattern.sub(month_abbr, stem, count=1)
    updated = year_pattern.sub(str(year), updated, count=1)

    if not has_month and not has_year:
        updated = f"{stem} - {month_abbr} - {year}"

    return f"{updated}{suffix}"


def sheet_name_for(month_abbr: str, year: int) -> str:
    return f"{month_abbr} {year}"


def clear_column_contents(worksheet, col: int, start_row: int, end_row: int) -> None:
    worksheet.Range(
        worksheet.Cells(start_row, col),
        worksheet.Cells(end_row, col),
    ).ClearContents()


def excel_serial_date(year: int, month: int, day: int) -> int:
    base = datetime(1899, 12, 30)
    return (datetime(year, month, day) - base).days


def normalize_cell_text(value: object) -> str:
    return str(value or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def parse_template_month(worksheet) -> int:
    text = normalize_cell_text(worksheet.Cells(*DATE_CELL).Text)
    match = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$", text)
    if not match:
        return 1

    month_abbr = match.group(2).upper()
    if month_abbr not in MONTH_ABBRS:
        return 1
    return MONTH_ABBRS.index(month_abbr) + 1


def classify_schedule_text(text: str) -> str | None:
    normalized = normalize_cell_text(text).upper()
    if not normalized:
        return None
    if normalized == DE_DROSS_TEXT.upper():
        return "de_dross"
    if normalized.startswith("PM TEAM") or normalized.startswith("PM PLAN"):
        return "pm_plan"
    return None


def to_pm_plan_text(text: str) -> str:
    normalized = normalize_cell_text(text)
    if not normalized:
        return "PM PLAN"
    return re.sub(r"^PM\s+TEAM\b", "PM PLAN", normalized, flags=re.IGNORECASE)


def uses_pm_anchor_for_de_dross(machine_name: str) -> bool:
    return (
        re.match(r"^(?:BT0[1-9]|A12|A13)\b", machine_name.strip(), flags=re.IGNORECASE)
        is not None
    )


def extract_schedule_rules(worksheet) -> tuple[int, list[RowScheduleRule], CellRef | None]:
    template_month = parse_template_month(worksheet)
    rules: list[RowScheduleRule] = []
    default_de_dross_source: CellRef | None = None

    for row in range(PLAN_START_ROW, PLAN_END_ROW + 1, 2):
        machine_name = normalize_cell_text(worksheet.Cells(row, 2).Text)
        blank_source_col: int | None = None
        de_dross_source_col: int | None = None
        pm_plan_source_col: int | None = None
        de_dross_start_day: int | None = None
        pm_plan_start_day: int | None = None
        pm_plan_text: str | None = None

        for col in range(DAY_START_COL, DAY_END_COL + 1):
            text = normalize_cell_text(worksheet.Cells(row, col).Text)
            kind = classify_schedule_text(text)

            if kind is None and blank_source_col is None:
                blank_source_col = col
                continue

            day = col - DAY_START_COL + 1
            if kind == "de_dross":
                if de_dross_source_col is None:
                    de_dross_source_col = col
                if default_de_dross_source is None:
                    default_de_dross_source = CellRef(row=row, col=col)
                if de_dross_start_day is None:
                    de_dross_start_day = day
            elif kind == "pm_plan":
                if pm_plan_source_col is None:
                    pm_plan_source_col = col
                if pm_plan_start_day is None:
                    pm_plan_start_day = day
                if pm_plan_text is None:
                    pm_plan_text = to_pm_plan_text(text)

        if blank_source_col is None:
            blank_source_col = DAY_START_COL

        if de_dross_start_day is None and pm_plan_start_day is None:
            continue

        rules.append(
            RowScheduleRule(
                row=row,
                machine_name=machine_name,
                blank_source_col=blank_source_col,
                de_dross_source_col=de_dross_source_col,
                pm_plan_source_col=pm_plan_source_col,
                de_dross_start_day=de_dross_start_day,
                pm_plan_start_day=pm_plan_start_day,
                pm_plan_text=pm_plan_text,
            )
        )

    return template_month, rules, default_de_dross_source


def copy_cell(
    source_worksheet,
    source_row: int,
    source_col: int,
    target_worksheet,
    target_row: int,
    target_col: int,
) -> None:
    source_worksheet.Cells(source_row, source_col).Copy(
        target_worksheet.Cells(target_row, target_col)
    )


def reset_row_schedule(worksheet, template_worksheet, row: int, blank_source_col: int) -> None:
    source = template_worksheet.Cells(row, blank_source_col)
    target = worksheet.Range(
        worksheet.Cells(row, DAY_START_COL),
        worksheet.Cells(row, DAY_END_COL),
    )
    source.Copy()
    target.PasteSpecial(XL_PASTE_FORMATS)
    target.ClearContents()
    worksheet.Application.CutCopyMode = False


def iter_occurrences(start_date: date, interval_days: int, year: int, month: int):
    first_of_month = date(year, month, 1)
    last_of_month = date(year, month, calendar.monthrange(year, month)[1])

    occurrence = start_date
    if occurrence < first_of_month:
        delta_days = (first_of_month - occurrence).days
        jumps = (delta_days + interval_days - 1) // interval_days
        occurrence += timedelta(days=jumps * interval_days)

    while occurrence <= last_of_month:
        yield occurrence
        occurrence += timedelta(days=interval_days)


def iter_occurrences_from_anchor(anchor_date: date, interval_days: int, year: int, month: int):
    first_of_month = date(year, month, 1)
    last_of_month = date(year, month, calendar.monthrange(year, month)[1])

    offset = (first_of_month - anchor_date).days % interval_days
    occurrence = first_of_month if offset == 0 else first_of_month + timedelta(days=interval_days - offset)

    while occurrence <= last_of_month:
        yield occurrence
        occurrence += timedelta(days=interval_days)


def apply_schedule_rule(
    worksheet,
    template_worksheet,
    template_month: int,
    rule: RowScheduleRule,
    default_de_dross_source: CellRef | None,
    year: int,
    month: int,
) -> None:
    reset_row_schedule(worksheet, template_worksheet, rule.row, rule.blank_source_col)

    pm_plan_days: set[int] = set()
    use_pm_anchor = (
        uses_pm_anchor_for_de_dross(rule.machine_name)
        and rule.pm_plan_start_day is not None
    )

    if rule.pm_plan_start_day is not None and rule.pm_plan_source_col is not None:
        anchor_date = date(year, template_month, rule.pm_plan_start_day)
        for occurrence in iter_occurrences(anchor_date, 28, year, month):
            pm_plan_days.add(occurrence.day)
            target_col = DAY_START_COL + occurrence.day - 1
            copy_cell(
                template_worksheet,
                rule.row,
                rule.pm_plan_source_col,
                worksheet,
                rule.row,
                target_col,
            )
            worksheet.Cells(rule.row, target_col).Value = rule.pm_plan_text or "PM PLAN"

    de_dross_source_row = rule.row
    de_dross_source_col = rule.de_dross_source_col
    if de_dross_source_col is None and use_pm_anchor and default_de_dross_source is not None:
        de_dross_source_row = default_de_dross_source.row
        de_dross_source_col = default_de_dross_source.col

    if use_pm_anchor and de_dross_source_col is not None:
        anchor_date = date(year, template_month, rule.pm_plan_start_day)
        occurrences = iter_occurrences_from_anchor(anchor_date, 7, year, month)
    elif rule.de_dross_start_day is not None and de_dross_source_col is not None:
        start_date = date(year, template_month, rule.de_dross_start_day)
        occurrences = iter_occurrences(start_date, 7, year, month)
    else:
        occurrences = None

    if occurrences is not None:
        for occurrence in occurrences:
            if occurrence.day in pm_plan_days:
                continue
            target_col = DAY_START_COL + occurrence.day - 1
            copy_cell(
                template_worksheet,
                de_dross_source_row,
                de_dross_source_col,
                worksheet,
                rule.row,
                target_col,
            )
            worksheet.Cells(rule.row, target_col).Value = DE_DROSS_TEXT

    worksheet.Application.CutCopyMode = False


def configure_month(
    worksheet,
    template_worksheet,
    template_month: int,
    schedule_rules: list[RowScheduleRule],
    default_de_dross_source: CellRef | None,
    year: int,
    month: int,
    log: LogFn,
) -> None:
    month_abbr = MONTH_ABBRS[month - 1]
    days_in_month = calendar.monthrange(year, month)[1]

    worksheet.Name = sheet_name_for(month_abbr, year)
    worksheet.Cells(*MONTH_LABEL_CELL).Value = month_abbr
    worksheet.Cells(*YEAR_CELL).Value = year
    worksheet.Cells(*DATE_CELL).Value = excel_serial_date(year, month, 1)

    for col in range(DAY_START_COL, DAY_END_COL + 1):
        worksheet.Columns(col).Hidden = False

    for day in range(1, 32):
        col = DAY_START_COL + day - 1
        header_cell = worksheet.Cells(16, col)
        if day <= days_in_month:
            header_cell.Value = day
        else:
            header_cell.Value = ""
            clear_column_contents(worksheet, col, PLAN_START_ROW, 60)
            worksheet.Columns(col).Hidden = True

    for rule in schedule_rules:
        apply_schedule_rule(
            worksheet,
            template_worksheet,
            template_month,
            rule,
            default_de_dross_source,
            year,
            month,
        )

    log(f"Configured {month_abbr} {year} with {days_in_month} day(s).")


def ensure_excel_available() -> None:
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
    except Exception as exc:  # pragma: no cover
        raise GenerationError(
            "Microsoft Excel is required to generate files from this template."
        ) from exc
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def generate_year_files(
    template_path: str | os.PathLike[str],
    output_dir: str | os.PathLike[str],
    year: int,
    log: LogFn | None = None,
) -> list[GeneratedFile]:
    template = Path(template_path).expanduser().resolve()
    target_dir = Path(output_dir).expanduser().resolve()

    if not template.exists():
        raise GenerationError(f"Template file was not found: {template}")
    if template.suffix.lower() not in {".xls", ".xlsx", ".xlsm"}:
        raise GenerationError("Template must be an Excel file (.xls, .xlsx, or .xlsm).")
    if year < 1900 or year > 9999:
        raise GenerationError("Year must be between 1900 and 9999.")

    target_dir.mkdir(parents=True, exist_ok=True)
    logger = log or (lambda message: None)
    results: list[GeneratedFile] = []

    pythoncom.CoInitialize()
    excel = None
    template_workbook = None
    temp_template_copy: Path | None = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        with tempfile.NamedTemporaryFile(
            prefix="pm-plan-template-",
            suffix=template.suffix,
            dir=target_dir,
            delete=False,
        ) as handle:
            temp_template_copy = Path(handle.name)

        shutil.copy2(template, temp_template_copy)
        template_workbook = excel.Workbooks.Open(str(temp_template_copy), 0, True)
        template_worksheet = template_workbook.Worksheets(SHEET_INDEX)
        template_month, schedule_rules, default_de_dross_source = extract_schedule_rules(
            template_worksheet
        )

        for month in range(1, 13):
            month_abbr = MONTH_ABBRS[month - 1]
            output_name = build_output_filename(template, month_abbr, year)
            output_path = target_dir / output_name

            if output_path.resolve() == template:
                raise GenerationError(
                    "Output path would overwrite the template file. Choose a different output folder."
                )

            shutil.copy2(template, output_path)
            logger(f"Copied template to {output_path.name}")

            workbook = None
            try:
                workbook = excel.Workbooks.Open(str(output_path))
                worksheet = workbook.Worksheets(SHEET_INDEX)
                configure_month(
                    worksheet,
                    template_worksheet,
                    template_month,
                    schedule_rules,
                    default_de_dross_source,
                    year,
                    month,
                    logger,
                )
                workbook.Save()
                results.append(GeneratedFile(month=month, path=output_path))
                logger(f"Saved {output_path.name}")
            finally:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
    except GenerationError:
        raise
    except Exception as exc:
        raise GenerationError(f"Failed while generating files: {exc}") from exc
    finally:
        if template_workbook is not None:
            template_workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
        if temp_template_copy is not None and temp_template_copy.exists():
            temp_template_copy.unlink(missing_ok=True)

    return results
