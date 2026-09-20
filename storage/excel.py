from datetime import datetime
from pathlib import Path
import shutil
import os
import tempfile

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

from config import DATA_DIR, TRACKER_PATH


# ============================================================================
# Excel tracker structure
# ============================================================================

JOB_HEADERS = [
    "Job ID",
    "Company",
    "Job Title",
    "Job URL",
    "Company URL",
    "Location",
    "Posted Date",
    "Posted Text",
    "Job Age",
    "Experience",
    "Min Experience",
    "Max Experience",
    "Employment Type",
    "Salary",
    "Min LPA",
    "Max LPA",
    "Employees",
    "Min Employees",
    "Max Employees",
    "AI Score",
    "Full Stack Score",
    "Relevance Score",
    "Resume Type",
    "Resume Name",
    "Resume Path",
    "Matched Skills",
    "Search Terms",
    "Application Status",
    "Referral Status",
    "Referral Source",
    "Referral Source Role",
    "Referral Date",
    "Created",
    "Updated",
]

CONTACT_HEADERS = [
    "Job ID",
    "Company",
    "Person",
    "LinkedIn URL",
    "Role",
    "Connection Degree",
    "Connection Count",
    "Activity",
    "Contact Score",
    "Referral Priority Score",
    "High Value Role",
    "Contact Status",
    "Request Status",
    "Response Status",
    "Referral Status",
    "Last Contacted",
    "Last Response",
    "Notes",
    "Created",
    "Updated",
]

REFERRAL_HEADERS = [
    "Job ID",
    "Company",
    "Person",
    "Role",
    "Referral Requested",
    "Referral Offered",
    "Referral Received",
    "Referral Date",
    "Is Referral Source",
    "Status",
    "Notes",
    "Created",
    "Updated",
]

MESSAGE_HEADERS = [
    "Job ID",
    "Company",
    "Person",
    "Message Type",
    "Message",
    "Resume Type",
    "Resume Name",
    "Resume Path",
    "Job URL",
    "Attachment Required",
    "Status",
    "Created Date",
    "Sent Date",
]

ACTIVITY_HEADERS = [
    "Job ID",
    "Company",
    "Person",
    "Activity",
    "Details",
    "Date",
]

DASHBOARD_HEADERS = ["Metric", "Value"]


# ============================================================================
# Formatting
# ============================================================================

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SUBHEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
SUBHEADER_FONT = Font(bold=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9E1F2"),
    right=Side(style="thin", color="D9E1F2"),
    top=Side(style="thin", color="D9E1F2"),
    bottom=Side(style="thin", color="D9E1F2"),
)

STATUS_FILLS = {
    "received": PatternFill("solid", fgColor="C6EFCE"),
    "willing to refer": PatternFill("solid", fgColor="C6EFCE"),
    "offered": PatternFill("solid", fgColor="C6EFCE"),
    "applied": PatternFill("solid", fgColor="C6EFCE"),
    "interview": PatternFill("solid", fgColor="C6EFCE"),
    "offer": PatternFill("solid", fgColor="C6EFCE"),
    "declined": PatternFill("solid", fgColor="FFC7CE"),
    "rejected": PatternFill("solid", fgColor="FFC7CE"),
    "withdrawn": PatternFill("solid", fgColor="FFC7CE"),
    "no response": PatternFill("solid", fgColor="FFF2CC"),
    "needs review": PatternFill("solid", fgColor="FFF2CC"),
    "prepared": PatternFill("solid", fgColor="D9EAF7"),
    "contacted": PatternFill("solid", fgColor="D9EAF7"),
    "requested": PatternFill("solid", fgColor="D9EAF7"),
    "referral requested": PatternFill("solid", fgColor="D9EAF7"),
    "connection requested": PatternFill("solid", fgColor="D9EAF7"),
}


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _ensure_data_dir():
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def _sheet_definitions():
    return {
        "Jobs": JOB_HEADERS,
        "Contacts": CONTACT_HEADERS,
        "Referrals": REFERRAL_HEADERS,
        "Messages": MESSAGE_HEADERS,
        "Activity": ACTIVITY_HEADERS,
    }


def _create_sheet_if_missing(workbook, sheet_name, headers):
    if sheet_name not in workbook.sheetnames:
        sheet = workbook.create_sheet(sheet_name)
        sheet.append(headers)
        return sheet

    sheet = workbook[sheet_name]

    # A sheet may exist but be empty. Add the header row in that case.
    if sheet.max_row == 1 and all(
        cell.value is None for cell in sheet[1]
    ):
        sheet.delete_rows(1, 1)
        sheet.append(headers)

    return sheet


def _backup_broken_tracker(path):
    """Keep a copy before recreating an unreadable/corrupt workbook."""
    if not path.exists():
        return ""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = path.with_name(
        f"{path.stem}_backup_{timestamp}{path.suffix}"
    )

    try:
        shutil.copy2(path, backup)
        return str(backup)
    except Exception:
        return ""


def _atomic_save_workbook(workbook, path):
    """
    Save the workbook safely without partially overwriting the live tracker.

    The workbook is written to a temporary .xlsx in the same directory first.
    After the temporary workbook is successfully written and validated, it is
    atomically replaced into the live tracker path.

    If the process is interrupted while saving, the existing live tracker
    remains untouched.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = None

    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{path.stem}_",
            suffix=".tmp.xlsx",
            dir=str(path.parent),
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)

        # Write the complete workbook to the temporary file.
        workbook.save(temp_path)

        # Validate the generated OOXML before replacing the live tracker.
        validation_workbook = load_workbook(temp_path, read_only=True)
        validation_workbook.close()

        # os.replace() is atomic when source and destination are on the same
        # filesystem, which they are because the temp file is created beside
        # the live tracker.
        os.replace(temp_path, path)
        temp_path = None

    finally:
        # Never remove or modify the live tracker here. Only clean up the
        # temporary file if the save/validation was interrupted or failed.
        if temp_path is not None:
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except Exception:
                pass


def _load_or_create_workbook(path):
    """
    Load a valid workbook or create a fresh one.

    A backup is made if the existing workbook cannot be opened.
    """
    path = Path(path)

    if not path.exists():
        return Workbook()

    try:
        workbook = load_workbook(path)
    except Exception:
        _backup_broken_tracker(path)
        return Workbook()

    # OpenPyXL cannot save a workbook with all sheets hidden.
    visible = [s for s in workbook.worksheets if s.sheet_state == "visible"]
    if not visible:
        workbook.worksheets[0].sheet_state = "visible"

    return workbook


def _set_header_style(sheet):
    if sheet.max_row < 1:
        return

    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        cell.border = THIN_BORDER

    sheet.row_dimensions[1].height = 30
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions


def _safe_table_name(sheet_name):
    return "".join(
        character if character.isalnum() else "_"
        for character in sheet_name
    ) + "Table"


def _apply_table(sheet):
    """
    Add an Excel table when there is at least a header row and one data row.
    Existing project tables are removed/recreated so the range stays correct.
    """
    if sheet.max_row < 2 or sheet.max_column < 1:
        return

    # Remove old tables safely.
    for table_name in list(sheet.tables.keys()):
        del sheet.tables[table_name]

    ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"
    table = Table(displayName=_safe_table_name(sheet.title), ref=ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style
    sheet.add_table(table)


def _set_hyperlinks(sheet):
    headers = {
        str(cell.value).strip(): cell.column
        for cell in sheet[1]
        if cell.value is not None
    }

    url_columns = [
        headers.get("Job URL"),
        headers.get("Company URL"),
        headers.get("LinkedIn URL"),
    ]

    for column_index in [c for c in url_columns if c]:
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            value = str(cell.value or "").strip()

            if value.startswith(("http://", "https://")):
                cell.hyperlink = value
                cell.style = "Hyperlink"


def _apply_number_formats(sheet):
    headers = {
        str(cell.value).strip(): cell.column
        for cell in sheet[1]
        if cell.value is not None
    }

    integer_columns = [
        "Connection Count",
        "Employees",
        "Min Employees",
        "Max Employees",
    ]

    decimal_columns = [
        "Min Experience",
        "Max Experience",
        "Min LPA",
        "Max LPA",
        "AI Score",
        "Full Stack Score",
        "Relevance Score",
        "Contact Score",
        "Referral Priority Score",
    ]

    for name in integer_columns:
        column = headers.get(name)
        if not column:
            continue
        for row in range(2, sheet.max_row + 1):
            sheet.cell(row=row, column=column).number_format = "0"

    for name in decimal_columns:
        column = headers.get(name)
        if not column:
            continue
        for row in range(2, sheet.max_row + 1):
            sheet.cell(row=row, column=column).number_format = "0.0"


def _apply_status_formatting(sheet):
    if sheet.max_row < 2:
        return

    status_columns = []
    for cell in sheet[1]:
        if str(cell.value or "").strip() in {
            "Application Status",
            "Referral Status",
            "Contact Status",
            "Request Status",
            "Response Status",
            "Status",
        }:
            status_columns.append(cell.column)

    for column in status_columns:
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column)
            value = str(cell.value or "").strip().lower()

            if value in STATUS_FILLS:
                cell.fill = STATUS_FILLS[value]


def _apply_score_formatting(sheet):
    header_map = {
        str(cell.value).strip(): cell.column
        for cell in sheet[1]
        if cell.value is not None
    }

    for name in [
        "AI Score",
        "Full Stack Score",
        "Relevance Score",
        "Contact Score",
        "Referral Priority Score",
    ]:
        column = header_map.get(name)
        if not column or sheet.max_row < 2:
            continue

        start = f"{get_column_letter(column)}2"
        end = f"{get_column_letter(column)}{sheet.max_row}"

        sheet.conditional_formatting.add(
            f"{start}:{end}",
            ColorScaleRule(
                start_type="min",
                start_color="F8696B",
                mid_type="percentile",
                mid_value=50,
                mid_color="FFEB84",
                end_type="max",
                end_color="63BE7B",
            ),
        )


def _apply_widths(sheet):
    """
    Use sensible widths without making the workbook enormous.
    """
    special_widths = {
        "Job ID": 18,
        "Company": 24,
        "Job Title": 32,
        "Job URL": 38,
        "Company URL": 38,
        "LinkedIn URL": 38,
        "Location": 25,
        "Posted Text": 18,
        "Experience": 20,
        "Employment Type": 20,
        "Salary": 18,
        "Resume Type": 20,
        "Resume Name": 38,
        "Resume Path": 55,
        "Matched Skills": 45,
        "Search Terms": 45,
        "Application Status": 22,
        "Referral Status": 22,
        "Referral Source": 25,
        "Referral Source Role": 28,
        "Person": 25,
        "Role": 32,
        "Activity": 28,
        "Notes": 60,
        "Message": 80,
        "Message Type": 28,
        "Details": 70,
    }

    for column_cells in sheet.columns:
        if not column_cells:
            continue

        header = str(column_cells[0].value or "").strip()
        column_letter = get_column_letter(column_cells[0].column)

        if header in special_widths:
            width = special_widths[header]
        else:
            max_length = 0
            for cell in column_cells[:80]:
                value = str(cell.value or "")
                max_length = max(max_length, len(value))

            width = min(max(max_length + 2, 10), 28)

        sheet.column_dimensions[column_letter].width = width

    # Long text is easier to read when wrapped.
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if cell.column <= sheet.max_column:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=cell.column in {
                        5, 9, 18, 26, 27, 38, 43, 45, 55
                    } or isinstance(cell.value, str) and len(cell.value) > 60,
                )


def _format_sheet(sheet):
    _set_header_style(sheet)
    _apply_widths(sheet)
    _set_hyperlinks(sheet)
    _apply_number_formats(sheet)
    _apply_status_formatting(sheet)
    _apply_score_formatting(sheet)
    _apply_table(sheet)


def _write_dataframe_to_sheet(workbook, sheet_name, df, headers):
    """
    Replace sheet contents using openpyxl while retaining workbook-level
    formatting and then reapply the project's standard formatting.
    """
    if sheet_name in workbook.sheetnames:
        old_sheet = workbook[sheet_name]
        index = workbook.index(old_sheet)
        workbook.remove(old_sheet)
        sheet = workbook.create_sheet(sheet_name, index)
    else:
        sheet = workbook.create_sheet(sheet_name)

    df = _ensure_columns(df, headers)

    sheet.append(headers)

    for values in df[headers].itertuples(index=False, name=None):
        sheet.append(list(values))

    return sheet


def _status_count(df, column, value):
    if df.empty or column not in df.columns:
        return 0

    return int(
        (
            df[column]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
            == value.lower()
        ).sum()
    )


def _contains_count(df, column, value):
    if df.empty or column not in df.columns:
        return 0

    return int(
        df[column]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.contains(value.lower(), regex=False)
        .sum()
    )


def _dashboard_metrics():
    jobs = _read_sheet_raw("Jobs")
    contacts = _read_sheet_raw("Contacts")
    referrals = _read_sheet_raw("Referrals")
    messages = _read_sheet_raw("Messages")

    relevant_jobs = 0
    if not jobs.empty and "Relevance Score" in jobs.columns:
        scores = pd.to_numeric(
            jobs["Relevance Score"], errors="coerce"
        ).fillna(0)
        relevant_jobs = int((scores >= 60).sum())

    metrics = [
        ("Last Updated", _now()),
        ("Jobs Tracked", len(jobs)),
        ("Relevant Jobs", relevant_jobs),
        ("Jobs Applied", _status_count(jobs, "Application Status", "Applied")),
        ("Jobs in Interview", _contains_count(jobs, "Application Status", "interview")),
        ("Offers", _status_count(jobs, "Application Status", "Offer")),
        ("Jobs Rejected", _status_count(jobs, "Application Status", "Rejected")),
        ("Referrals Requested", _contains_count(jobs, "Referral Status", "requested")),
        ("Referrals Received", _status_count(jobs, "Referral Status", "Received")),
        ("Referral Contacts", len(contacts)),
        ("Willing to Refer", _status_count(contacts, "Response Status", "Willing to Refer")),
        ("Referral Responses Received", _contains_count(contacts, "Response Status", "referral")),
        ("Connection Requests", _contains_count(contacts, "Request Status", "connection requested")),
        ("No Response Contacts", _status_count(contacts, "Response Status", "No Response")),
        ("Messages Tracked", len(messages)),
        ("Messages Sent", _status_count(messages, "Status", "Sent")),
        ("Referrals Recorded", len(referrals)),
    ]

    return metrics


def _create_dashboard(workbook):
    if "Dashboard" in workbook.sheetnames:
        sheet = workbook["Dashboard"]
        workbook.remove(sheet)

    dashboard = workbook.create_sheet("Dashboard", 0)

    dashboard["A1"] = "LinkedIn Job & Referral Tracker"
    dashboard["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    dashboard["A1"].fill = HEADER_FILL
    dashboard["A1"].alignment = Alignment(horizontal="center")

    dashboard.merge_cells("A1:B1")
    dashboard.row_dimensions[1].height = 32

    dashboard["A3"] = "Metric"
    dashboard["B3"] = "Value"

    for cell in dashboard[3]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")

    metrics = _dashboard_metrics()

    for row_number, (metric, value) in enumerate(metrics, start=4):
        dashboard.cell(row=row_number, column=1, value=metric)
        dashboard.cell(row=row_number, column=2, value=value)

        dashboard.cell(row=row_number, column=1).border = THIN_BORDER
        dashboard.cell(row=row_number, column=2).border = THIN_BORDER

    dashboard["D3"] = "Workflow"
    dashboard["D3"].fill = HEADER_FILL
    dashboard["D3"].font = HEADER_FONT

    workflow = [
        ("1", "Search configured LinkedIn job terms"),
        ("2", "Filter by experience, age and company size"),
        ("3", "Score against AI/Data Science and Full Stack resumes"),
        ("4", "Find relevant people and prioritize referral contacts"),
        ("5", "Send/prepare connection or referral outreach"),
        ("6", "Inspect accessible LinkedIn conversations"),
        ("7", "Classify responses and update referral state"),
        ("8", "Stop referral requests after a referral is received"),
    ]

    for row_number, (number, text) in enumerate(workflow, start=4):
        dashboard.cell(row=row_number, column=4, value=number)
        dashboard.cell(row=row_number, column=5, value=text)
        dashboard.cell(row=row_number, column=4).alignment = Alignment(
            horizontal="center"
        )
        dashboard.cell(row=row_number, column=5).alignment = Alignment(
            wrap_text=True
        )

    dashboard.column_dimensions["A"].width = 32
    dashboard.column_dimensions["B"].width = 18
    dashboard.column_dimensions["C"].width = 4
    dashboard.column_dimensions["D"].width = 10
    dashboard.column_dimensions["E"].width = 72

    dashboard.freeze_panes = "A4"


def _rebuild_dashboard_and_format(workbook):
    """
    Format all operational sheets and rebuild the dashboard from current data.
    """
    for sheet_name, headers in _sheet_definitions().items():
        sheet = _create_sheet_if_missing(
            workbook, sheet_name, headers
        )
        _format_sheet(sheet)

    _create_dashboard(workbook)

    # Dashboard should remain the first visible sheet.
    for sheet in workbook.worksheets:
        sheet.sheet_state = "visible"

    workbook.active = 0


# ============================================================================
# Tracker creation and raw I/O
# ============================================================================

def create_tracker(path=TRACKER_PATH):
    """
    Create or repair the tracker workbook.

    IMPORTANT: this function does NOT rewrite an existing healthy workbook.
    Reads must never mutate the live tracker. A save happens only when the
    workbook is new, unreadable/corrupt, or missing required structure.
    Normal workflow writes use _write_sheet(), which performs an atomic save.
    """
    _ensure_data_dir()

    path = Path(path)
    existed = path.exists()
    workbook = _load_or_create_workbook(path)
    needs_save = not existed

    # Ensure every operational sheet exists.
    for sheet_name, headers in _sheet_definitions().items():
        if sheet_name not in workbook.sheetnames:
            _create_sheet_if_missing(workbook, sheet_name, headers)
            needs_save = True

    # If this is a brand-new workbook, remove only the default Sheet if it
    # contains no useful project data.
    if "Sheet" in workbook.sheetnames:
        default_sheet = workbook["Sheet"]
        if default_sheet.max_row <= 1 and default_sheet.max_column <= 1:
            if default_sheet["A1"].value is None:
                workbook.remove(default_sheet)
                needs_save = True

    # Never allow a workbook with no visible worksheets.
    if not workbook.worksheets:
        workbook.create_sheet("Jobs")
        needs_save = True

    for sheet in workbook.worksheets:
        if sheet.sheet_state != "visible":
            sheet.sheet_state = "visible"
            needs_save = True

    # Dashboard is required. Do not rebuild a healthy dashboard merely because
    # somebody is reading a sheet; rebuilding it causes an unnecessary save.
    if "Dashboard" not in workbook.sheetnames:
        _rebuild_dashboard_and_format(workbook)
        needs_save = True

    if needs_save:
        _rebuild_dashboard_and_format(workbook)
        _atomic_save_workbook(workbook, path)

    return path


def _read_sheet_raw(sheet_name):
    """
    Read one sheet without recursively calling create_tracker().
    """
    _ensure_data_dir()

    if not Path(TRACKER_PATH).exists():
        return pd.DataFrame()

    try:
        return pd.read_excel(
            TRACKER_PATH,
            sheet_name=sheet_name,
        )
    except Exception:
        return pd.DataFrame()


def _read_sheet(sheet_name):
    """Read a sheet without rewriting the tracker.

    Only create the tracker when the file does not exist. A healthy existing
    workbook is read directly so get/read operations can never overwrite it.
    """
    _ensure_data_dir()
    if not Path(TRACKER_PATH).exists():
        create_tracker()
    return _read_sheet_raw(sheet_name)


def _write_sheet(df, sheet_name):
    """
    Replace one operational sheet and then repair/format the complete workbook.
    """
    create_tracker()

    path = Path(TRACKER_PATH)
    workbook = load_workbook(path)

    headers = _sheet_definitions().get(sheet_name)
    if headers is None:
        raise ValueError(f"Unknown tracker sheet: {sheet_name}")

    _write_dataframe_to_sheet(
        workbook,
        sheet_name,
        df,
        headers,
    )

    _rebuild_dashboard_and_format(workbook)
    _atomic_save_workbook(workbook, path)


def _ensure_columns(df, headers):
    if df is None or df.empty:
        return pd.DataFrame(columns=headers)

    df = df.copy()

    for column in headers:
        if column not in df.columns:
            df[column] = ""

    df = df[headers].copy()

    # Tracker columns can contain text, numbers, blanks and timestamps.
    # Pandas may infer an all-blank Excel column as float64; converting every
    # operational column to object prevents later string/date assignments from
    # failing with "Invalid value ... for dtype float64".
    for column in headers:
        df[column] = df[column].astype(object)

    # Job IDs are identifiers, never numeric values. Normalize Excel values
    # such as 4448345691.0 back to 4448345691 so lookups remain stable.
    if "Job ID" in df.columns:
        def _normalize_job_id(value):
            if pd.isna(value):
                return ""
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            text = str(value).strip()
            if text.endswith(".0") and text[:-2].isdigit():
                return text[:-2]
            return text

        df["Job ID"] = df["Job ID"].map(_normalize_job_id)

    return df


# ============================================================================
# Jobs
# ============================================================================

def upsert_job(job):
    df = _read_sheet("Jobs")

    job_id = str(job.get("Job ID", "")).strip()
    if not job_id:
        return

    if df.empty:
        df = pd.DataFrame(columns=JOB_HEADERS)

    df = _ensure_columns(df, JOB_HEADERS)

    # LinkedIn Job IDs are identifiers, not numeric values. Pandas may infer
    # the existing Excel column as int64, which then rejects a string Job ID
    # such as "4448345691" during an update. Normalize the column to string
    # before any comparison or assignment.
    df["Job ID"] = df["Job ID"].fillna("").astype(str).str.strip()

    row = {
        header: job.get(header, "")
        for header in JOB_HEADERS
    }
    row["Job ID"] = job_id

    now = _now()

    existing = df[
        df["Job ID"].astype(str).str.strip() == job_id
    ]

    if existing.empty:
        row["Created"] = now
        row["Updated"] = now
        df.loc[len(df)] = row
    else:
        index = existing.index[-1]

        if not row.get("Created"):
            row["Created"] = df.loc[index, "Created"]

        # Preserve application/referral progress if the new discovery row
        # does not contain a meaningful value.
        protected = {
            "Application Status",
            "Referral Status",
            "Referral Source",
            "Referral Source Role",
            "Referral Date",
        }

        row["Updated"] = now

        for column in JOB_HEADERS:
            value = row.get(column, "")

            if column in protected and (
                value is None or str(value).strip() == ""
            ):
                continue

            if value != "":
                df.loc[index, column] = value

        df.loc[index, "Updated"] = now

    _write_sheet(df, "Jobs")


def get_job_state(job_id):
    df = _read_sheet("Jobs")

    if df.empty:
        return {}

    rows = df[
        df["Job ID"].astype(str).str.strip()
        == str(job_id).strip()
    ]

    if rows.empty:
        return {}

    return rows.iloc[-1].to_dict()


# ============================================================================
# Contacts
# ============================================================================

def upsert_contact(contact):
    """
    Insert/update a contact using either Excel-style keys
    ("Job ID", "Person", ...) or people.py snake_case keys
    ("job_id", "person", ...).
    """
    df = _read_sheet("Contacts")

    if df.empty:
        df = pd.DataFrame(columns=CONTACT_HEADERS)

    df = _ensure_columns(df, CONTACT_HEADERS)
    contact = dict(contact or {})

    def value_for(header):
        if header in contact:
            return contact.get(header, "")

        snake = header.lower().replace(" ", "_")
        return contact.get(snake, "")

    job_id = str(value_for("Job ID") or "").strip()
    person = str(value_for("Person") or "").strip()

    if not job_id or not person:
        return

    existing = df[
        (df["Job ID"].astype(str).str.strip() == job_id)
        & (
            df["Person"].astype(str).str.strip().str.lower()
            == person.lower()
        )
    ]

    row = {
        header: value_for(header)
        for header in CONTACT_HEADERS
    }

    now = _now()

    if existing.empty:
        row["Created"] = now
        row["Updated"] = now
        df.loc[len(df)] = row
    else:
        index = existing.index[-1]

        if not row.get("Created"):
            row["Created"] = df.loc[index, "Created"]

        row["Updated"] = now

        # Never erase a workflow state with a discovery default.
        protected_defaults = {
            "Request Status": {"Not Requested", ""},
            "Response Status": {"No Response", ""},
            "Referral Status": {"Not Requested", ""},
            "Contact Status": {"Discovered", ""},
        }

        for column in CONTACT_HEADERS:
            value = row.get(column, "")

            if column in protected_defaults:
                current = str(df.loc[index, column] or "").strip()
                if current and value in protected_defaults[column]:
                    continue

            if value != "":
                df.loc[index, column] = value

    _write_sheet(df, "Contacts")


def get_contacts_for_job(job_id):
    df = _read_sheet("Contacts")

    if df.empty:
        return []

    rows = df[
        df["Job ID"].astype(str).str.strip()
        == str(job_id).strip()
    ]

    return rows.to_dict("records")


def get_contact(job_id, person):
    df = _read_sheet("Contacts")

    if df.empty:
        return {}

    rows = df[
        (df["Job ID"].astype(str).str.strip() == str(job_id).strip())
        & (
            df["Person"].astype(str).str.strip().str.lower()
            == str(person).strip().lower()
        )
    ]

    if rows.empty:
        return {}

    return rows.iloc[-1].to_dict()


def update_contact_response(
    job_id,
    person,
    response_status,
    response_text="",
):
    df = _read_sheet("Contacts")

    if df.empty:
        return

    df = _ensure_columns(df, CONTACT_HEADERS)

    rows = df[
        (df["Job ID"].astype(str).str.strip() == str(job_id).strip())
        & (
            df["Person"].astype(str).str.strip().str.lower()
            == str(person).strip().lower()
        )
    ]

    if rows.empty:
        return

    index = rows.index[-1]

    df.loc[index, "Response Status"] = response_status
    df.loc[index, "Last Response"] = _now()

    if response_text:
        old_notes = str(df.loc[index, "Notes"] or "").strip()
        new_note = f"LinkedIn response: {response_text}"

        if old_notes:
            df.loc[index, "Notes"] = (
                old_notes + " | " + new_note
            )
        else:
            df.loc[index, "Notes"] = new_note

    if response_status == "Willing to Refer":
        df.loc[index, "Referral Status"] = "Offered"

    elif response_status == "Referral Received":
        df.loc[index, "Referral Status"] = "Received"

    elif response_status == "Declined":
        df.loc[index, "Referral Status"] = "Declined"

    elif response_status == "Alternative Process":
        df.loc[index, "Referral Status"] = "Alternative Process"

    df.loc[index, "Updated"] = _now()

    _write_sheet(df, "Contacts")


def mark_connection_accepted(job_id, person):
    df = _read_sheet("Contacts")

    if df.empty:
        return

    df = _ensure_columns(df, CONTACT_HEADERS)

    rows = df[
        (df["Job ID"].astype(str).str.strip() == str(job_id).strip())
        & (
            df["Person"].astype(str).str.strip().str.lower()
            == str(person).strip().lower()
        )
    ]

    if rows.empty:
        return

    index = rows.index[-1]

    df.loc[index, "Connection Degree"] = "1st"
    df.loc[index, "Request Status"] = "Accepted"
    df.loc[index, "Contact Status"] = "Connected"
    df.loc[index, "Updated"] = _now()

    _write_sheet(df, "Contacts")


def mark_referral_requested(job_id, person):
    df = _read_sheet("Contacts")

    if df.empty:
        return

    df = _ensure_columns(df, CONTACT_HEADERS)

    rows = df[
        (df["Job ID"].astype(str).str.strip() == str(job_id).strip())
        & (
            df["Person"].astype(str).str.strip().str.lower()
            == str(person).strip().lower()
        )
    ]

    if rows.empty:
        return

    index = rows.index[-1]

    df.loc[index, "Request Status"] = "Referral Requested"
    df.loc[index, "Contact Status"] = "Contacted"
    df.loc[index, "Referral Status"] = "Requested"
    df.loc[index, "Last Contacted"] = _now()
    df.loc[index, "Updated"] = _now()

    _write_sheet(df, "Contacts")


def mark_message_sent(
    job_id,
    person,
    message_type=None,
):
    df = _read_sheet("Contacts")

    if df.empty:
        return

    df = _ensure_columns(df, CONTACT_HEADERS)

    rows = df[
        (df["Job ID"].astype(str).str.strip() == str(job_id).strip())
        & (
            df["Person"].astype(str).str.strip().str.lower()
            == str(person).strip().lower()
        )
    ]

    if rows.empty:
        return

    index = rows.index[-1]

    df.loc[index, "Contact Status"] = "Contacted"
    df.loc[index, "Last Contacted"] = _now()

    if message_type:
        message_type_lower = str(message_type).lower()

        if "connection" in message_type_lower:
            df.loc[index, "Request Status"] = "Connection Requested"

        elif "referral" in message_type_lower:
            df.loc[index, "Request Status"] = "Referral Requested"
            df.loc[index, "Referral Status"] = "Requested"

        old_notes = str(df.loc[index, "Notes"] or "").strip()
        note = f"Message sent: {message_type}."

        df.loc[index, "Notes"] = (
            f"{old_notes} {note}".strip()
        )

    df.loc[index, "Updated"] = _now()

    _write_sheet(df, "Contacts")

    # Synchronize the corresponding message record.
    messages = _read_sheet("Messages")

    if not messages.empty:
        messages = _ensure_columns(
            messages,
            MESSAGE_HEADERS,
        )

        message_rows = messages[
            (messages["Job ID"].astype(str).str.strip()
             == str(job_id).strip())
            & (
                messages["Person"].astype(str).str.strip().str.lower()
                == str(person).strip().lower()
            )
            & (
                messages["Message Type"].astype(str).str.strip().str.lower()
                == str(message_type or "").strip().lower()
            )
        ]

        if not message_rows.empty:
            message_index = message_rows.index[-1]
            messages.loc[message_index, "Status"] = "Sent"
            messages.loc[message_index, "Sent Date"] = _now()
            _write_sheet(messages, "Messages")


# ============================================================================
# Referral state
# ============================================================================

def referral_already_received(job_id):
    df = _read_sheet("Referrals")

    if df.empty:
        return False

    rows = df[
        (df["Job ID"].astype(str).str.strip() == str(job_id).strip())
        & (
            df["Referral Received"]
            .astype(str)
            .str.lower()
            .isin(["yes", "true", "1", "received"])
        )
    ]

    return not rows.empty


def referral_received(
    job_id,
    person,
    role="",
    company="",
    notes="",
):
    referrals = _read_sheet("Referrals")

    if referrals.empty:
        referrals = pd.DataFrame(columns=REFERRAL_HEADERS)

    referrals = _ensure_columns(
        referrals,
        REFERRAL_HEADERS,
    )

    now = _now()

    # Job-specific: only one referral source is recorded for a job.
    existing = referrals[
        (referrals["Job ID"].astype(str).str.strip()
         == str(job_id).strip())
        & (
            referrals["Referral Received"]
            .astype(str)
            .str.lower()
            .isin(["yes", "true", "1", "received"])
        )
    ]

    if not existing.empty:
        return

    row = {
        "Job ID": job_id,
        "Company": company,
        "Person": person,
        "Role": role,
        "Referral Requested": "Yes",
        "Referral Offered": "Yes",
        "Referral Received": "Yes",
        "Referral Date": now,
        "Is Referral Source": "Yes",
        "Status": "Received",
        "Notes": notes,
        "Created": now,
        "Updated": now,
    }

    referrals.loc[len(referrals)] = row

    _write_sheet(
        referrals,
        "Referrals",
    )

    # Update the corresponding Jobs row.
    jobs = _read_sheet("Jobs")

    if not jobs.empty:
        jobs = _ensure_columns(
            jobs,
            JOB_HEADERS,
        )

        rows = jobs[
            jobs["Job ID"].astype(str).str.strip()
            == str(job_id).strip()
        ]

        if not rows.empty:
            index = rows.index[-1]

            jobs.loc[index, "Referral Status"] = "Received"
            jobs.loc[index, "Referral Source"] = person
            jobs.loc[index, "Referral Source Role"] = role
            jobs.loc[index, "Referral Date"] = now
            jobs.loc[index, "Updated"] = now

            _write_sheet(
                jobs,
                "Jobs",
            )

    # Also update the source contact, when available.
    contacts = _read_sheet("Contacts")

    if not contacts.empty:
        contacts = _ensure_columns(
            contacts,
            CONTACT_HEADERS,
        )

        contact_rows = contacts[
            (contacts["Job ID"].astype(str).str.strip()
             == str(job_id).strip())
            & (
                contacts["Person"].astype(str).str.strip().str.lower()
                == str(person).strip().lower()
            )
        ]

        if not contact_rows.empty:
            index = contact_rows.index[-1]
            contacts.loc[index, "Response Status"] = "Referral Received"
            contacts.loc[index, "Referral Status"] = "Received"
            contacts.loc[index, "Updated"] = now

            _write_sheet(
                contacts,
                "Contacts",
            )


def get_willing_contacts_except(job_id, excluded_person=""):
    df = _read_sheet("Contacts")

    if df.empty:
        return []

    rows = df[
        (df["Job ID"].astype(str).str.strip() == str(job_id).strip())
        & (
            df["Response Status"]
            .astype(str)
            .str.lower()
            .isin(
                [
                    "willing to refer",
                    "referral offered",
                    "referral given",
                ]
            )
        )
    ]

    if excluded_person:
        rows = rows[
            rows["Person"].astype(str).str.strip().str.lower()
            != str(excluded_person).strip().lower()
        ]

    return rows.to_dict("records")


# ============================================================================
# Messages
# ============================================================================

def message_exists(
    job_id,
    person,
    message_type,
):
    df = _read_sheet("Messages")

    if df.empty:
        return False

    rows = df[
        (df["Job ID"].astype(str).str.strip()
         == str(job_id).strip())
        & (
            df["Person"].astype(str).str.strip().str.lower()
            == str(person).strip().lower()
        )
        & (
            df["Message Type"].astype(str).str.strip().str.lower()
            == str(message_type).strip().lower()
        )
    ]

    return not rows.empty


def add_message(
    job_id,
    company,
    person,
    message_type,
    message,
    resume_type="",
    resume_name="",
    resume_path="",
    job_url="",
    attachment_required="No",
    status="Prepared",
):
    if message_exists(
        job_id,
        person,
        message_type,
    ):
        return

    df = _read_sheet("Messages")

    if df.empty:
        df = pd.DataFrame(columns=MESSAGE_HEADERS)

    df = _ensure_columns(
        df,
        MESSAGE_HEADERS,
    )

    row = {
        "Job ID": job_id,
        "Company": company,
        "Person": person,
        "Message Type": message_type,
        "Message": message,
        "Resume Type": resume_type,
        "Resume Name": resume_name,
        "Resume Path": resume_path,
        "Job URL": job_url,
        "Attachment Required": attachment_required,
        "Status": status,
        "Created Date": _now(),
        "Sent Date": "",
    }

    df.loc[len(df)] = row

    _write_sheet(
        df,
        "Messages",
    )


# ============================================================================
# Activity
# ============================================================================

def add_activity(
    job_id,
    company,
    person,
    activity,
    details="",
):
    df = _read_sheet("Activity")

    if df.empty:
        df = pd.DataFrame(columns=ACTIVITY_HEADERS)

    df = _ensure_columns(
        df,
        ACTIVITY_HEADERS,
    )

    df.loc[len(df)] = {
        "Job ID": job_id,
        "Company": company,
        "Person": person,
        "Activity": activity,
        "Details": details,
        "Date": _now(),
    }

    _write_sheet(
        df,
        "Activity",
    )
