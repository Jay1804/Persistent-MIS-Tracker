import asyncio
import os
import zipfile
from pathlib import Path

import pandas as pd
import pythoncom
import win32com.client as win32
from dateutil import parser as date_parser
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_COUNT = -4112
XL_CENTER = -4108
XL_SRC_RANGE = 1
XL_YES = 1
LIGHT_GREY = 0xD9D9D9
MAX_COLUMN_WIDTH = 60
ROW_HEIGHT = 20

load_dotenv()

LOGIN_URL = "https://mis.authbridge.com/export_query/login.php"
USERNAME = os.getenv("MIS_USERNAME")
PASSWORD = os.getenv("MIS_PASSWORD")

HOST = "Bridge Live"
DATABASE = "Bridge Live"
DATA_TIME_SLAB = "Client Specific Tracker"
QUERY = "Persistent Systems Limited Cleint Code-7084"

DATA_TIME_SLAB_2 = "Tracker - Query"
QUERY_2 = "Client Wise All Cases - Bridge"
CLIENT_CODE = "7084"

DOWNLOAD_DIR = Path(__file__).resolve().parent
REPORT_PATH = DOWNLOAD_DIR / "Persistent MIS Tracker Report.xlsx"

COLUMN_NAMES = [
    "Emp Code",
    "Name",
    "Vendor ID",
    "In Date",
    "Overall case status",
    "CBC Report Date",
    "CBC Report Color Code",
    "Final Report Date",
    "Final Report Colour code",
    "Latest_Report_Sent",
    "Latest_Report_Severity",
    "Insufficiency Raised",
    "Pending Checks",
    "Initiated Checks",
]

EXCLUDED_STATUSES = ["Duplicity", "Closed by Authbridge"]


def login(page):
    page.goto(LOGIN_URL)
    page.fill('input[name="username"]', USERNAME)
    page.fill('input[name="password"]', PASSWORD)
    page.click('input[name="login"]')
    page.wait_for_load_state("networkidle")


def select_report(page):
    page.select_option('select[name="hostname"]', HOST)
    page.select_option('select[name="database"]', DATABASE)
    page.select_option('select[name="access_time"]', DATA_TIME_SLAB)
    page.select_option('select[name="csv_query"]', QUERY)


def select_tracker_report(page):
    page.select_option('select[name="hostname"]', HOST)
    page.select_option('select[name="database"]', DATABASE)
    page.select_option('select[name="access_time"]', DATA_TIME_SLAB_2)
    page.select_option('select[name="csv_query"]', QUERY_2)
    page.fill('input[name="client1"]', CLIENT_CODE)


def export_report(page):
    with page.expect_download() as download_info:
        page.click("#run_query")
    download = download_info.value
    target_path = DOWNLOAD_DIR / download.suggested_filename
    download.save_as(target_path)
    return target_path


def unzip_report(zip_path):
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(DOWNLOAD_DIR)
        csv_name = zf.namelist()[0]
    zip_path.unlink()
    return DOWNLOAD_DIR / csv_name


def parse_doj(value):
    value = value.strip()
    if not value or value.lower() == "not mentioned":
        return None
    try:
        return date_parser.parse(value).date()
    except (ValueError, OverflowError):
        return None


def build_doj_map(tracker_csv_path):
    df = pd.read_csv(
        tracker_csv_path,
        usecols=["case_ars_no", "case_flex_field2"],
        dtype=str,
        keep_default_na=False,
    )
    doj_map = {}
    for vendor_id, raw_value in zip(df["case_ars_no"], df["case_flex_field2"]):
        vendor_id = vendor_id.strip()
        if not vendor_id or vendor_id in doj_map:
            continue
        parsed = parse_doj(raw_value)
        if parsed is not None:
            doj_map[vendor_id] = parsed
    tracker_csv_path.unlink()
    return doj_map


def normalize_emp_code(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text.isdigit():
        return int(text)
    return text


def build_report_workbook(csv_path, doj_map):
    df = pd.read_csv(csv_path, dtype={0: str})
    df = df.iloc[:, :14]
    df.columns = COLUMN_NAMES
    df["Emp Code"] = df["Emp Code"].map(normalize_emp_code)
    df["Date of joining"] = df["Vendor ID"].astype(str).str.strip().map(doj_map)
    df = df[~df["Overall case status"].isin(EXCLUDED_STATUSES)]
    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Data Tab", index=False)
    return REPORT_PATH


def add_pivot_table(pivot_cache, sheet, anchor_cell, title, field_name):
    sheet.Range(anchor_cell).Value = title
    row, col = sheet.Range(anchor_cell).Row, sheet.Range(anchor_cell).Column
    title_range = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col + 1))
    title_range.Merge()
    title_range.Font.Bold = True
    title_range.HorizontalAlignment = XL_CENTER

    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=sheet.Cells(row + 2, col),
        TableName=f"{field_name.replace(' ', '')}Pivot",
    )
    pivot_table.TableStyle2 = "PivotStyleMedium9"
    row_field = pivot_table.PivotFields(field_name)
    row_field.Orientation = XL_ROW_FIELD
    pivot_table.AddDataField(pivot_table.PivotFields("Vendor ID"), "Count", XL_COUNT)

    for item in row_field.PivotItems():
        item.Visible = item.Name not in ("(blank)", "")

    sheet.Cells(row + 2, col).Value = "Case Status"
    header_range = sheet.Range(sheet.Cells(row + 2, col), sheet.Cells(row + 2, col + 1))
    header_range.Font.Bold = True
    header_range.Interior.Color = LIGHT_GREY


def build_pivot_summary(report_path):
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(report_path))
        data_sheet = workbook.Sheets("Data Tab")
        last_row = data_sheet.Cells(data_sheet.Rows.Count, 1).End(-4162).Row
        last_col = data_sheet.Cells(1, data_sheet.Columns.Count).End(-4159).Column
        source_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(last_row, last_col))

        data_table = data_sheet.ListObjects.Add(XL_SRC_RANGE, source_range, None, XL_YES)
        data_table.Name = "DataTab"
        data_table.TableStyle = "TableStyleMedium9"
        header_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(1, last_col))
        header_range.HorizontalAlignment = XL_CENTER

        source_range.EntireColumn.AutoFit()
        for column in source_range.Columns:
            if column.ColumnWidth > MAX_COLUMN_WIDTH:
                column.ColumnWidth = MAX_COLUMN_WIDTH
                column.WrapText = True
        source_range.EntireRow.RowHeight = ROW_HEIGHT

        data_sheet.Activate()
        data_sheet.Range("A2").Select()
        excel.ActiveWindow.FreezePanes = True

        for sheet in workbook.Sheets:
            if sheet.Name == "Summary":
                sheet.Delete()
                break

        summary_sheet = workbook.Sheets.Add(Before=workbook.Sheets(1))
        summary_sheet.Name = "Summary"

        pivot_cache = workbook.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=source_range)
        add_pivot_table(pivot_cache, summary_sheet, "A1", "Overall Case wise Status", "Overall case status")
        add_pivot_table(pivot_cache, summary_sheet, "E1", "CBC  Severity", "CBC Report Color Code")

        summary_sheet.Columns("A:F").AutoFit()
        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def generate_report(headless=False):
    previous_policy = asyncio.get_event_loop_policy()
    if os.name == "nt":
        # Streamlit's Tornado server forces WindowsSelectorEventLoopPolicy process-wide,
        # but Playwright's sync API needs Proactor to launch the browser subprocess.
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    try:
        with sync_playwright() as p:
            # Playwright's bundled Chromium binary gets its outbound HTTPS blocked by
            # endpoint security software on this machine; the system-installed Edge
            # (already trusted) works fine.
            browser = p.chromium.launch(headless=headless, channel="msedge")
            page = browser.new_page()
            login(page)

            select_report(page)
            zip_path = export_report(page)

            select_tracker_report(page)
            tracker_zip_path = export_report(page)

            browser.close()
    finally:
        asyncio.set_event_loop_policy(previous_policy)

    csv_path = unzip_report(zip_path)
    tracker_csv_path = unzip_report(tracker_zip_path)
    doj_map = build_doj_map(tracker_csv_path)
    report_path = build_report_workbook(csv_path, doj_map)
    build_pivot_summary(report_path)
    return report_path


def main():
    report_path = generate_report(headless=False)
    print(f"Report saved to: {report_path}")


if __name__ == "__main__":
    main()
