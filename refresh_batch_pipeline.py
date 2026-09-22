#!/usr/bin/env python3
"""
refresh_batch_pipeline.py

Rebuilds Batch_Start_to_Picked_date.xlsx end-to-end, twice a day, with no manual CSV
upload step:

  1. Batch Enrollment (Metabase) tab  <- Metabase REST API (native SQL question)
  2. Picked Students (raw) tab        <- Google Sheets API, pulling ONLY the columns
                                          needed from the FlyWheel tracker (UserID,
                                          Batch, Status, Picked Date) instead of the
                                          whole 82-column sheet
  3. Batch Start to Picked date tab   <- rebuilt with the same INDEX/MATCH,
                                          "-" placeholder, and % of Currently Enrolled
                                          logic as the last manual build.

Run twice daily from cron / GitHub Actions (see refresh_batch_pipeline.yml).

Only two secrets are needed at runtime — everything else about *where* the data lives
is hardcoded below in the CONFIG block, since it doesn't change run to run:

  METABASE_API_KEY             Metabase API key (Admin > Settings > Authentication > API Keys)
  GOOGLE_SERVICE_ACCOUNT_JSON   The service account's JSON key, either as a file path OR as the
                                 raw JSON text itself (both are accepted — see load below).
                                 Share the FlyWheel sheet with that service account's
                                 client_email as Viewer.

pip install requests google-api-python-client google-auth openpyxl
"""
import csv
import datetime as dt
import json
import os
import sys
import tempfile

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter, column_index_from_string

# ----------------------------------------------------------------------------
# CONFIG — fixed facts about this pipeline. Edit here, not via env vars, if any
# of these ever change (new Metabase host, new saved question, sheet moved, etc).
# ----------------------------------------------------------------------------
METABASE_URL = "https://metabase-lierhfgoeiwhr.newtonschool.co"
METABASE_DATABASE_ID = 4  # Newton School Postgres
METABASE_QUESTION_ID = 12957  # saved question "Batch Status" — re-run via /api/card/12957/query

FLYWHEEL_SHEET_ID = "1Ue49enEEpgNaOEdQVgwgsehWvHb3HEI0Q-qekvAzYyU"
# Tab name inside that spreadsheet — confirmed from the tab bar screenshot (gid=1147350782).
FLYWHEEL_TAB_NAME = "Prog<>Placement"

OUTPUT_PATH = "Batch_Start_to_Picked_date.xlsx"

BATCH_TITLE_FILTER = "Professional Certificate Course In Data Science%"

# ----------------------------------------------------------------------------
# Secrets — the only two things read from the environment.
# ----------------------------------------------------------------------------
METABASE_API_KEY = os.environ["METABASE_API_KEY"]

_raw_sa = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
_stripped = _raw_sa.strip()
if _stripped.startswith("{"):
    # secret holds the raw JSON key text -> write it to a temp file for google-auth
    _sa_fh = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
    _sa_fh.write(_stripped)
    _sa_fh.close()
    GOOGLE_SERVICE_ACCOUNT_JSON = _sa_fh.name
else:
    # secret holds a file path instead (e.g. a path checked out by an earlier CI step)
    GOOGLE_SERVICE_ACCOUNT_JSON = _raw_sa

ENROLLMENT_SQL = f"""
select
  c.id as course_id,
  c.title as batch_name,
  c.start_timestamp as batch_start_date,
  count(distinct cum.user_id) filter (
    where cum.status in (8,11,12,30)
  ) as initially_enrolled,
  count(distinct cum.user_id) filter (
    where cum.status = 8
    and not exists (
      select 1
      from courses_courseuserlabelmapping cclm
      join technologies_label tl on tl.id = cclm.label_id
      where cclm.course_user_mapping_id = cum.id
        and tl.name = 'NBFC - Refund Requested'
    )
  ) as currently_enrolled,
  count(distinct cum.user_id) filter (
    where cum.status = 8
    and exists (
      select 1
      from courses_courseuserlabelmapping cclm
      join technologies_label tl on tl.id = cclm.label_id
      where cclm.course_user_mapping_id = cum.id
        and tl.name = 'NBFC - Refund Requested'
    )
  ) as refund_requested,
  count(distinct cum.user_id) filter (where cum.status in (11,12)) as course_cancellation,
  count(distinct cum.user_id) filter (where cum.status = 30) as deferred
from courses_course c
left join courses_courseusermapping cum
  on cum.course_id = c.id
where c.title ilike '{BATCH_TITLE_FILTER}'
group by c.id, c.title, c.start_timestamp
order by c.start_timestamp;
""".strip()


# ----------------------------------------------------------------------------
# 1. Metabase — Batch Enrollment
# ----------------------------------------------------------------------------
def fetch_enrollment_rows():
    headers = {"X-API-Key": METABASE_API_KEY, "Content-Type": "application/json"}

    if METABASE_QUESTION_ID:
        url = f"{METABASE_URL}/api/card/{METABASE_QUESTION_ID}/query"
        resp = requests.post(url, headers=headers, timeout=60)
    else:
        url = f"{METABASE_URL}/api/dataset"
        payload = {
            "type": "native",
            "native": {"query": ENROLLMENT_SQL},
            "database": METABASE_DATABASE_ID,
        }
        resp = requests.post(url, headers=headers, json=payload, timeout=60)

    resp.raise_for_status()
    data = resp.json()["data"]
    cols = [c["name"] for c in data["cols"]]
    rows = [dict(zip(cols, r)) for r in data["rows"]]
    if not rows:
        raise RuntimeError("Metabase returned 0 enrollment rows — check the query/credentials.")
    return rows


# ----------------------------------------------------------------------------
# 2. Google Sheets — Picked Students (raw), pulling only the needed columns
# ----------------------------------------------------------------------------
def fetch_picked_rows(target_batches):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds = service_account.Credentials.from_service_account_file(
        GOOGLE_SERVICE_ACCOUNT_JSON,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    service = build("sheets", "v4", credentials=creds)

    # Pull only the 4 columns we actually use instead of the whole 82-column sheet:
    # A=UserID, D=Batch, U=Status, V=Picked Date  (adjust letters if the FlyWheel
    # sheet's column order changes — these match cols 0,3,20,21 in the CSV export
    # this pipeline was originally built from).
    ranges = [
        f"'{FLYWHEEL_TAB_NAME}'!A2:A",
        f"'{FLYWHEEL_TAB_NAME}'!D2:D",
        f"'{FLYWHEEL_TAB_NAME}'!U2:U",
        f"'{FLYWHEEL_TAB_NAME}'!V2:V",
    ]
    result = service.spreadsheets().values().batchGet(
        spreadsheetId=FLYWHEEL_SHEET_ID, ranges=ranges
    ).execute()
    value_ranges = result["valueRanges"]
    user_ids = [r[0] if r else "" for r in value_ranges[0].get("values", [])]
    batches = [r[0] if r else "" for r in value_ranges[1].get("values", [])]
    statuses = [r[0] if r else "" for r in value_ranges[2].get("values", [])]
    picked_dates = [r[0] if r else "" for r in value_ranges[3].get("values", [])]

    n = max(len(user_ids), len(batches), len(statuses), len(picked_dates))

    def get(lst, i):
        return lst[i] if i < len(lst) else ""

    out = []
    for i in range(n):
        batch = get(batches, i)
        status = get(statuses, i)
        if batch not in target_batches or status != "Picked":
            continue
        uid = get(user_ids, i)
        pdate_raw = get(picked_dates, i).strip()
        pdate = ""
        if pdate_raw and pdate_raw not in ("NA", "#N/A", "#REF!"):
            for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
                try:
                    pdate = dt.datetime.strptime(pdate_raw, fmt).strftime("%Y-%m-%d")
                    break
                except ValueError:
                    continue
        out.append((uid, batch, pdate))
    return out


# ----------------------------------------------------------------------------
# 3. Build the workbook (same layout/logic as the manual build)
# ----------------------------------------------------------------------------
def build_workbook(enroll_rows, picked_rows, output_path):
    NAVY, WHITE, GREY = "FF1F3864", "FFFFFFFF", "FF333333"
    header_font = Font(name="Arial", bold=True, color=WHITE)
    header_fill = PatternFill("solid", fgColor=NAVY, bgColor=GREY)
    data_font = Font(name="Arial")
    note_font = Font(name="Arial", italic=True, size=9, color="FF808080")
    manual_fill = PatternFill("solid", fgColor="FFFFF2CC")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- Batch Enrollment (Metabase) ---
    ws1 = wb.create_sheet("Batch Enrollment (Metabase)")
    headers1 = ["Batch", "Batch Start Date", "Initially Enrolled", "Currently Enrolled",
                "Refund Requested", "Course Cancellation", "Deferred"]
    for j, h in enumerate(headers1, start=1):
        c = ws1.cell(row=1, column=j, value=h)
        c.font, c.fill = header_font, header_fill
    for i, row in enumerate(enroll_rows, start=2):
        start_date = row["batch_start_date"]
        if isinstance(start_date, str):
            start_date = dt.datetime.fromisoformat(start_date.replace("Z", "+00:00")).replace(tzinfo=None)
        vals = [row["batch_name"], start_date, int(row["initially_enrolled"]),
                int(row["currently_enrolled"]), int(row["refund_requested"]),
                int(row["course_cancellation"]), int(row["deferred"])]
        for j, v in enumerate(vals, start=1):
            c = ws1.cell(row=i, column=j, value=v)
            c.font = data_font
            if j == 2:
                c.number_format = "dd-mmm-yyyy"
    last_enroll_row = len(enroll_rows) + 1
    ws1.cell(row=1, column=9,
              value=f"Source: Metabase (Newton School DB). Refreshed: {dt.date.today().isoformat()}.").font = note_font
    ws1.column_dimensions["A"].width = 55
    for col in "BCDEFG":
        ws1.column_dimensions[col].width = 16
    ws1.freeze_panes = "A2"

    # --- Picked Students (raw) ---
    ws2 = wb.create_sheet("Picked Students (raw)")
    for j, h in enumerate(["UserID", "Batch", "Picked Date"], start=1):
        c = ws2.cell(row=1, column=j, value=h)
        c.font, c.fill = header_font, header_fill
    for i, (uid, batch, pdate_str) in enumerate(picked_rows, start=2):
        uid_val = int(uid) if str(uid).strip().isdigit() else uid
        pdate = dt.datetime.strptime(pdate_str, "%Y-%m-%d") if pdate_str else None
        ws2.cell(row=i, column=1, value=uid_val).font = data_font
        ws2.cell(row=i, column=2, value=batch).font = data_font
        c3 = ws2.cell(row=i, column=3, value=pdate)
        c3.font = data_font
        if pdate:
            c3.number_format = "dd-mmm-yyyy"
    last_picked_row = len(picked_rows) + 1
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 60
    ws2.column_dimensions["C"].width = 14
    ws2.freeze_panes = "A2"
    ws2.cell(row=1, column=5,
              value=f"Source: FlyWheel Google Sheet via Sheets API. Refreshed: {dt.date.today().isoformat()}.").font = note_font

    # --- Batch Start to Picked date ---
    ws3 = wb.create_sheet("Batch Start to Picked date")
    NWINDOWS = 18
    headers3 = ["Batch", "Batch (A/B)", "Batch Start Date", "Initially Enrolled",
                "Currently Enrolled", "Picked Students"]
    for w in range(NWINDOWS):
        days = (w + 1) * 30
        headers3 += [f"Within {days} Days", "% of Currently Enrolled"]
    headers3.append("Not Placed yet")
    for j, h in enumerate(headers3, start=1):
        c = ws3.cell(row=1, column=j, value=h)
        c.font, c.fill = header_font, header_fill

    N_BATCH_ROWS = max(30, len(enroll_rows) + 9)
    first_data_row, last_data_row = 2, 2 + N_BATCH_ROWS - 1
    known_batches = [row["batch_name"] for row in enroll_rows]

    enroll_range_name = f"'Batch Enrollment (Metabase)'!$A$2:$A${last_enroll_row}"
    enroll_col = lambda letter: f"'Batch Enrollment (Metabase)'!${letter}$2:${letter}${last_enroll_row}"
    picked_batch_range = f"'Picked Students (raw)'!$B$2:$B${last_picked_row}"
    picked_date_range = f"'Picked Students (raw)'!$C$2:$C${last_picked_row}"

    for r in range(first_data_row, last_data_row + 1):
        idx = r - first_data_row
        a_cell = ws3.cell(row=r, column=1)
        if idx < len(known_batches):
            a_cell.value = known_batches[idx]
        a_cell.font = data_font

        ab_cell = ws3.cell(row=r, column=2)
        ab_cell.font = data_font
        ab_cell.fill = manual_fill
        ab_cell.alignment = Alignment(horizontal="center")

        Ar = f"$A{r}"
        b = ws3.cell(row=r, column=3,
                      value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("B")},MATCH({Ar},{enroll_range_name},0)),""))')
        b.font = data_font
        b.number_format = "dd-mmm-yyyy"
        ws3.cell(row=r, column=4,
                  value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("C")},MATCH({Ar},{enroll_range_name},0)),""))').font = data_font
        ws3.cell(row=r, column=5,
                  value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("D")},MATCH({Ar},{enroll_range_name},0)),""))').font = data_font
        ws3.cell(row=r, column=6,
                  value=f'=IF({Ar}="","",COUNTIF({picked_batch_range},{Ar}))').font = data_font

        Dr, Br, Er = f"$E{r}", f"$C{r}", f"$F{r}"
        for w in range(NWINDOWS):
            days = (w + 1) * 30
            within_col = 6 + 1 + w * 2
            pct_col = within_col + 1
            within_letter = get_column_letter(within_col)
            within_formula = (
                f'=IF({Ar}="","",IF({Br}="","",IF(TODAY()<{Br}+{days},"-",'
                f"SUMPRODUCT(({picked_batch_range}={Ar})*({picked_date_range}<>\"\")*"
                f'(({picked_date_range}-{Br})<{days})))))'
            )
            ws3.cell(row=r, column=within_col, value=within_formula).font = data_font
            pct_formula = (
                f'=IF({Ar}="","",IF({within_letter}{r}="-","-",'
                f'IF({Dr}=0,0,{within_letter}{r}/{Dr})))'
            )
            pcell = ws3.cell(row=r, column=pct_col, value=pct_formula)
            pcell.font = data_font
            pcell.number_format = "0%"

        last_within_letter = get_column_letter(6 + 1 + (NWINDOWS - 1) * 2)
        ws3.cell(row=r, column=len(headers3),
                  value=f'=IF({Ar}="","",IF({last_within_letter}{r}="-","-",{Er}-{last_within_letter}{r}))').font = data_font

    total_row = last_data_row + 1
    ws3.cell(row=total_row, column=1, value="TOTAL").font = Font(name="Arial", bold=True)
    for letter in ("D", "E", "F"):
        ws3.cell(row=total_row, column=column_index_from_string(letter),
                  value=f"=SUM({letter}{first_data_row}:{letter}{last_data_row})").font = Font(name="Arial", bold=True)
    for w in range(NWINDOWS):
        within_col = 6 + 1 + w * 2
        pct_col = within_col + 1
        wl, pl = get_column_letter(within_col), get_column_letter(pct_col)
        ws3.cell(row=total_row, column=within_col,
                  value=f"=SUM({wl}{first_data_row}:{wl}{last_data_row})").font = Font(name="Arial", bold=True)
        pcell = ws3.cell(row=total_row, column=pct_col,
                          value=(f'=IF(SUM($E${first_data_row}:$E${last_data_row})=0,0,'
                                 f'{wl}{total_row}/SUM($E${first_data_row}:$E${last_data_row}))'))
        pcell.font = Font(name="Arial", bold=True)
        pcell.number_format = "0%"
    last_within_letter = get_column_letter(6 + 1 + (NWINDOWS - 1) * 2)
    ws3.cell(row=total_row, column=len(headers3),
              value=f"=F{total_row}-{last_within_letter}{total_row}").font = Font(name="Arial", bold=True)

    ws3.column_dimensions["A"].width = 55
    ws3.column_dimensions["B"].width = 13
    ws3.column_dimensions["C"].width = 14
    for j in range(4, len(headers3) + 1):
        ws3.column_dimensions[get_column_letter(j)].width = 13
    ws3.freeze_panes = "D2"

    for w in range(NWINDOWS):
        pct_col = 6 + 1 + w * 2 + 1
        letter = get_column_letter(pct_col)
        rng = f"{letter}{first_data_row}:{letter}{last_data_row}"
        ws3.conditional_formatting.add(rng, ColorScaleRule(
            start_type="min", start_color="F8696B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="63BE7B"))
    np_letter = get_column_letter(len(headers3))
    ws3.conditional_formatting.add(f"{np_letter}{first_data_row}:{np_letter}{last_data_row}", ColorScaleRule(
        start_type="min", start_color="63BE7B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="F8696B"))
    ws3.conditional_formatting.add(f"F{first_data_row}:F{last_data_row}", ColorScaleRule(
        start_type="min", start_color="FFFFFFFF", end_type="max", end_color="FF9DC3E6"))

    wb.save(output_path)
    return output_path


def main():
    print("Fetching enrollment data from Metabase...", file=sys.stderr)
    enroll_rows = fetch_enrollment_rows()
    print(f"  -> {len(enroll_rows)} batches", file=sys.stderr)

    target_batches = {r["batch_name"] for r in enroll_rows}

    print("Fetching Picked rows from the FlyWheel Google Sheet...", file=sys.stderr)
    picked_rows = fetch_picked_rows(target_batches)
    print(f"  -> {len(picked_rows)} picked rows", file=sys.stderr)

    print(f"Building {OUTPUT_PATH} ...", file=sys.stderr)
    build_workbook(enroll_rows, picked_rows, OUTPUT_PATH)
    print("Done.", file=sys.stderr)


if __name__ == "__main__":
    main()
