#!/usr/bin/env python3
"""
refresh_batch_pipeline.py

Refreshes the "Batch Start to Picked date" tracker directly inside a Google Sheet,
twice a day, with no manual CSV upload step and no separate .xlsx file to distribute:

  1. Batch Enrollment (Metabase) tab  <- Metabase REST API (saved question / native SQL),
                                          including the GEM / Non-GEM A-B split: any batch
                                          that has real GEM(728)/Non-GEM(729) label tagging
                                          gets 2 extra rows under it ("<batch> — GEM (A)",
                                          "<batch> — Non-GEM (B)") alongside its own
                                          unsplit total row.
  2. Picked Students (raw) tab        <- Google Sheets API, reading ONLY the columns
                                          needed from the FlyWheel tracker (UserID,
                                          Batch, Type of Experience, Status, Picked
                                          Date), tagged with each user's GEM Status via
                                          a second Metabase query.
  3. Batch Start to Picked date tab   <- rebuilt with name-keyed INDEX/MATCH, "-" for
                                          windows that haven't elapsed yet, "% of
                                          Currently Enrolled", and the GEM(A)/Non-GEM(B)
                                          rows auto-inserted under their parent batch
                                          (TOTAL only sums the unsplit rows).
  4. Batch Start-Picked (Freshers)    <- same layout, Picked Students filtered to
     Batch Start-Picked (Unemployed)     FlyWheel "Type of Experience" = "Fresher" /
                                          "Career Gap/ Non Working".

All 5 tabs are written straight into the destination Google Sheet (OUTPUT_SHEET_ID
below) via the Sheets API — each run clears and rewrites those tabs from scratch, so
the sheet always reflects the latest Metabase + FlyWheel data. Formulas are written as
live Google Sheets formulas (not baked-in values), so anyone can inspect/extend them.

Run twice daily from cron / GitHub Actions (see .github/workflows/refresh_batch_pipeline.yml).

Only two secrets are needed at runtime — everything else about *where* the data lives
is hardcoded below in the CONFIG block, since it doesn't change run to run:

  METABASE_API_KEY             Metabase API key (Admin > Settings > Authentication > API Keys)
  GOOGLE_SERVICE_ACCOUNT_JSON   The service account's JSON key, either as a file path OR as the
                                 raw JSON text itself (both are accepted — see load below).
                                 This service account needs:
                                   - Viewer access to the FlyWheel sheet (reads Picked data)
                                   - Editor access to the OUTPUT sheet (writes the 5 tabs)

pip install requests google-api-python-client google-auth
"""
import datetime as dt
import os
import sys
import tempfile

import requests

# ----------------------------------------------------------------------------
# CONFIG — fixed facts about this pipeline. Edit here, not via env vars, if any
# of these ever change (new Metabase host, new saved question, sheet moved, etc).
# ----------------------------------------------------------------------------
METABASE_URL = "https://metabase-lierhfgoeiwhr.newtonschool.co"
METABASE_DATABASE_ID = 4  # Newton School Postgres
METABASE_QUESTION_ID = 12957  # saved question "Batch Status" — re-run via /api/card/12957/query
# NOTE: question 12957's SQL must be updated to the GEM-split version in enrollment_query.sql
# (adds the gem_status column) for this pipeline to pick up the A/B split. If it still has the
# old (unsplit) query, the pipeline falls back gracefully — every batch just stays unsplit.

FLYWHEEL_SHEET_ID = "1Ue49enEEpgNaOEdQVgwgsehWvHb3HEI0Q-qekvAzYyU"
FLYWHEEL_TAB_NAME = "Prog<>Placement"  # confirmed from the tab bar screenshot (gid=1147350782)

# The Google Sheet this pipeline writes its output into (5 tabs, cleared & rewritten
# each run). Must be shared with the service account's client_email as Editor.
OUTPUT_SHEET_ID = "1hU8fx6qYS_A6RYHfG5n_O63Dr_AnsjcfrHFtGzHX3Xk"

TAB_ENROLLMENT = "Batch Enrollment (Metabase)"
TAB_PICKED = "Picked Students (raw)"
TAB_MAIN_PIVOT = "Batch Start to Picked date"
TAB_FRESHERS = "Batch Start-Picked (Freshers)"     # Excel/Sheets tab-name limit is 31 chars
TAB_UNEMPLOYED = "Batch Start-Picked (Unemployed)"

BATCH_TITLE_FILTER = "Professional Certificate Course In Data Science%"
GEM_LABEL_IDS = (728, 729)  # technologies_label: 728 = GEM, 729 = Non-GEM

# "Type of Experience" column in the Prog<>Placement tab — used for the Freshers /
# Unemployed views. Column letter and value spellings confirmed against the live sheet.
TYPE_OF_EXPERIENCE_COLUMN = "I"
FRESHER_VALUE = "Fresher"
UNEMPLOYED_VALUE = "Career Gap/ Non Working"

NWINDOWS = 18  # Within 30/60/.../540 Days

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

# Fallback native SQL, used only if METABASE_QUESTION_ID is unset. Keep this in sync
# with enrollment_query.sql's first query.
ENROLLMENT_SQL = f"""
select
  c.id as course_id,
  c.title as batch_name,
  c.start_timestamp as batch_start_date,
  coalesce(tl.name, '') as gem_status,
  count(distinct cum.user_id) filter (where cum.status in (8,11,12,30)) as initially_enrolled,
  count(distinct cum.user_id) filter (
    where cum.status = 8
    and not exists (
      select 1 from courses_courseuserlabelmapping cclm2
      join technologies_label tl2 on tl2.id = cclm2.label_id
      where cclm2.course_user_mapping_id = cum.id and tl2.name = 'NBFC - Refund Requested'
    )
  ) as currently_enrolled,
  count(distinct cum.user_id) filter (
    where cum.status = 8
    and exists (
      select 1 from courses_courseuserlabelmapping cclm2
      join technologies_label tl2 on tl2.id = cclm2.label_id
      where cclm2.course_user_mapping_id = cum.id and tl2.name = 'NBFC - Refund Requested'
    )
  ) as refund_requested,
  count(distinct cum.user_id) filter (where cum.status in (11,12)) as course_cancellation,
  count(distinct cum.user_id) filter (where cum.status = 30) as deferred
from courses_course c
join courses_courseusermapping cum on cum.course_id = c.id
left join courses_courseuserlabelmapping cclm on cclm.course_user_mapping_id = cum.id and cclm.label_id in {GEM_LABEL_IDS}
left join technologies_label tl on tl.id = cclm.label_id
where c.title ilike '{BATCH_TITLE_FILTER}'
group by c.id, c.title, c.start_timestamp, tl.name
order by c.start_timestamp, gem_status;
""".strip()

GEM_MAP_SQL = f"""
select cum.user_id, c.title as batch_name, tl.name as gem_status
from courses_courseusermapping cum
join courses_course c on c.id = cum.course_id
join courses_courseuserlabelmapping cclm on cclm.course_user_mapping_id = cum.id and cclm.label_id in {GEM_LABEL_IDS}
join technologies_label tl on tl.id = cclm.label_id
where c.title ilike '{BATCH_TITLE_FILTER}'
order by c.title, cum.user_id;
""".strip()


def _metabase_query(sql):
    headers = {"X-API-Key": METABASE_API_KEY, "Content-Type": "application/json"}
    url = f"{METABASE_URL}/api/dataset"
    payload = {"type": "native", "native": {"query": sql}, "database": METABASE_DATABASE_ID}
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    data = resp.json()["data"]
    cols = [c["name"] for c in data["cols"]]
    return [dict(zip(cols, r)) for r in data["rows"]]


# ----------------------------------------------------------------------------
# 1. Metabase — Batch Enrollment, with the GEM/Non-GEM A-B split folded in
# ----------------------------------------------------------------------------
def fetch_enrollment_rows():
    """Returns rows shaped for the sheet: for a batch with a real GEM+Non-GEM
    split, 3 rows come out (unsplit original, "<batch> — GEM (A)", "<batch> —
    Non-GEM (B)"); every other batch comes out as a single unsplit row."""
    if METABASE_QUESTION_ID:
        headers = {"X-API-Key": METABASE_API_KEY, "Content-Type": "application/json"}
        url = f"{METABASE_URL}/api/card/{METABASE_QUESTION_ID}/query"
        resp = requests.post(url, headers=headers, timeout=60)
        resp.raise_for_status()
        data = resp.json()["data"]
        cols = [c["name"] for c in data["cols"]]
        raw_rows = [dict(zip(cols, r)) for r in data["rows"]]
    else:
        raw_rows = _metabase_query(ENROLLMENT_SQL)

    if not raw_rows:
        raise RuntimeError("Metabase returned 0 enrollment rows — check the query/credentials.")

    by_batch = {}
    order = []
    for r in raw_rows:
        bname = r["batch_name"]
        if bname not in by_batch:
            by_batch[bname] = {"": None, "GEM": None, "Non-GEM": None}
            order.append(bname)
        gem_status = r.get("gem_status") or ""
        by_batch[bname][gem_status if gem_status in ("GEM", "Non-GEM") else ""] = r

    out = []
    for bname in order:
        bucket = by_batch[bname]
        base = bucket[""] or bucket["GEM"] or bucket["Non-GEM"]
        out.append({
            "batch_name": bname, "base_batch": bname, "gem_status": "",
            "batch_start_date": base["batch_start_date"],
            "initially_enrolled": bucket[""]["initially_enrolled"] if bucket[""] else 0,
            "currently_enrolled": bucket[""]["currently_enrolled"] if bucket[""] else 0,
            "refund_requested": bucket[""]["refund_requested"] if bucket[""] else 0,
            "course_cancellation": bucket[""]["course_cancellation"] if bucket[""] else 0,
            "deferred": bucket[""]["deferred"] if bucket[""] else 0,
        })
        if bucket["GEM"] and bucket["Non-GEM"]:
            for tag, key in (("GEM (A)", "GEM"), ("Non-GEM (B)", "Non-GEM")):
                g = bucket[key]
                out.append({
                    "batch_name": f"{bname} — {tag}", "base_batch": bname, "gem_status": key,
                    "batch_start_date": g["batch_start_date"],
                    "initially_enrolled": g["initially_enrolled"],
                    "currently_enrolled": g["currently_enrolled"],
                    "refund_requested": g["refund_requested"],
                    "course_cancellation": g["course_cancellation"],
                    "deferred": g["deferred"],
                })
    return out


def fetch_gem_map():
    """(batch_name, user_id) -> 'GEM' | 'Non-GEM', used to tag Picked Students raw."""
    try:
        rows = _metabase_query(GEM_MAP_SQL)
    except Exception as e:
        print(f"  ! GEM map query failed ({e}) — Picked rows will be left untagged.", file=sys.stderr)
        return {}
    return {(r["batch_name"], str(r["user_id"])): r["gem_status"] for r in rows}


# ----------------------------------------------------------------------------
# 2. Google Sheets — auth, and Picked Students (raw) read from FlyWheel
# ----------------------------------------------------------------------------
def _sheets_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds = service_account.Credentials.from_service_account_file(
        GOOGLE_SERVICE_ACCOUNT_JSON,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def fetch_picked_rows(service, target_batches, gem_map):
    # Pull only the 5 columns we actually use instead of the whole 82-column sheet:
    # A=UserID, D=Batch, I=Type of Experience, U=Status, V=Picked Date (adjust letters
    # if the FlyWheel sheet's column order changes).
    ranges = [
        f"'{FLYWHEEL_TAB_NAME}'!A2:A",
        f"'{FLYWHEEL_TAB_NAME}'!D2:D",
        f"'{FLYWHEEL_TAB_NAME}'!{TYPE_OF_EXPERIENCE_COLUMN}2:{TYPE_OF_EXPERIENCE_COLUMN}",
        f"'{FLYWHEEL_TAB_NAME}'!U2:U",
        f"'{FLYWHEEL_TAB_NAME}'!V2:V",
    ]
    result = service.spreadsheets().values().batchGet(
        spreadsheetId=FLYWHEEL_SHEET_ID, ranges=ranges
    ).execute()
    value_ranges = result["valueRanges"]
    user_ids = [r[0] if r else "" for r in value_ranges[0].get("values", [])]
    batches = [r[0] if r else "" for r in value_ranges[1].get("values", [])]
    experience_types = [r[0] if r else "" for r in value_ranges[2].get("values", [])]
    statuses = [r[0] if r else "" for r in value_ranges[3].get("values", [])]
    picked_dates = [r[0] if r else "" for r in value_ranges[4].get("values", [])]

    n = max(len(user_ids), len(batches), len(experience_types), len(statuses), len(picked_dates))

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
        gem_status = gem_map.get((batch, str(uid)), "")
        experience_type = get(experience_types, i).strip()
        out.append((uid, batch, pdate, gem_status, experience_type))
    return out


# ----------------------------------------------------------------------------
# 3. Push helpers — write a tab's values into the OUTPUT sheet and format it
# ----------------------------------------------------------------------------
NAVY_RGB = {"red": 0.122, "green": 0.220, "blue": 0.392}


def _get_or_create_sheet(service, spreadsheet_id, title):
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]
    resp = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": title}}}]},
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]


def _push_values(service, spreadsheet_id, title, values):
    sheet_id = _get_or_create_sheet(service, spreadsheet_id, title)
    # Clear the whole tab first so a shorter run doesn't leave stale rows behind.
    service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=f"'{title}'").execute()
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=f"'{title}'!A1",
        valueInputOption="USER_ENTERED", body={"values": values},
    ).execute()
    return sheet_id


def _apply_formatting(service, spreadsheet_id, sheet_id, num_cols, num_data_rows,
                       date_col=None, pct_cols=None):
    requests_ = [
        {  # bold white-on-navy header row
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                           "startColumnIndex": 0, "endColumnIndex": num_cols},
                "cell": {"userEnteredFormat": {
                    "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                    "backgroundColor": NAVY_RGB,
                }},
                "fields": "userEnteredFormat(textFormat,backgroundColor)",
            }
        },
        {  # freeze header row
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
    ]
    if date_col is not None:
        requests_.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 1 + num_data_rows,
                           "startColumnIndex": date_col, "endColumnIndex": date_col + 1},
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "dd-mmm-yyyy"}}},
                "fields": "userEnteredFormat.numberFormat",
            }
        })
    for col in (pct_cols or []):
        requests_.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 1 + num_data_rows,
                           "startColumnIndex": col, "endColumnIndex": col + 1},
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}},
                "fields": "userEnteredFormat.numberFormat",
            }
        })
        requests_.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 1 + num_data_rows,
                                 "startColumnIndex": col, "endColumnIndex": col + 1}],
                    "gradientRule": {
                        "minpoint": {"type": "MIN", "color": {"red": 0.973, "green": 0.412, "blue": 0.420}},
                        "midpoint": {"type": "PERCENTILE", "value": "50",
                                     "color": {"red": 1, "green": 0.922, "blue": 0.518}},
                        "maxpoint": {"type": "MAX", "color": {"red": 0.388, "green": 0.745, "blue": 0.482}},
                    },
                },
                "index": 0,
            }
        })
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests_}).execute()


# ----------------------------------------------------------------------------
# 4. Build each tab's values (headers + rows of literal values / formula strings)
# ----------------------------------------------------------------------------
def _enrollment_tab_values(enroll_rows):
    headers = ["Batch", "Base Batch", "GEM Status", "Batch Start Date", "Initially Enrolled",
               "Currently Enrolled", "Refund Requested", "Course Cancellation", "Deferred"]
    rows = [headers]
    for row in enroll_rows:
        start_date = row["batch_start_date"]
        if isinstance(start_date, str):
            start_date = start_date[:10]  # keep just YYYY-MM-DD, Sheets parses it as a date
        rows.append([
            row["batch_name"], row["base_batch"], row["gem_status"], start_date,
            int(row["initially_enrolled"]), int(row["currently_enrolled"]),
            int(row["refund_requested"]), int(row["course_cancellation"]), int(row["deferred"]),
        ])
    return rows


def _picked_tab_values(picked_rows):
    headers = ["UserID", "Batch", "Picked Date", "GEM Status", "Type of Experience"]
    rows = [headers]
    for uid, batch, pdate_str, gem_status, experience_type in picked_rows:
        uid_val = int(uid) if str(uid).strip().isdigit() else uid
        rows.append([uid_val, batch, pdate_str, gem_status, experience_type])
    return rows


def _main_pivot_values(enroll_rows, last_enroll_row, last_picked_row):
    headers = ["Batch", "Batch (A/B)", "Base Batch", "GEM Status (raw)", "Batch Start Date",
               "Initially Enrolled", "Currently Enrolled", "Picked Students"]
    for w in range(NWINDOWS):
        headers += [f"Within {(w + 1) * 30} Days", "% of Currently Enrolled"]
    headers.append("Not Placed yet")

    N_TAIL_ROWS = 9
    data_rows = enroll_rows + [None] * N_TAIL_ROWS
    first_data_row, last_data_row = 2, 2 + len(data_rows) - 1

    enroll_range_name = f"'{TAB_ENROLLMENT}'!$A$2:$A${last_enroll_row}"
    enroll_col = lambda letter: f"'{TAB_ENROLLMENT}'!${letter}$2:${letter}${last_enroll_row}"
    picked_batch_range = f"'{TAB_PICKED}'!$B$2:$B${last_picked_row}"
    picked_date_range = f"'{TAB_PICKED}'!$C$2:$C${last_picked_row}"
    picked_gem_range = f"'{TAB_PICKED}'!$D$2:$D${last_picked_row}"

    rows = [headers]
    for idx, r in enumerate(range(first_data_row, last_data_row + 1)):
        src = data_rows[idx]
        Ar = f"$A{r}"
        row_vals = [src["batch_name"] if src else ""]

        row_vals.append(
            f'=IF({Ar}="","",IFERROR(IF(INDEX({enroll_col("C")},MATCH({Ar},{enroll_range_name},0))="GEM","A",'
            f'IF(INDEX({enroll_col("C")},MATCH({Ar},{enroll_range_name},0))="Non-GEM","B","")),""))'
        )
        row_vals.append(f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("B")},MATCH({Ar},{enroll_range_name},0)),""))')
        row_vals.append(f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("C")},MATCH({Ar},{enroll_range_name},0)),""))')
        row_vals.append(f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("D")},MATCH({Ar},{enroll_range_name},0)),""))')
        row_vals.append(f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("E")},MATCH({Ar},{enroll_range_name},0)),""))')
        row_vals.append(f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("F")},MATCH({Ar},{enroll_range_name},0)),""))')

        BaseBatch, GemFilter = f"$C{r}", f"$D{r}"
        Dr, Br = f"$G{r}", f"$E{r}"

        row_vals.append(
            f'=IF({Ar}="","",IF({GemFilter}="",COUNTIF({picked_batch_range},{BaseBatch}),'
            f"COUNTIFS({picked_batch_range},{BaseBatch},{picked_gem_range},{GemFilter})))"
        )

        col_base = 8  # 1-indexed column of "Picked Students"
        for w in range(NWINDOWS):
            days = (w + 1) * 30
            within_col = col_base + 1 + w * 2
            within_letter = _col_letter(within_col)

            within_formula = (
                f'=IF({Ar}="","",IF({Br}="","",IF(TODAY()<{Br}+{days},"-",'
                f'IF({GemFilter}="",'
                f'SUMPRODUCT(({picked_batch_range}={BaseBatch})*({picked_date_range}<>"")*(({picked_date_range}-{Br})<{days})),'
                f'SUMPRODUCT(({picked_batch_range}={BaseBatch})*({picked_gem_range}={GemFilter})*({picked_date_range}<>"")*(({picked_date_range}-{Br})<{days}))'
                f"))))"
            )
            row_vals.append(within_formula)

            pct_formula = f'=IF({Ar}="","",IF({within_letter}{r}="-","-",IF({Dr}=0,0,{within_letter}{r}/{Dr})))'
            row_vals.append(pct_formula)

        last_within_letter = _col_letter(col_base + 1 + (NWINDOWS - 1) * 2)
        row_vals.append(f'=IF({Ar}="","",IF({last_within_letter}{r}="-","-",$H{r}-{last_within_letter}{r}))')
        rows.append(row_vals)

    total_row = last_data_row + 1
    ab_range = f"$B${first_data_row}:$B${last_data_row}"
    total_vals = ["TOTAL", "", "", "", ""]
    for letter in ("F", "G", "H"):
        total_vals.append(f'=SUMIF({ab_range},"",{letter}{first_data_row}:{letter}{last_data_row})')
    for w in range(NWINDOWS):
        within_col = col_base + 1 + w * 2
        wl = _col_letter(within_col)
        total_vals.append(f'=SUMIF({ab_range},"",{wl}{first_data_row}:{wl}{last_data_row})')
        total_vals.append(
            f'=IF(SUMIF({ab_range},"",$G${first_data_row}:$G${last_data_row})=0,0,'
            f'{wl}{total_row}/SUMIF({ab_range},"",$G${first_data_row}:$G${last_data_row}))'
        )
    last_within_letter = _col_letter(col_base + 1 + (NWINDOWS - 1) * 2)
    total_vals.append(f"=H{total_row}-{last_within_letter}{total_row}")
    rows.append(total_vals)

    pct_cols_0based = [col_base + w * 2 + 1 for w in range(NWINDOWS)]  # 0-indexed
    return rows, first_data_row, last_data_row, len(headers), pct_cols_0based


def _experience_pivot_values(target_value, base_batches, last_enroll_row, last_picked_row):
    headers = ["Batch", "Batch Start Date", "Initially Enrolled", "Currently Enrolled", "Picked Students"]
    for w in range(NWINDOWS):
        headers += [f"Within {(w + 1) * 30} Days", "% of Currently Enrolled"]
    headers.append("Not Placed yet")

    N_TAIL_ROWS = 9
    data_rows = base_batches + [None] * N_TAIL_ROWS
    first_data_row, last_data_row = 2, 2 + len(data_rows) - 1

    enroll_range_name = f"'{TAB_ENROLLMENT}'!$A$2:$A${last_enroll_row}"
    enroll_col = lambda letter: f"'{TAB_ENROLLMENT}'!${letter}$2:${letter}${last_enroll_row}"
    picked_batch_range = f"'{TAB_PICKED}'!$B$2:$B${last_picked_row}"
    picked_date_range = f"'{TAB_PICKED}'!$C$2:$C${last_picked_row}"
    picked_experience_range = f"'{TAB_PICKED}'!$E$2:$E${last_picked_row}"
    target_lit = target_value.replace('"', '""')

    col_base = 5  # 1-indexed column of "Picked Students"
    rows = [headers]
    for idx, r in enumerate(range(first_data_row, last_data_row + 1)):
        src = data_rows[idx]
        Ar = f"$A{r}"
        row_vals = [src["batch_name"] if src else ""]
        row_vals.append(f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("D")},MATCH({Ar},{enroll_range_name},0)),""))')
        row_vals.append(f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("E")},MATCH({Ar},{enroll_range_name},0)),""))')
        row_vals.append(f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("F")},MATCH({Ar},{enroll_range_name},0)),""))')

        Dr, Br = f"$D{r}", f"$B{r}"
        row_vals.append(
            f'=IF({Ar}="","",COUNTIFS({picked_batch_range},{Ar},{picked_experience_range},"{target_lit}"))'
        )

        for w in range(NWINDOWS):
            days = (w + 1) * 30
            within_col = col_base + 1 + w * 2
            within_letter = _col_letter(within_col)
            within_formula = (
                f'=IF({Ar}="","",IF({Br}="","",IF(TODAY()<{Br}+{days},"-",'
                f'SUMPRODUCT(({picked_batch_range}={Ar})*({picked_experience_range}="{target_lit}")*'
                f'({picked_date_range}<>"")*(({picked_date_range}-{Br})<{days})))))'
            )
            row_vals.append(within_formula)
            pct_formula = f'=IF({Ar}="","",IF({within_letter}{r}="-","-",IF({Dr}=0,0,{within_letter}{r}/{Dr})))'
            row_vals.append(pct_formula)

        last_within_letter = _col_letter(col_base + 1 + (NWINDOWS - 1) * 2)
        row_vals.append(f'=IF({Ar}="","",IF({last_within_letter}{r}="-","-",$E{r}-{last_within_letter}{r}))')
        rows.append(row_vals)

    total_row = last_data_row + 1
    total_vals = ["TOTAL", ""]
    for letter in ("C", "D", "E"):
        total_vals.append(f"=SUM({letter}{first_data_row}:{letter}{last_data_row})")
    for w in range(NWINDOWS):
        within_col = col_base + 1 + w * 2
        wl = _col_letter(within_col)
        total_vals.append(f"=SUM({wl}{first_data_row}:{wl}{last_data_row})")
        total_vals.append(
            f'=IF(SUM($D${first_data_row}:$D${last_data_row})=0,0,'
            f'{wl}{total_row}/SUM($D${first_data_row}:$D${last_data_row}))'
        )
    last_within_letter = _col_letter(col_base + 1 + (NWINDOWS - 1) * 2)
    total_vals.append(f"=E{total_row}-{last_within_letter}{total_row}")
    rows.append(total_vals)

    pct_cols_0based = [col_base + w * 2 + 1 for w in range(NWINDOWS)]
    return rows, first_data_row, last_data_row, len(headers), pct_cols_0based


def _col_letter(n):
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------
def main():
    print("Fetching enrollment data from Metabase (with GEM/Non-GEM split)...", file=sys.stderr)
    enroll_rows = fetch_enrollment_rows()
    print(f"  -> {len(enroll_rows)} rows (incl. A/B sub-rows)", file=sys.stderr)

    target_batches = {r["base_batch"] for r in enroll_rows}

    print("Fetching GEM/Non-GEM user-level mapping from Metabase...", file=sys.stderr)
    gem_map = fetch_gem_map()
    print(f"  -> {len(gem_map)} tagged users", file=sys.stderr)

    service = _sheets_service()

    print("Fetching Picked rows from the FlyWheel Google Sheet...", file=sys.stderr)
    picked_rows = fetch_picked_rows(service, target_batches, gem_map)
    print(f"  -> {len(picked_rows)} picked rows", file=sys.stderr)

    last_enroll_row = len(enroll_rows) + 1
    last_picked_row = len(picked_rows) + 1

    print(f"Writing {TAB_ENROLLMENT} ...", file=sys.stderr)
    enroll_values = _enrollment_tab_values(enroll_rows)
    sid = _push_values(service, OUTPUT_SHEET_ID, TAB_ENROLLMENT, enroll_values)
    _apply_formatting(service, OUTPUT_SHEET_ID, sid, len(enroll_values[0]), len(enroll_values) - 1, date_col=3)

    print(f"Writing {TAB_PICKED} ...", file=sys.stderr)
    picked_values = _picked_tab_values(picked_rows)
    sid = _push_values(service, OUTPUT_SHEET_ID, TAB_PICKED, picked_values)
    _apply_formatting(service, OUTPUT_SHEET_ID, sid, len(picked_values[0]), len(picked_values) - 1, date_col=2)

    print(f"Writing {TAB_MAIN_PIVOT} ...", file=sys.stderr)
    rows, first_r, last_r, ncols, pct_cols = _main_pivot_values(enroll_rows, last_enroll_row, last_picked_row)
    sid = _push_values(service, OUTPUT_SHEET_ID, TAB_MAIN_PIVOT, rows)
    _apply_formatting(service, OUTPUT_SHEET_ID, sid, ncols, last_r - first_r + 2, date_col=4, pct_cols=pct_cols)

    base_batches_only = [row for row in enroll_rows if not row["gem_status"]]

    print(f"Writing {TAB_FRESHERS} ...", file=sys.stderr)
    rows, first_r, last_r, ncols, pct_cols = _experience_pivot_values(
        FRESHER_VALUE, base_batches_only, last_enroll_row, last_picked_row)
    sid = _push_values(service, OUTPUT_SHEET_ID, TAB_FRESHERS, rows)
    _apply_formatting(service, OUTPUT_SHEET_ID, sid, ncols, last_r - first_r + 2, date_col=1, pct_cols=pct_cols)

    print(f"Writing {TAB_UNEMPLOYED} ...", file=sys.stderr)
    rows, first_r, last_r, ncols, pct_cols = _experience_pivot_values(
        UNEMPLOYED_VALUE, base_batches_only, last_enroll_row, last_picked_row)
    sid = _push_values(service, OUTPUT_SHEET_ID, TAB_UNEMPLOYED, rows)
    _apply_formatting(service, OUTPUT_SHEET_ID, sid, ncols, last_r - first_r + 2, date_col=1, pct_cols=pct_cols)

    print("Done.", file=sys.stderr)


if __name__ == "__main__":
    main()
