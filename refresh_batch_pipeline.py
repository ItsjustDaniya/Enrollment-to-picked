#!/usr/bin/env python3
"""
refresh_batch_pipeline.py

Rebuilds Batch_Start_to_Picked_date.xlsx end-to-end, twice a day, with no manual CSV
upload step:

  1. Batch Enrollment (Metabase) tab  <- Metabase REST API (saved question / native SQL),
                                          including the GEM / Non-GEM A-B split: any batch
                                          that has real GEM(728)/Non-GEM(729) label tagging
                                          gets 2 extra rows under it ("<batch> — GEM (A)",
                                          "<batch> — Non-GEM (B)") alongside its own
                                          unsplit total row.
  2. Picked Students (raw) tab        <- Google Sheets API, pulling ONLY the columns
                                          needed from the FlyWheel tracker (UserID,
                                          Batch, Status, Picked Date), tagged with each
                                          user's GEM Status via a second Metabase query
                                          so the A/B split also applies to Picked counts.
  3. Batch Start to Picked date tab   <- rebuilt with name-keyed INDEX/MATCH, "-" for
                                          windows that haven't elapsed yet, "% of Currently
                                          Enrolled", and the GEM(A)/Non-GEM(B) rows are
                                          auto-inserted under their parent batch (TOTAL
                                          only sums the unsplit rows, so nothing double-counts).

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
import os
import sys
import tempfile

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter, column_index_from_string

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

OUTPUT_PATH = "Batch_Start_to_Picked_date.xlsx"

BATCH_TITLE_FILTER = "Professional Certificate Course In Data Science%"
GEM_LABEL_IDS = (728, 729)  # technologies_label: 728 = GEM, 729 = Non-GEM

# "Type of Experience" column in the Prog<>Placement tab — used for the two extra
# Freshers / Unemployed "Batch Start to Picked date" views. Column letter and the
# exact value spellings were confirmed against the live sheet.
TYPE_OF_EXPERIENCE_COLUMN = "I"
FRESHER_VALUE = "Fresher"
UNEMPLOYED_VALUE = "Career Gap/ Non Working"

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
    """Returns rows shaped for the workbook: for a batch with a real GEM+Non-GEM
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

    # Group by base batch name (gem_status may be missing entirely if the saved
    # question hasn't been updated yet — treat that as "no split data available").
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
# 2. Google Sheets — Picked Students (raw), pulling only the needed columns
# ----------------------------------------------------------------------------
def fetch_picked_rows(target_batches, gem_map):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds = service_account.Credentials.from_service_account_file(
        GOOGLE_SERVICE_ACCOUNT_JSON,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    service = build("sheets", "v4", credentials=creds)

    # Pull only the 5 columns we actually use instead of the whole 82-column sheet:
    # A=UserID, D=Batch, I=Type of Experience, U=Status, V=Picked Date (adjust letters
    # if the FlyWheel sheet's column order changes — A/D/U/V match cols 0,3,20,21 in
    # the CSV export this pipeline was originally built from; I was confirmed from a
    # screenshot of the live sheet's "Type of Experience" column).
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


def _build_experience_pivot(wb, sheet_name, target_value, base_batches, last_enroll_row,
                             picked_batch_range, picked_date_range, picked_experience_range,
                             header_font, header_fill, data_font, note_font):
    """A simplified 'Batch Start to Picked date' view (no GEM/A-B columns) filtered to
    Picked Students whose 'Type of Experience' equals target_value exactly."""
    NAVY_BOLD = Font(name="Arial", bold=True)
    NWINDOWS = 18
    ws = wb.create_sheet(sheet_name)
    headers = ["Batch", "Batch Start Date", "Initially Enrolled", "Currently Enrolled", "Picked Students"]
    for w in range(NWINDOWS):
        days = (w + 1) * 30
        headers += [f"Within {days} Days", "% of Currently Enrolled"]
    headers.append("Not Placed yet")
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font, c.fill = header_font, header_fill

    N_TAIL_ROWS = 9
    data_rows = base_batches + [None] * N_TAIL_ROWS
    first_data_row, last_data_row = 2, 2 + len(data_rows) - 1

    enroll_range_name = f"'Batch Enrollment (Metabase)'!$A$2:$A${last_enroll_row}"
    enroll_col = lambda letter: f"'Batch Enrollment (Metabase)'!${letter}$2:${letter}${last_enroll_row}"
    target_lit = target_value.replace('"', '""')

    for idx, r in enumerate(range(first_data_row, last_data_row + 1)):
        src = data_rows[idx]
        a_cell = ws.cell(row=r, column=1)
        if src:
            a_cell.value = src["batch_name"]
        a_cell.font = data_font
        Ar = f"$A{r}"

        b = ws.cell(row=r, column=2,
                    value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("D")},MATCH({Ar},{enroll_range_name},0)),""))')
        b.font, b.number_format = data_font, "dd-mmm-yyyy"

        c_ = ws.cell(row=r, column=3,
                     value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("E")},MATCH({Ar},{enroll_range_name},0)),""))')
        c_.font = data_font

        d_ = ws.cell(row=r, column=4,
                     value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("F")},MATCH({Ar},{enroll_range_name},0)),""))')
        d_.font = data_font

        Dr, Br = f"$D{r}", f"$B{r}"

        e_ = ws.cell(row=r, column=5,
                     value=(f'=IF({Ar}="","",COUNTIFS({picked_batch_range},{Ar},'
                            f'{picked_experience_range},"{target_lit}"))'))
        e_.font = data_font

        col_base = 5
        for w in range(NWINDOWS):
            days = (w + 1) * 30
            within_col = col_base + 1 + w * 2
            pct_col = within_col + 1
            within_letter = get_column_letter(within_col)

            within_formula = (
                f'=IF({Ar}="","",IF({Br}="","",IF(TODAY()<{Br}+{days},"-",'
                f'SUMPRODUCT(({picked_batch_range}={Ar})*({picked_experience_range}="{target_lit}")*'
                f'({picked_date_range}<>"")*(({picked_date_range}-{Br})<{days})))))'
            )
            wcell = ws.cell(row=r, column=within_col, value=within_formula)
            wcell.font = data_font

            pct_formula = (
                f'=IF({Ar}="","",IF({within_letter}{r}="-","-",'
                f"IF({Dr}=0,0,{within_letter}{r}/{Dr})))"
            )
            pcell = ws.cell(row=r, column=pct_col, value=pct_formula)
            pcell.font, pcell.number_format = data_font, "0%"

        last_within_letter = get_column_letter(col_base + 1 + (NWINDOWS - 1) * 2)
        np_cell = ws.cell(row=r, column=len(headers),
                           value=f'=IF({Ar}="","",IF({last_within_letter}{r}="-","-",$E{r}-{last_within_letter}{r}))')
        np_cell.font = data_font

    total_row = last_data_row + 1
    ws.cell(row=total_row, column=1, value="TOTAL").font = NAVY_BOLD
    for letter in ("C", "D", "E"):
        cell = ws.cell(row=total_row, column=column_index_from_string(letter),
                        value=f"=SUM({letter}{first_data_row}:{letter}{last_data_row})")
        cell.font = NAVY_BOLD
    for w in range(NWINDOWS):
        within_col = col_base + 1 + w * 2
        pct_col = within_col + 1
        wl = get_column_letter(within_col)
        wcell = ws.cell(row=total_row, column=within_col,
                         value=f"=SUM({wl}{first_data_row}:{wl}{last_data_row})")
        wcell.font = NAVY_BOLD
        pcell = ws.cell(row=total_row, column=pct_col,
                         value=(f'=IF(SUM($D${first_data_row}:$D${last_data_row})=0,0,'
                                f'{wl}{total_row}/SUM($D${first_data_row}:$D${last_data_row}))'))
        pcell.font, pcell.number_format = NAVY_BOLD, "0%"
    last_within_letter = get_column_letter(col_base + 1 + (NWINDOWS - 1) * 2)
    ws.cell(row=total_row, column=len(headers),
            value=f"=E{total_row}-{last_within_letter}{total_row}").font = NAVY_BOLD

    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 14
    for j in range(3, len(headers) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 13
    ws.freeze_panes = "C2"

    for w in range(NWINDOWS):
        pct_col = col_base + 1 + w * 2 + 1
        letter = get_column_letter(pct_col)
        rng = f"{letter}{first_data_row}:{letter}{last_data_row}"
        ws.conditional_formatting.add(rng, ColorScaleRule(
            start_type="min", start_color="F8696B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="63BE7B"))
    np_letter = get_column_letter(len(headers))
    ws.conditional_formatting.add(f"{np_letter}{first_data_row}:{np_letter}{last_data_row}", ColorScaleRule(
        start_type="min", start_color="63BE7B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="F8696B"))
    ws.conditional_formatting.add(f"E{first_data_row}:E{last_data_row}", ColorScaleRule(
        start_type="min", start_color="FFFFFFFF", end_type="max", end_color="FF9DC3E6"))

    ws.cell(row=1, column=len(headers) + 2,
            value=f'Filtered to Picked Students where "Type of Experience" = "{target_value}" '
                  f'(FlyWheel column {TYPE_OF_EXPERIENCE_COLUMN}). Same "-" and % of Currently '
                  f'Enrolled logic as the main tab.').font = note_font


# ----------------------------------------------------------------------------
# 3. Build the workbook
# ----------------------------------------------------------------------------
def build_workbook(enroll_rows, picked_rows, output_path):
    NAVY, WHITE, GREY = "FF1F3864", "FFFFFFFF", "FF333333"
    header_font = Font(name="Arial", bold=True, color=WHITE)
    header_fill = PatternFill("solid", fgColor=NAVY, bgColor=GREY)
    data_font = Font(name="Arial")
    sub_font = Font(name="Arial", italic=True)
    note_font = Font(name="Arial", italic=True, size=9, color="FF808080")
    sub_fill = PatternFill("solid", fgColor="FFF4F7FC")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- Batch Enrollment (Metabase) ---
    ws1 = wb.create_sheet("Batch Enrollment (Metabase)")
    headers1 = ["Batch", "Base Batch", "GEM Status", "Batch Start Date", "Initially Enrolled",
                "Currently Enrolled", "Refund Requested", "Course Cancellation", "Deferred"]
    for j, h in enumerate(headers1, start=1):
        c = ws1.cell(row=1, column=j, value=h)
        c.font, c.fill = header_font, header_fill
    for i, row in enumerate(enroll_rows, start=2):
        start_date = row["batch_start_date"]
        if isinstance(start_date, str):
            start_date = dt.datetime.fromisoformat(start_date.replace("Z", "+00:00")).replace(tzinfo=None)
        is_sub = bool(row["gem_status"])
        vals = [row["batch_name"], row["base_batch"], row["gem_status"], start_date,
                int(row["initially_enrolled"]), int(row["currently_enrolled"]),
                int(row["refund_requested"]), int(row["course_cancellation"]), int(row["deferred"])]
        for j, v in enumerate(vals, start=1):
            c = ws1.cell(row=i, column=j, value=v)
            c.font = sub_font if is_sub else data_font
            if is_sub:
                c.fill = sub_fill
            if j == 4:
                c.number_format = "dd-mmm-yyyy"
    last_enroll_row = len(enroll_rows) + 1
    ws1.cell(row=1, column=11,
              value=f"Source: Metabase (Newton School DB). GEM/Non-GEM split from technologies_label "
                    f"ids 728/729. Refreshed: {dt.date.today().isoformat()}.").font = note_font
    ws1.column_dimensions["A"].width = 62
    ws1.column_dimensions["B"].width = 55
    for col in "CDEFGHI":
        ws1.column_dimensions[col].width = 16
    ws1.freeze_panes = "A2"

    # --- Picked Students (raw) ---
    ws2 = wb.create_sheet("Picked Students (raw)")
    for j, h in enumerate(["UserID", "Batch", "Picked Date", "GEM Status", "Type of Experience"], start=1):
        c = ws2.cell(row=1, column=j, value=h)
        c.font, c.fill = header_font, header_fill
    for i, (uid, batch, pdate_str, gem_status, experience_type) in enumerate(picked_rows, start=2):
        uid_val = int(uid) if str(uid).strip().isdigit() else uid
        pdate = dt.datetime.strptime(pdate_str, "%Y-%m-%d") if pdate_str else None
        ws2.cell(row=i, column=1, value=uid_val).font = data_font
        ws2.cell(row=i, column=2, value=batch).font = data_font
        c3 = ws2.cell(row=i, column=3, value=pdate)
        c3.font = data_font
        if pdate:
            c3.number_format = "dd-mmm-yyyy"
        ws2.cell(row=i, column=4, value=gem_status).font = data_font
        ws2.cell(row=i, column=5, value=experience_type).font = data_font
    last_picked_row = len(picked_rows) + 1
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 60
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 12
    ws2.column_dimensions["E"].width = 26
    ws2.freeze_panes = "A2"
    ws2.cell(row=1, column=7,
              value=f"Source: FlyWheel Google Sheet via Sheets API. GEM Status via Metabase label "
                    f"mapping, Type of Experience via col {TYPE_OF_EXPERIENCE_COLUMN} of "
                    f"'{FLYWHEEL_TAB_NAME}'. Refreshed: {dt.date.today().isoformat()}.").font = note_font
    picked_experience_range = f"'Picked Students (raw)'!$E$2:$E${last_picked_row}"

    # --- Batch Start to Picked date ---
    ws3 = wb.create_sheet("Batch Start to Picked date")
    NWINDOWS = 18
    headers3 = ["Batch", "Batch (A/B)", "Base Batch", "GEM Status (raw)", "Batch Start Date",
                "Initially Enrolled", "Currently Enrolled", "Picked Students"]
    for w in range(NWINDOWS):
        days = (w + 1) * 30
        headers3 += [f"Within {days} Days", "% of Currently Enrolled"]
    headers3.append("Not Placed yet")
    for j, h in enumerate(headers3, start=1):
        c = ws3.cell(row=1, column=j, value=h)
        c.font, c.fill = header_font, header_fill

    N_TAIL_ROWS = 9
    data_rows = enroll_rows + [None] * N_TAIL_ROWS
    first_data_row, last_data_row = 2, 2 + len(data_rows) - 1

    enroll_range_name = f"'Batch Enrollment (Metabase)'!$A$2:$A${last_enroll_row}"
    enroll_col = lambda letter: f"'Batch Enrollment (Metabase)'!${letter}$2:${letter}${last_enroll_row}"
    picked_batch_range = f"'Picked Students (raw)'!$B$2:$B${last_picked_row}"
    picked_date_range = f"'Picked Students (raw)'!$C$2:$C${last_picked_row}"
    picked_gem_range = f"'Picked Students (raw)'!$D$2:$D${last_picked_row}"

    for idx, r in enumerate(range(first_data_row, last_data_row + 1)):
        src = data_rows[idx]
        is_sub = bool(src and src["gem_status"])
        a_cell = ws3.cell(row=r, column=1)
        if src:
            a_cell.value = src["batch_name"]
        a_cell.font = sub_font if is_sub else data_font
        if is_sub:
            a_cell.fill = sub_fill

        Ar = f"$A{r}"

        b = ws3.cell(row=r, column=2,
                     value=(f'=IF({Ar}="","",IFERROR(IF(INDEX({enroll_col("C")},MATCH({Ar},{enroll_range_name},0))="GEM","A",'
                            f'IF(INDEX({enroll_col("C")},MATCH({Ar},{enroll_range_name},0))="Non-GEM","B","")),""))'))
        b.font = sub_font if is_sub else data_font
        if is_sub:
            b.fill = sub_fill

        c_ = ws3.cell(row=r, column=3,
                      value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("B")},MATCH({Ar},{enroll_range_name},0)),""))')
        c_.font = sub_font if is_sub else data_font

        d_ = ws3.cell(row=r, column=4,
                      value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("C")},MATCH({Ar},{enroll_range_name},0)),""))')
        d_.font = sub_font if is_sub else data_font

        e = ws3.cell(row=r, column=5,
                     value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("D")},MATCH({Ar},{enroll_range_name},0)),""))')
        e.font = sub_font if is_sub else data_font
        e.number_format = "dd-mmm-yyyy"

        f_ = ws3.cell(row=r, column=6,
                      value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("E")},MATCH({Ar},{enroll_range_name},0)),""))')
        f_.font = sub_font if is_sub else data_font

        g_ = ws3.cell(row=r, column=7,
                      value=f'=IF({Ar}="","",IFERROR(INDEX({enroll_col("F")},MATCH({Ar},{enroll_range_name},0)),""))')
        g_.font = sub_font if is_sub else data_font

        BaseBatch, GemFilter = f"$C{r}", f"$D{r}"
        Dr, Br, Er = f"$G{r}", f"$E{r}", f"$H{r}"

        h_ = ws3.cell(row=r, column=8,
                      value=(f'=IF({Ar}="","",IF({GemFilter}="",COUNTIF({picked_batch_range},{BaseBatch}),'
                             f"COUNTIFS({picked_batch_range},{BaseBatch},{picked_gem_range},{GemFilter})))"))
        h_.font = sub_font if is_sub else data_font

        col_base = 8
        for w in range(NWINDOWS):
            days = (w + 1) * 30
            within_col = col_base + 1 + w * 2
            pct_col = within_col + 1
            within_letter = get_column_letter(within_col)

            within_formula = (
                f'=IF({Ar}="","",IF({Br}="","",IF(TODAY()<{Br}+{days},"-",'
                f'IF({GemFilter}="",'
                f'SUMPRODUCT(({picked_batch_range}={BaseBatch})*({picked_date_range}<>"")*(({picked_date_range}-{Br})<{days})),'
                f'SUMPRODUCT(({picked_batch_range}={BaseBatch})*({picked_gem_range}={GemFilter})*({picked_date_range}<>"")*(({picked_date_range}-{Br})<{days}))'
                f"))))"
            )
            wcell = ws3.cell(row=r, column=within_col, value=within_formula)
            wcell.font = sub_font if is_sub else data_font

            pct_formula = (
                f'=IF({Ar}="","",IF({within_letter}{r}="-","-",'
                f"IF({Dr}=0,0,{within_letter}{r}/{Dr})))"
            )
            pcell = ws3.cell(row=r, column=pct_col, value=pct_formula)
            pcell.font = sub_font if is_sub else data_font
            pcell.number_format = "0%"
            if is_sub:
                wcell.fill = sub_fill
                pcell.fill = sub_fill

        last_within_letter = get_column_letter(col_base + 1 + (NWINDOWS - 1) * 2)
        np_cell = ws3.cell(row=r, column=len(headers3),
                            value=f'=IF({Ar}="","",IF({last_within_letter}{r}="-","-",{Er}-{last_within_letter}{r}))')
        np_cell.font = sub_font if is_sub else data_font
        if is_sub:
            np_cell.fill = sub_fill

    total_row = last_data_row + 1
    ws3.cell(row=total_row, column=1, value="TOTAL").font = Font(name="Arial", bold=True)
    ab_range = f"$B${first_data_row}:$B${last_data_row}"
    for letter in ("F", "G", "H"):
        cell = ws3.cell(row=total_row, column=column_index_from_string(letter),
                         value=f'=SUMIF({ab_range},"",{letter}{first_data_row}:{letter}{last_data_row})')
        cell.font = Font(name="Arial", bold=True)
    for w in range(NWINDOWS):
        within_col = 8 + 1 + w * 2
        pct_col = within_col + 1
        wl = get_column_letter(within_col)
        wcell = ws3.cell(row=total_row, column=within_col,
                          value=f'=SUMIF({ab_range},"",{wl}{first_data_row}:{wl}{last_data_row})')
        wcell.font = Font(name="Arial", bold=True)
        pcell = ws3.cell(row=total_row, column=pct_col,
                          value=(f'=IF(SUMIF({ab_range},"",$G${first_data_row}:$G${last_data_row})=0,0,'
                                 f'{wl}{total_row}/SUMIF({ab_range},"",$G${first_data_row}:$G${last_data_row}))'))
        pcell.font = Font(name="Arial", bold=True)
        pcell.number_format = "0%"
    last_within_letter = get_column_letter(8 + 1 + (NWINDOWS - 1) * 2)
    ws3.cell(row=total_row, column=len(headers3),
              value=f"=H{total_row}-{last_within_letter}{total_row}").font = Font(name="Arial", bold=True)

    ws3.column_dimensions["A"].width = 62
    ws3.column_dimensions["B"].width = 11
    ws3.column_dimensions["C"].width = 55
    ws3.column_dimensions["D"].width = 13
    ws3.column_dimensions["E"].width = 14
    for j in range(6, len(headers3) + 1):
        ws3.column_dimensions[get_column_letter(j)].width = 13
    ws3.freeze_panes = "F2"

    for w in range(NWINDOWS):
        pct_col = 8 + 1 + w * 2 + 1
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
    ws3.conditional_formatting.add(f"H{first_data_row}:H{last_data_row}", ColorScaleRule(
        start_type="min", start_color="FFFFFFFF", end_type="max", end_color="FF9DC3E6"))

    # --- Batch Start to Picked Date (Freshers) / (Unemployed) ---
    # Same batches as the main tab, but Picked Students / Within-N-Days are filtered to
    # rows in "Picked Students (raw)" whose Type of Experience matches exactly. Uses the
    # unsplit batch rows only (GEM/Non-GEM split is unrelated to experience type).
    # Note: Excel caps sheet names at 31 characters, so "Batch Start to Picked Date
    # (Freshers/Unemployed)" is shortened to "Batch Start-Picked (...)" — same tab,
    # shorter title only.
    base_batches_only = [row for row in enroll_rows if not row["gem_status"]]
    _build_experience_pivot(
        wb, "Batch Start-Picked (Freshers)", FRESHER_VALUE, base_batches_only,
        last_enroll_row, picked_batch_range, picked_date_range, picked_experience_range,
        header_font, header_fill, data_font, note_font,
    )
    _build_experience_pivot(
        wb, "Batch Start-Picked (Unemployed)", UNEMPLOYED_VALUE, base_batches_only,
        last_enroll_row, picked_batch_range, picked_date_range, picked_experience_range,
        header_font, header_fill, data_font, note_font,
    )

    wb.save(output_path)
    return output_path


def main():
    print("Fetching enrollment data from Metabase (with GEM/Non-GEM split)...", file=sys.stderr)
    enroll_rows = fetch_enrollment_rows()
    print(f"  -> {len(enroll_rows)} rows (incl. A/B sub-rows)", file=sys.stderr)

    target_batches = {r["base_batch"] for r in enroll_rows}

    print("Fetching GEM/Non-GEM user-level mapping from Metabase...", file=sys.stderr)
    gem_map = fetch_gem_map()
    print(f"  -> {len(gem_map)} tagged users", file=sys.stderr)

    print("Fetching Picked rows from the FlyWheel Google Sheet...", file=sys.stderr)
    picked_rows = fetch_picked_rows(target_batches, gem_map)
    print(f"  -> {len(picked_rows)} picked rows", file=sys.stderr)

    print(f"Building {OUTPUT_PATH} ...", file=sys.stderr)
    build_workbook(enroll_rows, picked_rows, OUTPUT_PATH)
    print("Done.", file=sys.stderr)


if __name__ == "__main__":
    main()
