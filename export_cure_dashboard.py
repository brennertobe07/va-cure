"""
export_cure_dashboard.py
DPVA Absentee Pipeline -- Cure Dashboard JSON Exporter
------------------------------------------------------
Queries INSTANCE-1 / Absentee and produces data/cure_data.json
for cure_dashboard.html (GitHub Pages, no live DB connection).

Source tables / views (all in Absentee database):
  dbo.Cure_History          - refreshed daily by usp_Refresh_Cure_History
  dbo.Daily_Absentee_List   - current snapshot
  van                       - view joining to Voter.dbo.Van

Party logic (SupportScore, resolved via Voter.dbo.ScoreCycle).
Threshold is 50/50, matching the absentee dashboard (changed from
60/40 on 2026-09-02). There is no 'Ind' bucket: an independent with
no score falls to 'Unk' alongside unmatched voters.
  sd / ld                            -> Dem
  sr / lr                            -> Rep
  nd / U / I  +  score >= 50         -> Dem
  nd / U / I  +  score <  50         -> Rep
  else                               -> Unk

Daily_Absentee_List is deduplicated to one row per voter before use,
matching build_absentee_json.py (most resolved status wins; Deleted and
Not Issued rows excluded). Without it, voters with a replacement ballot
were counted twice in total_vbm and listed twice in cure_voters.

Run:
  python export_cure_dashboard.py
  python export_cure_dashboard.py --election "2026 April 21 Special"
  python export_cure_dashboard.py --out C:/Repos/va-cure/data/cure_data.json
"""

import pyodbc
import json
import os
import argparse
from datetime import datetime

# ---------------------------------------------------------------
# Config
# ---------------------------------------------------------------
SERVER      = r'INSTANCE-1'
DATABASE    = 'Absentee'
DRIVER      = '{ODBC Driver 17 for SQL Server}'
OUTPUT_FILE = r'C:\Scripts\Python\Python_Absentee\dashboards\va-cure\data\cure_data.json'

# 2026 cycle targets the congressional districts, not house districts.
TARGET_CDS = ['01', '02', '05']

# CONG_CODE_VALUE is noisy: a locality that is really one CD picks up a
# handful of stray records in others. Locality x CD rows below this many
# cure records are suppressed so the table is not swamped by them.
MIN_CD_ROWS = 10


# ---------------------------------------------------------------
# Connection helpers
# ---------------------------------------------------------------
def get_conn():
    conn_str = (
        'DRIVER=' + DRIVER + ';'
        'SERVER=' + SERVER + ';'
        'DATABASE=' + DATABASE + ';'
        'Trusted_Connection=yes;'
    )
    return pyodbc.connect(conn_str, autocommit=True)

def rows_to_dicts(cursor):
    cols = [c[0] for c in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]

def safe_pct(num, den):
    try:
        if den and den > 0:
            return round(num / den, 6)
    except Exception:
        pass
    return None


# ---------------------------------------------------------------
# SQL -- Shared dedup CTE
#
# One row per voter, most resolved status wins. Same priority as
# build_absentee_json.py's dal_deduped, so total_vbm here equals
# MailCount on the absentee dashboard. Expects @Election declared.
# ---------------------------------------------------------------
DAL_DEDUPED_CTE = """
dal_deduped AS (
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY IDENTIFICATION_NUMBER
               ORDER BY
                   CASE BALLOT_STATUS
                       WHEN 'Marked'        THEN 1
                       WHEN 'Pre-Processed' THEN 2
                       WHEN 'On Machine'    THEN 3
                       WHEN 'Unmarked'      THEN 4
                       ELSE                      5
                   END,
                   BALLOT_RECEIPT_DATE DESC
           ) AS rn
    FROM dbo.Daily_Absentee_List
    WHERE ELECTION_NAME = @Election
      AND BALLOT_STATUS NOT IN ('Deleted', 'Not Issued')
)"""


# ---------------------------------------------------------------
# SQL -- Statewide
# ---------------------------------------------------------------
STATEWIDE_SQL = """
DECLARE @Election VARCHAR(150) = ?;

WITH""" + DAL_DEDUPED_CTE + """,
cure_party AS (
    SELECT
        ch.IDENTIFICATION_NUMBER,
        ch.Cure_Status,
        ch.Rejection_Reason,
        CASE
            WHEN v.likelyparty IN ('sd','ld')                                   THEN 'Dem'
            WHEN v.likelyparty IN ('sr','lr')                                   THEN 'Rep'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore >= 50   THEN 'Dem'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore <  50   THEN 'Rep'
            ELSE 'Unk'
        END AS Party
    FROM dbo.Cure_History ch
    LEFT JOIN van v ON ch.IDENTIFICATION_NUMBER = v.StateFileID
    WHERE ch.ELECTION_NAME = @Election
),
vbm_party AS (
    SELECT
        d.IDENTIFICATION_NUMBER,
        CASE
            WHEN v.likelyparty IN ('sd','ld')                                   THEN 'Dem'
            WHEN v.likelyparty IN ('sr','lr')                                   THEN 'Rep'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore >= 50   THEN 'Dem'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore <  50   THEN 'Rep'
            ELSE 'Unk'
        END AS Party
    FROM dal_deduped d
    LEFT JOIN van v ON d.IDENTIFICATION_NUMBER = v.StateFileID
    WHERE d.rn = 1
      AND d.BALLOT_STATUS IN ('Marked','Pre-Processed')
)
SELECT
    (SELECT COUNT(*)                                          FROM vbm_party)           AS total_vbm,
    (SELECT COUNT(*) FROM vbm_party   WHERE Party = 'Dem')                             AS vbm_dem,
    (SELECT COUNT(*) FROM vbm_party   WHERE Party = 'Rep')                             AS vbm_rep,
    (SELECT COUNT(*)                                          FROM cure_party)          AS total_in_history,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'Needs Cure')                AS currently_needs_cure,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'Needs Cure' AND Party='Dem') AS currently_needs_cure_dem,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'Needs Cure' AND Party='Rep') AS currently_needs_cure_rep,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'Cured')                     AS cured,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'Cured'      AND Party='Dem') AS cured_dem,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'EVIP')                      AS switched_to_evip,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'Re-Issued')                 AS reissued,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'Undeliverable')             AS undeliverable,
    (SELECT COUNT(*) FROM cure_party  WHERE Cure_Status = 'Other')                     AS other_status,
    (SELECT COUNT(*) FROM cure_party  WHERE Party = 'Dem')                             AS ever_rej_dem,
    (SELECT COUNT(*) FROM cure_party  WHERE Party = 'Rep')                             AS ever_rej_rep;
"""


# ---------------------------------------------------------------
# SQL -- Locality breakdown
# ---------------------------------------------------------------
LOCALITY_SQL = """
DECLARE @Election VARCHAR(150) = ?;

WITH""" + DAL_DEDUPED_CTE + """,
cure_party AS (
    SELECT
        ch.LOCALITY_NAME,
        ch.CONG_CODE_VALUE AS cd,
        ch.Cure_Status,
        CASE
            WHEN v.likelyparty IN ('sd','ld')                                   THEN 'Dem'
            WHEN v.likelyparty IN ('sr','lr')                                   THEN 'Rep'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore >= 50   THEN 'Dem'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore <  50   THEN 'Rep'
            ELSE 'Unk'
        END AS Party
    FROM dbo.Cure_History ch
    LEFT JOIN van v ON ch.IDENTIFICATION_NUMBER = v.StateFileID
    WHERE ch.ELECTION_NAME = @Election
),
vbm_party AS (
    SELECT
        d.LOCALITY_NAME,
        d.CONG_CODE_VALUE AS cd,
        CASE
            WHEN v.likelyparty IN ('sd','ld')                                   THEN 'Dem'
            WHEN v.likelyparty IN ('sr','lr')                                   THEN 'Rep'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore >= 50   THEN 'Dem'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore <  50   THEN 'Rep'
            ELSE 'Unk'
        END AS Party
    FROM dal_deduped d
    LEFT JOIN van v ON d.IDENTIFICATION_NUMBER = v.StateFileID
    WHERE d.rn = 1
      AND d.BALLOT_STATUS IN ('Marked','Pre-Processed')
),
cure_agg AS (
    SELECT
        LOCALITY_NAME,
        cd,
        COUNT(*)                                                            AS total_in_history,
        SUM(CASE WHEN Cure_Status = 'Needs Cure'              THEN 1 ELSE 0 END) AS currently_needs_cure,
        SUM(CASE WHEN Cure_Status = 'Needs Cure' AND Party='Dem' THEN 1 ELSE 0 END) AS currently_needs_cure_dem,
        SUM(CASE WHEN Cure_Status = 'Cured'                   THEN 1 ELSE 0 END) AS cured,
        SUM(CASE WHEN Cure_Status = 'Cured'      AND Party='Dem' THEN 1 ELSE 0 END) AS cured_dem,
        SUM(CASE WHEN Party = 'Dem'               THEN 1 ELSE 0 END)        AS ever_rej_dem
    FROM cure_party
    GROUP BY LOCALITY_NAME, cd
),
vbm_agg AS (
    SELECT
        LOCALITY_NAME,
        cd,
        COUNT(*)                                                            AS total_vbm,
        SUM(CASE WHEN Party = 'Dem' THEN 1 ELSE 0 END)                    AS vbm_dem
    FROM vbm_party
    GROUP BY LOCALITY_NAME, cd
)
SELECT
    ISNULL(va.LOCALITY_NAME, ca.LOCALITY_NAME)  AS locality,
    ISNULL(va.cd, ca.cd)                        AS cd,
    ISNULL(va.total_vbm, 0)                     AS total_vbm,
    ISNULL(va.vbm_dem, 0)                       AS vbm_dem,
    ISNULL(ca.total_in_history, 0)              AS ever_not_accepted,
    ISNULL(ca.ever_rej_dem, 0)                  AS ever_not_accepted_dem,
    ISNULL(ca.currently_needs_cure, 0)          AS currently_needs_cure,
    ISNULL(ca.currently_needs_cure_dem, 0)      AS currently_needs_cure_dem,
    ISNULL(ca.cured, 0)                         AS cured,
    ISNULL(ca.cured_dem, 0)                     AS cured_dem,
    ISNULL(va.total_vbm, 0)
        + ISNULL(ca.currently_needs_cure, 0)    AS attempts
FROM vbm_agg va
FULL OUTER JOIN cure_agg ca
    ON  va.LOCALITY_NAME = ca.LOCALITY_NAME
    AND va.cd            = ca.cd
ORDER BY ISNULL(va.LOCALITY_NAME, ca.LOCALITY_NAME),
         ISNULL(ca.total_in_history, 0) DESC;
"""


# ---------------------------------------------------------------
# SQL -- CD breakdown
# ---------------------------------------------------------------
CD_SQL = """
DECLARE @Election VARCHAR(150) = ?;

WITH cure_party AS (
    SELECT
        ch.CONG_CODE_VALUE      AS cd,
        ch.LOCALITY_NAME,
        ch.Cure_Status,
        ch.Initial_Rejection_Date,
        ch.Latest_Snap_Date,
        CASE
            WHEN v.likelyparty IN ('sd','ld')                                   THEN 'Dem'
            WHEN v.likelyparty IN ('sr','lr')                                   THEN 'Rep'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore >= 50   THEN 'Dem'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore <  50   THEN 'Rep'
            ELSE 'Unk'
        END AS Party
    FROM dbo.Cure_History ch
    LEFT JOIN van v ON ch.IDENTIFICATION_NUMBER = v.StateFileID
    WHERE ch.ELECTION_NAME = @Election
)
SELECT
    ISNULL(cd, 'UNK')                                                   AS cd,
    COUNT(DISTINCT LOCALITY_NAME)                                        AS locality_count,
    COUNT(*)                                                             AS total_in_history,
    SUM(CASE WHEN Cure_Status = 'Needs Cure'               THEN 1 ELSE 0 END) AS currently_needs_cure,
    SUM(CASE WHEN Cure_Status = 'Needs Cure' AND Party='Dem' THEN 1 ELSE 0 END) AS currently_needs_cure_dem,
    SUM(CASE WHEN Cure_Status = 'Cured'                    THEN 1 ELSE 0 END) AS cured,
    SUM(CASE WHEN Party = 'Dem'                            THEN 1 ELSE 0 END) AS ever_rej_dem,
    MIN(Initial_Rejection_Date)                                          AS first_rejection,
    MAX(Latest_Snap_Date)                                                AS latest_snap
FROM cure_party
GROUP BY cd
ORDER BY cd;
"""


# ---------------------------------------------------------------
# SQL -- Rejection reasons
# ---------------------------------------------------------------
REASONS_SQL = """
DECLARE @Election VARCHAR(150) = ?;

WITH cure_party AS (
    SELECT
        ch.Rejection_Reason,
        ch.Cure_Status,
        CASE
            WHEN v.likelyparty IN ('sd','ld')                                   THEN 'Dem'
            WHEN v.likelyparty IN ('sr','lr')                                   THEN 'Rep'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore >= 50   THEN 'Dem'
            WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore <  50   THEN 'Rep'
            ELSE 'Unk'
        END AS Party
    FROM dbo.Cure_History ch
    LEFT JOIN van v ON ch.IDENTIFICATION_NUMBER = v.StateFileID
    WHERE ch.ELECTION_NAME = @Election
)
SELECT
    ISNULL(Rejection_Reason, 'Unknown')                                  AS reason,
    COUNT(*)                                                             AS total,
    SUM(CASE WHEN Party = 'Dem'               THEN 1 ELSE 0 END)        AS dem,
    SUM(CASE WHEN Party = 'Rep'               THEN 1 ELSE 0 END)        AS rep,
    SUM(CASE WHEN Cure_Status = 'Cured'       THEN 1 ELSE 0 END)        AS cured,
    SUM(CASE WHEN Cure_Status = 'Needs Cure'  THEN 1 ELSE 0 END)        AS still_needs_cure
FROM cure_party
GROUP BY Rejection_Reason
ORDER BY total DESC;
"""


# ---------------------------------------------------------------
# SQL -- Status transitions (no party needed, just flow)
# ---------------------------------------------------------------
TRANSITIONS_SQL = """
DECLARE @Election VARCHAR(150) = ?;

SELECT
    ISNULL(Rejection_Reason, 'Unknown')                                  AS rejection_reason,
    SUM(CASE WHEN Cure_Status = 'Needs Cure'    THEN 1 ELSE 0 END)      AS still_needs_cure,
    SUM(CASE WHEN Cure_Status = 'Cured'         THEN 1 ELSE 0 END)      AS now_cured,
    SUM(CASE WHEN Cure_Status = 'EVIP'          THEN 1 ELSE 0 END)      AS switched_evip,
    SUM(CASE WHEN Cure_Status = 'Re-Issued'     THEN 1 ELSE 0 END)      AS reissued,
    SUM(CASE WHEN Cure_Status = 'Undeliverable' THEN 1 ELSE 0 END)      AS undeliverable,
    SUM(CASE WHEN Cure_Status = 'Other'         THEN 1 ELSE 0 END)      AS other_status,
    COUNT(*)                                                             AS total
FROM dbo.Cure_History
WHERE ELECTION_NAME = @Election
GROUP BY Rejection_Reason
ORDER BY total DESC;
"""


# ---------------------------------------------------------------
# SQL -- Cure Voter List (individual records)
#
# Exports ALL Cure_History voters (not just Needs Cure) so that
# the VB Survey export can cover the full cure universe.
# The dashboard Needs Cure tab filters client-side to cure_status = 'Needs Cure'.
# Joins to the deduplicated DAL so each voter appears exactly once.
# ---------------------------------------------------------------
CURE_VOTERS_SQL = """
DECLARE @Election VARCHAR(150) = ?;

WITH""" + DAL_DEDUPED_CTE + """
SELECT
    -- VAN / State IDs (VoterFileVANID is first column for VB bulk load)
    ISNULL(CAST(v.VoterFileVANID AS VARCHAR(20)), '')   AS voter_file_vanid,
    ISNULL(v.StateFileID, '')                           AS state_file_id,

    -- Name
    ISNULL(v.LastName,    '')                           AS last_name,
    ISNULL(v.FirstName,   '')                           AS first_name,
    ISNULL(v.MiddleName,  '')                           AS middle_name,

    -- Contact
    ISNULL(v.PreferredPhone, '')                        AS preferred_phone,
    ISNULL(v.CellPhone,      '')                        AS cell_phone,

    -- Address
    ISNULL(v.Address, '')                               AS address,
    ISNULL(v.City,    '')                               AS city,
    ISNULL(v.State,   '')                               AS state,
    ISNULL(v.Zip5,    '')                               AS zip5,
    ISNULL(v.Zip4,    '')                               AS zip4,

    -- Geography / precinct
    ISNULL(v.CountyName,   '')                          AS county_name,
    ISNULL(v.PrecinctName, '')                          AS precinct_name,

    -- Voter demographics
    ISNULL(CAST(v.Age AS VARCHAR(5)), '')               AS age,
    ISNULL(v.Sex, '')                                   AS sex,
    ISNULL(CONVERT(VARCHAR(10), v.DOB, 120), '')        AS dob,

    -- Cure tracking
    ISNULL(ch.LOCALITY_NAME, '')                        AS locality,
    ISNULL(ch.CONG_CODE_VALUE, '')                      AS cd,
    ISNULL(ch.Cure_Status, '')                          AS cure_status,
    ISNULL(ch.Rejection_Reason, '')                     AS rejection_reason,
    ISNULL(d.BALLOTSTATUSREASON, '')                    AS ballot_status_reason,
    ISNULL(d.Ballot_Comment, '')                        AS ballot_comment,

    -- Derived party
    CASE
        WHEN v.likelyparty IN ('sd','ld')                                   THEN 'Dem'
        WHEN v.likelyparty IN ('sr','lr')                                   THEN 'Rep'
        WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore >= 50   THEN 'Dem'
        WHEN v.likelyparty IN ('nd','U','I') AND v.SupportScore <  50   THEN 'Rep'
        ELSE 'Unk'
    END AS party

FROM dbo.Cure_History ch
LEFT JOIN van v
    ON ch.IDENTIFICATION_NUMBER = v.StateFileID
LEFT JOIN dal_deduped d
    ON  ch.IDENTIFICATION_NUMBER = d.IDENTIFICATION_NUMBER
    AND d.rn = 1
WHERE ch.ELECTION_NAME = @Election
ORDER BY ch.LOCALITY_NAME,
         v.LastName,
         v.FirstName;
"""


# ---------------------------------------------------------------
# Main
# ---------------------------------------------------------------
def main(election_name=None, out_path=None):
    print('Connecting to ' + SERVER + '/' + DATABASE + '...')
    conn = get_conn()
    cur  = conn.cursor()

    # Resolve election name
    if not election_name:
        cur.execute("SELECT MAX(ELECTION_NAME) FROM dbo.Cure_History")
        row = cur.fetchone()
        election_name = row[0] if row and row[0] else 'Unknown'
    print('  Election: ' + election_name)

    # Snapshot date
    cur.execute("SELECT MAX(Latest_Snap_Date) FROM dbo.Cure_History WHERE ELECTION_NAME = ?", election_name)
    row = cur.fetchone()
    snap_date = str(row[0])[:10] if row and row[0] else datetime.now().strftime('%Y-%m-%d')

    # Source record count
    cur.execute("SELECT COUNT(*) FROM dbo.Cure_History WHERE ELECTION_NAME = ?", election_name)
    source_records = cur.fetchone()[0]

    # -- Statewide
    print("  Building statewide...")
    cur.execute(STATEWIDE_SQL, election_name)
    sw = dict(zip([c[0] for c in cur.description], cur.fetchone()))

    total_vbm           = sw.get('total_vbm')               or 0
    vbm_dem             = sw.get('vbm_dem')                  or 0
    vbm_rep             = sw.get('vbm_rep')                  or 0
    ever_rej            = sw.get('total_in_history')         or 0
    curr_needs_cure     = sw.get('currently_needs_cure')     or 0
    curr_needs_cure_dem = sw.get('currently_needs_cure_dem') or 0
    curr_needs_cure_rep = sw.get('currently_needs_cure_rep') or 0
    cured               = sw.get('cured')                    or 0
    cured_dem           = sw.get('cured_dem')                or 0
    evip                = sw.get('switched_to_evip')         or 0
    reissued            = sw.get('reissued')                 or 0
    undeliverable       = sw.get('undeliverable')            or 0
    other_status        = sw.get('other_status')             or 0
    ever_rej_dem        = sw.get('ever_rej_dem')             or 0
    ever_rej_rep        = sw.get('ever_rej_rep')             or 0
    attempts            = total_vbm + curr_needs_cure + undeliverable
    never_cure          = max(0, total_vbm - cured)

    statewide = {
        'total_attempts':               attempts,
        'total_vbm':                    total_vbm,
        'vbm_dem':                      vbm_dem,
        'vbm_rep':                      vbm_rep,
        'dem_pct_vbm':                  safe_pct(vbm_dem, total_vbm),
        'never_needs_cure':             never_cure,
        'ever_not_accepted':            ever_rej,
        'ever_not_accepted_dem':        ever_rej_dem,
        'ever_not_accepted_rep':        ever_rej_rep,
        'ever_not_accepted_dem_pct':    safe_pct(ever_rej_dem, ever_rej),
        'cured_total':                  cured,
        'cured_dem':                    cured_dem,
        'currently_needs_cure':         curr_needs_cure,
        'currently_needs_cure_dem':     curr_needs_cure_dem,
        'currently_needs_cure_rep':     curr_needs_cure_rep,
        'switched_to_evip':             evip,
        'reissued':                     reissued,
        'undeliverable':                undeliverable,
        'other_status':                 other_status,
        'pre_cure_rej_rate':            safe_pct(ever_rej, attempts),
        'post_cure_rej_rate':           safe_pct(curr_needs_cure, attempts),
        'cure_rate':                    safe_pct(cured, ever_rej),
        'dem_pre_cure_rate':            safe_pct(ever_rej_dem, vbm_dem + ever_rej_dem),
        'dem_post_cure_rate':           safe_pct(curr_needs_cure_dem, vbm_dem + ever_rej_dem),
    }

    # -- Locality
    print("  Building locality breakdown...")
    cur.execute(LOCALITY_SQL, election_name)
    loc_rows = rows_to_dicts(cur)
    state_post = statewide['post_cure_rej_rate'] or 0

    by_locality = []
    suppressed = 0
    for r in loc_rows:
        locality = (r.get('locality') or '').strip()
        cd       = str(r.get('cd') or '').zfill(2)
        att      = r.get('attempts')            or 0
        ev       = r.get('ever_not_accepted')   or 0
        cn       = r.get('currently_needs_cure') or 0
        cu       = r.get('cured')               or 0
        post     = safe_pct(cn, att)

        by_locality.append({
            'locality':                 locality,
            'cd':                       cd,
            'target_cd':                cd if cd in TARGET_CDS else 'No',
            'attempts':                 att,
            'total_vbm':                r.get('total_vbm') or 0,
            'dem_pct_vbm':              safe_pct(r.get('vbm_dem'), r.get('total_vbm')),
            'ever_not_accepted':        ev,
            'ever_not_accepted_dem':    r.get('ever_not_accepted_dem') or 0,
            'currently_needs_cure':     cn,
            'currently_needs_cure_dem': r.get('currently_needs_cure_dem') or 0,
            'cured':                    cu,
            'cured_dem':                r.get('cured_dem') or 0,
            'pre_cure_rej_rate':        safe_pct(ev, att),
            'post_cure_rej_rate':       post,
            'cure_rate':                safe_pct(cu, ev),
            'rej_vs_statewide':         round(post - state_post, 6) if post is not None else None,
        })

    # Suppression is about splitting noise, not hiding places. Every locality
    # keeps its dominant row; only ADDITIONAL CD slivers below the floor are
    # dropped, and never one in a target CD.
    by_loc_name = {}
    for row in by_locality:
        by_loc_name.setdefault(row['locality'], []).append(row)

    kept = []
    for name, group in by_loc_name.items():
        group.sort(key=lambda r: (r['ever_not_accepted'], r['total_vbm']), reverse=True)
        for i, row in enumerate(group):
            if i == 0 or row['ever_not_accepted'] >= MIN_CD_ROWS or row['cd'] in TARGET_CDS:
                kept.append(row)
            else:
                suppressed += 1
    kept.sort(key=lambda r: (r['locality'], -r['ever_not_accepted']))
    by_locality = kept

    print(f"    {len(by_locality)} locality x CD rows across "
          f"{len(by_loc_name)} localities "
          f"({suppressed} minor CD rows suppressed under {MIN_CD_ROWS} cure records)")

    # -- CD
    print("  Building CD breakdown...")
    cur.execute(CD_SQL, election_name)
    cd_rows = rows_to_dicts(cur)
    by_cd = []
    for r in cd_rows:
        cd = str(r.get('cd') or '').zfill(2)
        ev = r.get('total_in_history')      or 0
        cn = r.get('currently_needs_cure')  or 0
        cu = r.get('cured')                 or 0
        by_cd.append({
            'cd':                       cd,
            'target':                   cd in TARGET_CDS,
            'localities':               r.get('locality_count'),
            'total_in_history':         ev,
            'ever_rej_dem':             r.get('ever_rej_dem') or 0,
            'currently_needs_cure':     cn,
            'currently_needs_cure_dem': r.get('currently_needs_cure_dem') or 0,
            'cured':                    cu,
            'post_cure_rej_rate':       safe_pct(cn, ev),
            'cure_rate':                safe_pct(cu, ev),
            'first_rejection':          str(r['first_rejection'])[:10] if r.get('first_rejection') else None,
        })

    # -- Rejection reasons
    print("  Building rejection reasons...")
    cur.execute(REASONS_SQL, election_name)
    rejection_reasons = [
        {
            'reason':           r['reason'],
            'total':            r['total'],
            'dem':              r['dem'],
            'rep':              r['rep'],
            'cured':            r['cured'],
            'still_needs_cure': r['still_needs_cure'],
            'cure_rate':        safe_pct(r['cured'], r['total']),
        }
        for r in rows_to_dicts(cur)
    ]

    # -- Status transitions
    print("  Building status transitions...")
    cur.execute(TRANSITIONS_SQL, election_name)
    status_transitions = [
        {
            'rejection_reason': r['rejection_reason'],
            'still_needs_cure': r['still_needs_cure'],
            'now_cured':        r['now_cured'],
            'switched_evip':    r['switched_evip'],
            'reissued':         r['reissued'],
            'undeliverable':    r['undeliverable'],
            'other_status':     r['other_status'],
            'total':            r['total'],
            'cure_rate':        safe_pct(r['now_cured'], r['total']),
        }
        for r in rows_to_dicts(cur)
    ]

    # -- Cure voter list (individual records — used for Needs Cure tab + VB Survey export)
    print("  Building cure voter list...")
    cur.execute(CURE_VOTERS_SQL, election_name)
    cure_voters = []
    for r in rows_to_dicts(cur):
        cure_voters.append({
            # IDs
            'voter_file_vanid':     str(r.get('voter_file_vanid') or '').strip(),
            'state_file_id':        str(r.get('state_file_id')    or '').strip(),
            # Name
            'last_name':            (r.get('last_name')   or '').strip(),
            'first_name':           (r.get('first_name')  or '').strip(),
            'middle_name':          (r.get('middle_name') or '').strip(),
            # Contact
            'preferred_phone':      (r.get('preferred_phone') or '').strip(),
            'cell_phone':           (r.get('cell_phone')      or '').strip(),
            # Address
            'address':              (r.get('address') or '').strip(),
            'city':                 (r.get('city')    or '').strip(),
            'state':                (r.get('state')   or '').strip(),
            'zip5':                 (r.get('zip5')    or '').strip(),
            'zip4':                 (r.get('zip4')    or '').strip(),
            # Geography
            'county_name':          (r.get('county_name')   or '').strip(),
            'precinct_name':        (r.get('precinct_name') or '').strip(),
            # Demographics
            'age':                  (r.get('age') or '').strip(),
            'sex':                  (r.get('sex') or '').strip(),
            'dob':                  (r.get('dob') or '').strip(),
            # Cure tracking
            'locality':             (r.get('locality') or '').strip(),
            'cd':                   str(r.get('cd') or '').zfill(2),
            'cure_status':          (r.get('cure_status')          or '').strip(),
            'rejection_reason':     (r.get('rejection_reason')     or '').strip(),
            'ballot_status_reason': (r.get('ballot_status_reason') or '').strip(),
            'ballot_comment':       (r.get('ballot_comment')       or '').strip(),
            'party':                (r.get('party') or 'Unk').strip(),
        })
    needs_cure_count = sum(1 for v in cure_voters if v['cure_status'] == 'Needs Cure')

    # -- Assemble and write
    export = {
        'meta': {
            'election_name':  election_name,
            'snapshot_date':  snap_date,
            'export_time':    datetime.now().strftime('%Y-%m-%d %H:%M'),
            'source_records': source_records,
            'target_cds':     TARGET_CDS,
            'min_cd_rows':    MIN_CD_ROWS,
            'suppressed_rows': suppressed,
        },
        'statewide':          statewide,
        'by_locality':        by_locality,
        'by_cd':              by_cd,
        'rejection_reasons':  rejection_reasons,
        'status_transitions': status_transitions,
        'cure_voters':        cure_voters,   # new — individual voter records
    }

    out = out_path or OUTPUT_FILE
    os.makedirs(os.path.dirname(out) if os.path.dirname(out) else '.', exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(export, f, indent=2, default=str)

    print('\nExported -> ' + out)
    print('  Election:       ' + election_name)
    print('  Snapshot date:  ' + snap_date)
    print('  Source records: ' + str(source_records))
    print('  Localities:     ' + str(len(by_locality)))
    print('  CDs:            ' + str(len(by_cd)))
    print('  Ever rejected:  ' + str(ever_rej) + '  |  Needs cure: ' + str(curr_needs_cure) + '  |  Cured: ' + str(cured))
    print('  Voter list:     ' + str(len(cure_voters)) + ' total records  |  ' + str(needs_cure_count) + ' currently need cure')
    conn.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Export cure dashboard JSON')
    parser.add_argument('--election', default=None,
                        help='Election name (default: latest in Cure_History)')
    parser.add_argument('--out', default=None,
                        help='Output path (default: data/cure_data.json)')
    args = parser.parse_args()
    main(election_name=args.election, out_path=args.out)
