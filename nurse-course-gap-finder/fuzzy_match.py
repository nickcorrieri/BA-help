#!/usr/bin/env python3
# =====================================================================
#  PHASE 2 - FUZZY COURSE MATCHER   (standard library ONLY)
# =====================================================================
#  No pip installs. Uses only modules that ship with Python:
#      csv, os, sys, datetime, difflib, re
#  So it runs on any stock Python 3 in VS Code with nothing added.
#
#  WHAT IT DOES
#    Phase 1 (the Excel macro) already removed every course the nurse
#    clearly still has (exact Course ID taken within 5 years). What's
#    left in "gaps.csv" might still be covered by a DIFFERENT course
#    with a similar title - e.g. "Hazardous Materials" vs
#    "Taking Care of Hazardous Materials". This script finds those.
#
#    For each gap it looks through that same nurse's taken courses
#    (within 5 years), finds the closest title, scores the similarity,
#    and writes "review.csv" so a human can confirm line by line.
#
#  HOW TO USE (in VS Code)
#    1. From Excel, export two sheets to CSV in this same folder:
#         - the "Gaps" sheet  ->  gaps.csv
#         - the "Taken" sheet ->  taken.csv
#       (File > Save As > CSV, or right-click the sheet tab > Move/Copy
#        to a new book, then Save As CSV.)
#    2. Open this file in VS Code and press the Run button (or in a
#       terminal:  python3 fuzzy_match.py )
#    3. Open review.csv. Each row shows the gap, the closest course she
#       took, a match %, and a prefilled "Needs Course? (yes/no)" guess.
#       Fix any you disagree with.
#    4. To produce the final enrollment list, set MODE = "finalize"
#       below and run again -> writes final.csv (employee + course).
#       (Or do the finalize step back in Excel with BuildFinal.)
# =====================================================================

import csv
import os
import re
import sys
from datetime import date, datetime
from difflib import SequenceMatcher

# ============================ CONFIG ================================
HERE = os.path.dirname(os.path.abspath(__file__))

GAPS_CSV   = os.path.join(HERE, "gaps.csv")     # output of the Phase 1 macro
TAKEN_CSV  = os.path.join(HERE, "taken.csv")    # Table 2, the courses each nurse has taken
REVIEW_CSV = os.path.join(HERE, "review.csv")   # this script writes this
FINAL_CSV  = os.path.join(HERE, "final.csv")    # written only when MODE = "finalize"

MODE = "match"        # "match" = build review.csv ;  "finalize" = build final.csv from review.csv

YEARS_VALID = 5       # only consider courses she took within this many years as possible matches

# ----- OPTIONS: how strict the fuzzy "same topic" match is -----------
#   Similarity runs 0.0 - 1.0. Two knobs set the label + the prefilled guess:
#     score >= STRONG (and recent)  ->  "LIKELY already covered" (prefill: no)
#     WEAK <= score < STRONG        ->  "MAYBE - check it"        (prefill: yes)
#     score <  WEAK                 ->  "no close match"          (prefill: yes)
#   Raise STRONG to send MORE rows to manual review (safer, more work).
#   Lower STRONG to auto-clear more look-alikes (faster, riskier).
STRONG = 0.75         # 0.0 - 1.0
WEAK   = 0.45         # 0.0 - 1.0
# --------------------------------------------------------------------

# gaps.csv is produced by the Phase-1 macro, so its headers are fixed/known:
G_EMP, G_COURSE, G_TITLE = "Employee ID", "Required Course ID", "Required Course Title"

# taken.csv is YOUR Table 2 export. If USE_COLUMN_LETTERS is True, the T_*
# values below are COLUMN LETTERS (A, B, C...) and taken.csv's header row is
# ignored. If False, they're header names matched against the first row.
USE_COLUMN_LETTERS = True
T_EMP, T_COURSE, T_TITLE, T_DATE = "A", "B", "C", "D"   # verify against your sheet

# ASK-vs-DON'T-ASK: True = the script asks you to type the Taken column letters
# when it runs (defaults shown in [brackets], press Enter to accept). False =
# just use the T_* letters above.
ASK_AT_RUNTIME = False

# Generic words stripped before comparing titles, so the score reflects
# the actual topic. Add your own org's filler words here.
STOPWORDS = {
    "the", "of", "a", "an", "to", "for", "in", "and", "on", "with", "your",
    "course", "training", "module", "intro", "introduction", "basic", "basics",
    "annual", "required", "mandatory", "online", "elearning", "part",
}
# ====================================================================


def norm_tokens(text):
    """Lowercase, drop punctuation and filler words, return the topic words."""
    text = (text or "").lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return [w for w in text.split() if w and w not in STOPWORDS]


def similarity(a, b):
    """0.0 - 1.0 topic similarity between two course titles.

    Combines three views and takes the strongest, because each catches a
    different shape of "same thing, named differently":
      - sequence ratio : handles typos / minor wording changes
      - jaccard        : overall word overlap
      - overlap coeff. : high when one title's words are a subset of the
                         other ("hazardous materials" inside "taking care
                         of hazardous materials")
    """
    ta, tb = norm_tokens(a), norm_tokens(b)
    if not ta or not tb:
        return 0.0
    sa, sb = " ".join(sorted(ta)), " ".join(sorted(tb))
    seq = SequenceMatcher(None, sa, sb).ratio()
    set_a, set_b = set(ta), set(tb)
    inter = len(set_a & set_b)
    union = len(set_a | set_b)
    jacc = inter / union if union else 0.0
    overlap = inter / min(len(set_a), len(set_b))     # subset-friendly
    return max(seq, jacc, 0.9 * overlap)


def parse_date(s):
    """Best-effort date parse. Returns a date or None."""
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y",
                "%m-%d-%Y", "%Y/%m/%d", "%b %d, %Y", "%B %d, %Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # last resort: ISO-ish leading 10 chars (e.g. "2023-04-01 00:00:00")
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def cutoff_date():
    today = date.today()
    try:
        return today.replace(year=today.year - YEARS_VALID)
    except ValueError:           # Feb 29 -> Feb 28
        return today.replace(year=today.year - YEARS_VALID, day=28)


def read_csv(path):
    if not os.path.exists(path):
        sys.exit(f"ERROR: can't find {path}\n"
                 f"Export the sheet to CSV into this folder first.")
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def need_col(rows, col, path):
    if rows and col not in rows[0]:
        sys.exit(f"ERROR: column '{col}' not found in {os.path.basename(path)}.\n"
                 f"Found columns: {list(rows[0].keys())}\n"
                 f"Fix the header names in the CONFIG section of this script.")


def letter_to_index(s):
    s = s.strip().upper()
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1   # "A" -> 0


def ask_taken_letters():
    """If ASK_AT_RUNTIME, prompt for the four Taken column letters; else defaults."""
    if not (ASK_AT_RUNTIME and USE_COLUMN_LETTERS):
        return T_EMP, T_COURSE, T_TITLE, T_DATE
    print("Enter the column LETTER for each Taken field (press Enter to keep the default):")

    def ask(label, default):
        try:
            v = input(f"  {label} [{default}]: ").strip()
        except EOFError:
            v = ""
        return v or default
    return (ask("Employee ID", T_EMP), ask("Course id", T_COURSE),
            ask("Course title", T_TITLE), ask("Date completed", T_DATE))


def load_taken(t_emp, t_course, t_title, t_date):
    """Return Table 2 as normalized dicts: emp, course, title, date.
    Honors USE_COLUMN_LETTERS (letters = positions, header row skipped)."""
    if not os.path.exists(TAKEN_CSV):
        sys.exit(f"ERROR: can't find {TAKEN_CSV}\nExport the Taken sheet to CSV here first.")
    with open(TAKEN_CSV, newline="", encoding="utf-8-sig") as f:
        if USE_COLUMN_LETTERS:
            rows = list(csv.reader(f))[1:]  # skip header row
            ie, ic, it, idt = (letter_to_index(x) for x in (t_emp, t_course, t_title, t_date))
            cell = lambda r, i: r[i] if 0 <= i < len(r) else ""
            return [{"emp": cell(r, ie), "course": cell(r, ic),
                     "title": cell(r, it), "date": cell(r, idt)} for r in rows]
        recs = list(csv.DictReader(f))
        for c in (t_emp, t_course, t_title):
            if recs and c not in recs[0]:
                sys.exit(f"ERROR: column '{c}' not in taken.csv. Found: {list(recs[0].keys())}")
        return [{"emp": r.get(t_emp, ""), "course": r.get(t_course, ""),
                 "title": r.get(t_title, ""), "date": r.get(t_date, "")} for r in recs]


# ----------------------------- MATCH --------------------------------
def run_match():
    gaps = read_csv(GAPS_CSV)
    for c in (G_EMP, G_COURSE, G_TITLE):
        need_col(gaps, c, GAPS_CSV)
    taken = load_taken(*ask_taken_letters())

    cutoff = cutoff_date()

    # group each nurse's recent taken courses
    by_emp = {}
    for r in taken:
        emp = (r["emp"] or "").strip().upper()
        if not emp:
            continue
        d = parse_date(r["date"])
        recent = (d is not None and d >= cutoff)
        by_emp.setdefault(emp, []).append({
            "title": r["title"],
            "course": r["course"],
            "date": d,
            "recent": recent,
        })

    out = []
    for g in gaps:
        emp = (g.get(G_EMP) or "").strip().upper()
        req_title = g.get(G_TITLE, "")
        candidates = by_emp.get(emp, [])

        # prefer only courses taken within 5 years; if none, fall back to all
        recent = [c for c in candidates if c["recent"]]
        pool = recent if recent else candidates

        best, best_score = None, 0.0
        for c in pool:
            s = similarity(req_title, c["title"])
            if s > best_score:
                best, best_score = c, s

        if best is None:
            suggestion, needs = "NO PRIOR COURSES - needs it", "yes"
            bt, bc, bd, expired = "", "", "", ""
        else:
            expired = "" if best["recent"] else "  (match is OLD/expired)"
            bt, bc = best["title"], best["course"]
            bd = best["date"].isoformat() if best["date"] else ""
            if best_score >= STRONG and best["recent"]:
                suggestion, needs = "LIKELY same topic - review", "no"
            elif best_score >= WEAK:
                suggestion, needs = "MAYBE - check it" + expired, "yes"
            else:
                suggestion, needs = "no close match - needs it", "yes"

        out.append({
            "Employee ID": g.get(G_EMP, ""),
            "New Job Code": g.get("New Job Code", ""),
            "Required Course ID": g.get(G_COURSE, ""),
            "Required Course Title": req_title,
            "Best Match Taken Title": bt,
            "Best Match Course ID": bc,
            "Similarity %": round(best_score * 100),
            "Taken Date": bd,
            "Suggestion": suggestion,
            "Needs Course? (yes/no)": needs,
        })

    fields = ["Employee ID", "New Job Code", "Required Course ID",
              "Required Course Title", "Best Match Taken Title",
              "Best Match Course ID", "Similarity %", "Taken Date",
              "Suggestion", "Needs Course? (yes/no)"]
    with open(REVIEW_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(out)

    likely = sum(1 for r in out if r["Needs Course? (yes/no)"] == "no")
    print(f"Wrote {len(out)} rows to {REVIEW_CSV}")
    print(f"  {likely} look already-covered by a similar course (prefilled 'no').")
    print(f"  {len(out) - likely} look like real gaps (prefilled 'yes').")
    print("Open review.csv, correct the 'Needs Course?' column, then set "
          "MODE='finalize' and run again (or use BuildFinal in Excel).")


# ---------------------------- FINALIZE ------------------------------
def run_finalize():
    rows = read_csv(REVIEW_CSV)
    need_col(rows, "Needs Course? (yes/no)", REVIEW_CSV)
    keep = [r for r in rows
            if (r.get("Needs Course? (yes/no)") or "").strip().lower() == "yes"]
    with open(FINAL_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["Employee ID", "Course ID", "Course Title"])
        w.writeheader()
        for r in keep:
            w.writerow({
                "Employee ID": r.get("Employee ID", ""),
                "Course ID": r.get("Required Course ID", ""),
                "Course Title": r.get("Required Course Title", ""),
            })
    print(f"Wrote {len(keep)} enrollment rows to {FINAL_CSV}")


if __name__ == "__main__":
    if MODE == "finalize":
        run_finalize()
    else:
        run_match()
