# Course Gap Finder for Transferring Nurses

Figures out, for each transferring nurse, which courses her **new** job
requires that she hasn't already taken (within the last 5 years) — including
"same topic, different course number" matches.

Everything here runs with **no installs**. Pick the tool your approved system
actually allows (see next section).

## Which tool? ("macro" is two different things in Excel)

Excel has **two** kinds of automation, and they're governed differently:

- **Office Scripts** — the **Automate** tab → *New Script*. TypeScript, runs in
  Microsoft's **cloud sandbox**, governed by your M365 admin, **no local file
  access, no "enable macros" prompt**. In locked-down / PII environments this is
  usually the **approved** one. ➜ **`GapFinder.osts.ts`** — recommended.
- **VBA** — the `Alt+F11` editor. Classic local macros. Powerful, but often
  **disabled by security policy** because it can touch the local machine.
  ➜ `01_GapFinder.bas` + `02_FuzzyFallback.bas`.

If she can open the **Automate** tab and create a script, use the Office Script —
it does the whole job inside the workbook with nothing else needed.

## The files

| File | What it is | Needs |
|------|-----------|-------|
| **`GapFinder.osts.ts`** | **Recommended.** Office Script — does Phase 1 **and** Phase 2 in one, all inside the workbook. | Excel Automate tab |
| `01_GapFinder.bas` | VBA, **Phase 1** (exact match + 5-yr date check). Has `BuildFinal`. | Excel + VBA enabled |
| `02_FuzzyFallback.bas` | VBA, **Phase 2** fuzzy match. | Excel + VBA enabled |
| `fuzzy_match.py` | **Phase 2** in Python (stdlib only, no pip). | Python / VS Code |
| `_proof/` | A test harness that runs the real Office Script logic to prove it works. *(Not for end use.)* | — |

All three approaches use the **same algorithm and the same WEAK/STRONG options**,
so they produce equivalent results — choose by what's allowed, not by quality.

## Set up your workbook (one time)

Put your data on sheets named exactly:

- **`Pairings`** (Table 1) — one row per job-code + course it requires.
  Headers: `Job Code`, `Course ID`, and ideally `Course Title`.
- **`Taken`** (Table 2) — every course the transferring nurses have taken.
  Headers: `Employee ID`, `Course ID`, `Course Title`, `Date Completed`.
  *(Format the date column as a Date so the 5-year check reads it correctly.)*
- **`Roster`** — the 12 nurses. Headers: `Employee ID`, `New Job Code`.
- *(optional)* **`Catalog`** — `Course ID`, `Course Title` — only needed if
  `Pairings` has no title column, so the fuzzy step still has titles to match.

If your real column names differ, edit the `CONFIG` block at the top of whichever
file you're using. IDs are matched trimmed + uppercased, so minor spacing/case
differences won't break it.

---

## Run it — Option A: Office Script (recommended)

1. In Excel: **Automate** tab → **New Script**. Delete the sample code.
2. Open `GapFinder.osts.ts`, copy **everything**, paste it into the script
   editor. Click **Save**, then **Run**.
3. Read the new **`Gaps`** and **`Review`** sheets it creates.
4. In **`Review`**, set the **`Needs Course? (yes/no)`** column for each row
   (it comes prefilled with a best-guess — override anything you disagree with).
5. Near the top of the script, change `MODE` from `"review"` to `"finalize"`,
   **Save**, **Run** again → it writes the **`Final`** sheet
   (`Employee ID | Course ID | Course Title`).

That's the whole thing — no terminal, no CSV exports.

## Run it — Option B: VBA (if Office Scripts aren't available)

1. `Alt+F11` → **File ▸ Import File…** → `01_GapFinder.bas`. `Alt+F8` →
   **BuildGaps** → **Run**. Read the **`Gaps`** sheet.
2. `Alt+F11` → import `02_FuzzyFallback.bas`. `Alt+F8` → **BuildReviewFuzzy** →
   **Run**. Read the **`Review`** sheet, set the `Needs Course?` column.
3. `Alt+F8` → **BuildFinal** → **Run** → produces the **`Final`** sheet.
   **File ▸ Save As ▸ CSV** to export.

## Run it — Option C: Python (if a terminal/VS Code is allowed)

1. From Excel, save two sheets as CSV **into this folder**:
   `Gaps` → `gaps.csv`, `Taken` → `taken.csv`. (Run **BuildGaps** in VBA first,
   or just export your raw Table 1/2 and adjust the headers in the script.)
2. Open `fuzzy_match.py` in VS Code → **Run** (or `python3 fuzzy_match.py`) →
   writes `review.csv`. Edit the `Needs Course?` column.
3. Set `MODE = "finalize"` near the top, run again → writes `final.csv`.

---

## How the matching works

- **Phase 1 (no judgment):** a required course counts as **done** only if the
  nurse took the **exact same Course ID** on/after the 5-year cutoff. Everything
  else is a gap, labelled `MISSING` (never taken) or `EXPIRED` (taken too long
  ago / no date).
- **Phase 2 (fuzzy):** each gap's title is compared to what she actually took,
  scored 0–100%. Close-but-not-exact look-alikes ("Hazardous Materials Handling"
  vs "Taking Care of Hazardous Materials") surface as **MAYBE – check it** so a
  human confirms them.
- **Result:** after you set `Needs Course?`, the finalize step keeps the `yes`
  rows → the enrollment list.

## Tuning the OPTIONS

Same knobs in every tool (top of each file):

- **`YEARS_VALID`** — the 5-year rule (relative to today).
- **`STRONG`** (default `0.75`) — at/above this *and* recent, a match is labelled
  "LIKELY already covered" and prefilled `no`. **Raise it** to push more rows
  into manual review (safer); **lower it** to auto-clear more look-alikes.
- **`WEAK`** (default `0.45`) — below this, "no close match" / real gap.
  Between WEAK and STRONG = "MAYBE – check it".
- **`STOPWORDS`** — filler words ignored in titles (e.g. "annual", "training").
  Add your org's boilerplate so the score reflects the real topic.

## Proof it works

`_proof/run_proof.mjs` loads the **actual** `GapFinder.osts.ts`, feeds it sample
data through a mock workbook, and checks the results (exact match satisfied,
expired flagged, fuzzy match found, finalize honors edits). Run it with:

```
node _proof/run_proof.mjs
```

All checks pass, and the Python tool produces the equivalent ranking on the same
data.

## Notes

- Python uses `difflib`; the Office Script / VBA use a Levenshtein + word-overlap
  score. Exact percentages differ by a few points, but both rank the same
  obvious match to the top and flag borderline ones for review — and a human
  confirms every row, so the small difference doesn't matter.
- Nothing here deletes or overwrites your source data; output always goes to new
  sheets (`Gaps`, `Review`, `Final`) or new CSV files.
