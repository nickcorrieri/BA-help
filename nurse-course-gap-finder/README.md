# Course Gap Finder — Nurse Transfer Tool

For each transferring nurse: which courses does her **new** job require that she hasn't already taken within the last 5 years?

No installs. Pick the tool your approved system allows.

---

## Quick start

Three sheets in your workbook, named exactly **`Pairings`**, **`Taken`**, **`Roster`**. Then pick a tool:

| Tool | How to run | Best for |
|------|-----------|---------|
| **`GapFinder.osts.ts`** ← start here | Excel → Automate tab → New Script → paste → Run | M365 / restricted PII systems (no "enable macros" prompt) |
| `01_GapFinder.bas` + `02_FuzzyFallback.bas` | Alt+F11 → Import → Alt+F8 → Run | When the Automate tab is unavailable |
| `fuzzy_match.py` | `python3 fuzzy_match.py` | VS Code / terminal |

---

## Column setup

**Pairings** letters are already set from your export format:

| Value | Column |
|-------|--------|
| Course ID (`Course-Ext-ID`) | **A** |
| Course Title (`Full-Course-Name`) | **B** |
| COMBO (`Dept-Code + Job-Code`) | **H** |
| State (optional) | **L** |

For **Taken** and **Roster**, edit the `Settings` sheet (auto-created on first run) or the CONFIG block at the top of each file. Defaults are `A/B/C/D` and `A/B` — verify against your actual sheet.

### COMBO format

DEPT-CODE + JOB-CODE glued with **no separator**: `4150` + `299` = **`4150299`**.

**If your Roster has them in separate columns** (common after an LMS export), set `Roster - DEPT-CODE col` in the Settings sheet to your dept column letter. The tool glues them automatically — no CONCAT formula needed in the sheet.

---

## Workflow

1. **Run** → creates `Gaps` and `Review` sheets (or CSVs for Python)
2. **Review** → each fuzzy-match row has a `Needs Course? (yes/no)` column prefilled with a best guess — override anything wrong
3. **Finalize** → change `MODE` to `"finalize"` at the top of the file, run again → `Final` sheet with the enrollment list

---

## Tuning knobs (same in every tool)

| Setting | Default | Effect |
|---------|---------|--------|
| `YEARS_VALID` | `5` | How old a completion can be to still count |
| `STRONG` | `0.75` | Above this → prefilled "no" (likely covered). Raise for more manual review. |
| `WEAK` | `0.45` | Below this → definitely a gap |
| `STOPWORDS` | see file | Filler words stripped before title comparison |

---

## Verify it works

```
node _proof/run_proof.mjs
```

Runs the actual Office Script logic against sample data. 13 checks pass. No extra installs needed.

---

## Office Script vs VBA — which "macro" is which?

- **Office Scripts** = the **Automate** tab in Excel. TypeScript, runs in Microsoft's cloud sandbox, governed by M365 admin. No local file access, no "enable macros" prompt. **Usually the approved path for PII.**
- **VBA** = `Alt+F11`. Classic local macros. Powerful, but often disabled by security policy.

If the Automate tab exists, use `GapFinder.osts.ts`. It does the whole job — Phase 1 and Phase 2 — in one script, inside the workbook.
