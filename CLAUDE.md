# Claude / Agent Instructions — BA Help Repo

## PII Scrubbing Before Any Commit

This repo is used by Minnesota healthcare employees. Sample and demo files frequently contain real-world references that must be abstracted **before** any file is committed or staged to git.

### What to scan for and replace

**Minnesota counties** — any county name that is or looks like a real MN county:
Hennepin, Ramsey, Anoka, Dakota, Washington, Scott, Carver, Wright, Sherburne, Stearns, Olmsted, St. Louis, Polk, Clay, etc. Minneapolis, Saint Paul, Duluth, Rochester and similar city names fall in this category too.
→ Replace with: `Metro County`, `North County`, `Region A`, etc.

**Hospital / health system names** — any name that reads like a real facility:
e.g. "Hennepin Healthcare", "M Health Fairview", "Allina", "Children's Minnesota", any "__ Medical Center / __ Hospital / __ Clinic"
→ Replace with: `General Hospital`, `Metro Medical Center`, `Facility A`, etc.

**Real person names** — any name that is not an obvious placeholder:
Acceptable placeholders: Jane Doe, John Doe, Joe Somebody, Employee 1001, Nurse A, etc.
NOT acceptable in committed files: first+last combos that look real, names pulled from actual rosters, manager names, etc.
→ Replace with: `Jane Doe`, `John Doe`, or sequential IDs (`EMP-001`, `EMP-002`).

**Employee / badge IDs** — real numeric IDs from an LMS or HR system
→ Replace with: `10001`, `10002`, `10003`, … (obviously synthetic sequences)

**Job codes or department codes** — only abstract if they appear alongside other identifiers that would make the combination re-identifying. Generic codes like `4150`, `299` used in isolation as demo values are fine.

### When this rule applies

- Any file being committed as a demo, sample, test, or example (CSV, JSON, text, etc.)
- Any inline sample data inside a script or proof harness
- File names themselves (e.g. `hennepin_roster.csv` → `sample_roster.csv`)

It does **not** apply to:
- Code logic, variable names, or comments that reference the domain conceptually ("MN county", "hospital dept")
- Files the user has explicitly marked as already-scrubbed

### Practical workflow

1. Before `git add` on any data file, scan each column/field for the patterns above.
2. If a match is found, abstract it in-place and note what was changed in the commit message (e.g. "scrub: replaced real county name with Metro County in sample data").
3. If uncertain whether a value is real or synthetic, treat it as real and abstract it.
4. Never commit a file with a real name, facility, or county when a placeholder works equally well for demonstration purposes.

Minneapolis and Hennepin County references are **especially common** in this dataset — always check for them specifically.
