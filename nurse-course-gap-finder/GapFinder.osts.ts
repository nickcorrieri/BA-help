/* =====================================================================
 *  COURSE GAP FINDER  --  OFFICE SCRIPT  (the "approved systems" version)
 * =====================================================================
 *  This is an Office Script (Excel's Automate tab > New Script), NOT VBA.
 *  It runs in Microsoft's cloud sandbox, governed by your M365 admin, with
 *  no local file access and no "enable macros" prompt -- which is usually
 *  the path that's allowed when working with PII.
 *
 *  It does the WHOLE job in one place, entirely inside the workbook:
 *    Phase 1  exact Course-ID match + 5-year date check  -> "Gaps" sheet
 *    Phase 2  fuzzy title match (same topic, diff number) -> "Review" sheet
 *  No terminal, no Python, no CSV exporting.
 *
 *  HOW TO RUN
 *    1. In Excel: Automate tab > New Script. Delete the sample code.
 *    2. Paste ALL of this file in. Click Save, then Run.
 *    3. Read the "Gaps" and "Review" sheets it creates.
 *    4. In "Review", set the "Needs Course? (yes/no)" column for each row.
 *    5. Change MODE below from "review" to "finalize", Save, Run again ->
 *       it writes the "Final" sheet (Employee ID + Course needed).
 *
 *  SET UP YOUR DATA (sheet names + headers) -- edit CONFIG if yours differ.
 * =====================================================================*/

// ============================ CONFIG ================================
const MODE: "review" | "finalize" = "review";   // flip to "finalize" for the last step

// If true, the HDR values below are COLUMN LETTERS (A, B, C...) and the header
// row is ignored. If false, they are header NAMES matched against row 1.
// Column letters are the easy path: just read them off the top of your sheet.
const USE_COLUMN_LETTERS = true;

// ASK-vs-DON'T-ASK: an Office Script can't pop a prompt, so "ask" = read the
// column letters from a "Settings" sheet you fill in. On first run the script
// creates that sheet prefilled with the defaults below; edit it and re-run.
//   true  = use the Settings sheet (auto-created)   <- "ask"
//   false = just use the HDR letters below          <- "don't ask"
const ASK_VIA_SETTINGS = true;

const SHEET = {
  settings: "Settings",
  pairings: "Pairings",   // Table 1: Job Code + Course ID it requires
  taken: "Taken",         // Table 2: courses each nurse has taken
  roster: "Roster",       // the 12 nurses + their new job code
  catalog: "Catalog",     // OPTIONAL Course ID -> Title (used only if Pairings has no title)
  gaps: "Gaps",           // output
  review: "Review",       // output
  final: "Final",         // output (finalize mode)
};

// With USE_COLUMN_LETTERS = true these are COLUMN LETTERS. (If you flip that to
// false, put the header NAMES here instead -- the names are in the comments.)
const HDR = {
  // ----- Table 1 (Pairings) -- prefilled from the column order you listed -----
  pairJob: "H",        // Combo (Dept-code + Job-code) -- the job key
  pairCourse: "A",     // Course-ext-id
  pairTitle: "B",      // Full-course-name   (set to "" if there's no title column)
  pairState: "L",      // State              (optional; only used if ACTIVE_STATE set)
  // ----- Table 2 (Taken) -- VERIFY these letters against your sheet -----
  takenEmp: "A",       // Employee ID
  takenCourse: "B",    // course id (whatever matches Course-ext-id from Table 1)
  takenTitle: "C",     // course title
  takenDate: "D",      // completion date
  // ----- Roster -- VERIFY. Two options:
  //   A) Pre-glued:  leave rosterDept blank; rosterNewJob holds the full COMBO (e.g. "4150299")
  //   B) Separate:   set rosterDept to the DEPT-CODE col and rosterNewJob to the JOB-CODE col;
  //                  the script concatenates them for you so no formula is needed in the sheet.
  rosterEmp: "A",      // Employee ID
  rosterDept: "",      // DEPT-CODE col -- blank = COMBO is already pre-glued in rosterNewJob
  rosterNewJob: "B",   // full COMBO (if rosterDept blank) OR JOB-CODE col (if rosterDept set)
  // ----- Catalog (optional, only used if pairTitle == "") -----
  catCourse: "A",
  catTitle: "B",
};

const YEARS_VALID = 5;          // a past course only counts if taken within this many years

const ACTIVE_STATE = "";        // "" = use ALL pairing rows. Set e.g. "Active" to only count
                                //   pairing rows whose State column equals this value.

// ----- OPTIONS: how strict the fuzzy "same topic" match is ----------
//   Similarity runs 0..100%. Two knobs decide the label + the prefilled guess:
//     score >= STRONG (and recent)  ->  "LIKELY already covered"  (prefill: no)
//     WEAK <= score < STRONG        ->  "MAYBE - check it"        (prefill: yes)
//     score <  WEAK                 ->  "no close match"          (prefill: yes)
//   Raise STRONG to send MORE rows to manual review (safer, more work).
//   Lower STRONG to auto-clear more look-alikes (faster, riskier).
const STRONG = 0.75;            // 0.00 - 1.00
const WEAK = 0.45;              // 0.00 - 1.00
// --------------------------------------------------------------------

// filler words ignored when comparing titles, so the score reflects the topic
const STOPWORDS = new Set([
  "the", "of", "a", "an", "to", "for", "in", "and", "on", "with", "your",
  "course", "training", "module", "intro", "introduction", "basic", "basics",
  "annual", "required", "mandatory", "online", "elearning", "part",
]);
// ====================================================================


function main(workbook: ExcelScript.Workbook) {
  if (MODE === "finalize") {
    buildFinal(workbook);
  } else {
    buildReview(workbook);
  }
}

// ----------------------- main flows ---------------------------------
function buildReview(workbook: ExcelScript.Workbook) {
  const C = resolveCols(workbook);   // letters from the Settings sheet, or HDR defaults

  const pairings = readSheet(workbook, SHEET.pairings, true);
  const taken = readSheet(workbook, SHEET.taken, true);
  const roster = readSheet(workbook, SHEET.roster, true);
  const catalog = readSheet(workbook, SHEET.catalog, false); // optional

  const pJob = col(pairings, C.pairJob), pCourse = col(pairings, C.pairCourse);
  const pTitle = C.pairTitle ? col(pairings, C.pairTitle) : -1;
  const pState = C.pairState ? col(pairings, C.pairState) : -1;
  const tEmp = col(taken, C.takenEmp), tCourse = col(taken, C.takenCourse);
  const tTitle = col(taken, C.takenTitle), tDate = col(taken, C.takenDate);
  const rEmp = col(roster, C.rosterEmp), rNew = col(roster, C.rosterNewJob);
  const rDept = C.rosterDept ? col(roster, C.rosterDept) : -1; // -1 = pre-glued mode

  const cutoff = cutoffNum();

  // optional catalog: Course ID -> Title
  const catTitle = new Map<string, string>();
  if (catalog && C.pairTitle === "") {
    const cC = col(catalog, HDR.catCourse), cT = col(catalog, HDR.catTitle);
    for (const row of catalog.rows) catTitle.set(key(row[cC]), str(row[cT]));
  }

  // taken: emp|course -> latest date ; and per-emp list for fuzzy
  const takenLatest = new Map<string, number>();
  const byEmp = new Map<string, { title: string; course: string; date: number; recent: boolean }[]>();
  for (const row of taken.rows) {
    const emp = key(row[tEmp]), course = key(row[tCourse]);
    if (!emp || !course) continue;
    const d = parseDate(str(row[tDate]));
    const k = emp + "|" + course;
    if (d !== null) {
      const prev = takenLatest.get(k);
      if (prev === undefined || d > prev) takenLatest.set(k, d);
    } else if (!takenLatest.has(k)) {
      takenLatest.set(k, 0); // taken but unknown date
    }
    if (!byEmp.has(emp)) byEmp.set(emp, []);
    byEmp.get(emp)!.push({
      title: str(row[tTitle]), course: str(row[tCourse]),
      date: d === null ? 0 : d, recent: d !== null && d >= cutoff,
    });
  }

  // job -> [{course,title}]
  const jobCourses = new Map<string, { course: string; title: string }[]>();
  for (const row of pairings.rows) {
    if (ACTIVE_STATE && pState >= 0 &&
        str(row[pState]).trim().toLowerCase() !== ACTIVE_STATE.toLowerCase()) continue;
    const job = key(row[pJob]), course = key(row[pCourse]);
    if (!job || !course) continue;
    let title = pTitle >= 0 ? str(row[pTitle]) : "";
    if (!title && catTitle.has(course)) title = catTitle.get(course)!;
    if (!jobCourses.has(job)) jobCourses.set(job, []);
    jobCourses.get(job)!.push({ course, title });
  }

  const gaps: string[][] = [];
  const review: (string | number)[][] = [];

  for (const row of roster.rows) {
    const emp = key(row[rEmp]);
    const newJob = rDept >= 0 ? key(row[rDept]) + key(row[rNew]) : key(row[rNew]);
    if (!emp) continue;
    const required = jobCourses.get(newJob);
    if (!required) continue; // no requirements found for this job code

    for (const req of required) {
      const k = emp + "|" + req.course;
      let status = "";
      if (takenLatest.has(k)) {
        if (takenLatest.get(k)! >= cutoff) status = "";                 // satisfied
        else status = "EXPIRED (>" + YEARS_VALID + "yr or no date)";
      } else {
        status = "MISSING (never taken)";
      }
      if (status === "") continue; // not a gap

      gaps.push([emp, newJob, req.course, req.title, status]);

      // fuzzy: best match among her recent taken courses (fall back to all)
      const all = byEmp.get(emp) || [];
      const recent = all.filter((c) => c.recent);
      const pool = recent.length ? recent : all;
      let best: { title: string; course: string; date: number; recent: boolean } | null = null;
      let bestScore = 0;
      for (const c of pool) {
        const s = similarity(req.title, c.title);
        if (s > bestScore) { bestScore = s; best = c; }
      }

      let suggestion: string, needs: string, bt = "", bc = "", bd = "";
      if (best === null) {
        suggestion = "NO PRIOR COURSES - needs it"; needs = "yes";
      } else {
        bt = best.title; bc = best.course;
        bd = best.date ? numToIso(best.date) : "";
        const oldNote = best.recent ? "" : "  (match is OLD/expired)";
        if (bestScore >= STRONG && best.recent) { suggestion = "LIKELY same topic - review"; needs = "no"; }
        else if (bestScore >= WEAK) { suggestion = "MAYBE - check it" + oldNote; needs = "yes"; }
        else { suggestion = "no close match - needs it"; needs = "yes"; }
      }

      review.push([emp, newJob, req.course, req.title, bt, bc,
        Math.round(bestScore * 100), bd, suggestion, needs]);
    }
  }

  writeSheet(workbook, SHEET.gaps,
    ["Employee ID", "New Job Code", "Required Course ID", "Required Course Title", "Status"],
    gaps);
  writeSheet(workbook, SHEET.review,
    ["Employee ID", "New Job Code", "Required Course ID", "Required Course Title",
     "Best Match Taken Title", "Best Match Course ID", "Similarity %", "Taken Date",
     "Suggestion", "Needs Course? (yes/no)"],
    review);

  console.log(`Gaps: ${gaps.length} row(s). Review: ${review.length} row(s).`);
  console.log(`Edit the 'Needs Course?' column in Review, then set MODE='finalize' and run again.`);
}

function buildFinal(workbook: ExcelScript.Workbook) {
  const r = readSheet(workbook, SHEET.review, true);
  const cEmp = colByHeader(r, "Employee ID");
  const cCourse = colByHeader(r, "Required Course ID");
  const cTitle = colByHeader(r, "Required Course Title");
  const cNeeds = colByHeader(r, "Needs Course? (yes/no)");

  const out: string[][] = [];
  for (const row of r.rows) {
    const keep = cNeeds < 0 ? true : str(row[cNeeds]).trim().toLowerCase() === "yes";
    if (keep && key(row[cEmp]) && key(row[cCourse])) {
      out.push([str(row[cEmp]), str(row[cCourse]), cTitle >= 0 ? str(row[cTitle]) : ""]);
    }
  }
  writeSheet(workbook, SHEET.final, ["Employee ID", "Course ID", "Course Title"], out);
  console.log(`Final: ${out.length} enrollment row(s) written to '${SHEET.final}'.`);
}

// ----------------------- fuzzy scoring ------------------------------
function normTokens(text: string): string[] {
  return (text || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").split(" ")
    .filter((w) => w && !STOPWORDS.has(w));
}

function similarity(a: string, b: string): number {
  const ta = normTokens(a), tb = normTokens(b);
  if (!ta.length || !tb.length) return 0;
  const sa = new Set(ta), sb = new Set(tb);
  let inter = 0;
  for (const w of sa) if (sb.has(w)) inter++;
  const union = sa.size + sb.size - inter;
  const jacc = union ? inter / union : 0;
  const overlap = inter / Math.min(sa.size, sb.size);          // subset-friendly
  const lev = levRatio([...ta].sort().join(" "), [...tb].sort().join(" "));
  return Math.max(jacc, lev, 0.9 * overlap);
}

function levRatio(a: string, b: string): number {
  const la = a.length, lb = b.length;
  if (la === 0 && lb === 0) return 1;
  if (la === 0 || lb === 0) return 0;
  const d: number[] = [];
  for (let j = 0; j <= lb; j++) d[j] = j;
  for (let i = 1; i <= la; i++) {
    let prev = d[0]; d[0] = i;
    for (let j = 1; j <= lb; j++) {
      const tmp = d[j];
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[j] = Math.min(d[j] + 1, d[j - 1] + 1, prev + cost);
      prev = tmp;
    }
  }
  return 1 - d[lb] / Math.max(la, lb);
}

// ----------------------- dates --------------------------------------
function parseDate(s: string): number | null {
  // returns yyyymmdd as an integer (easy to compare), or null
  s = (s || "").trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return +m[1] * 10000 + +m[2] * 100 + +m[3];
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; return y * 10000 + +m[1] * 100 + +m[2]; }
  const t = Date.parse(s);
  if (!isNaN(t)) { const d = new Date(t); return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate(); }
  return null;
}

function cutoffNum(): number {
  const d = new Date();
  return (d.getFullYear() - YEARS_VALID) * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
}

function numToIso(n: number): string {
  const y = Math.floor(n / 10000), mo = Math.floor((n % 10000) / 100), da = n % 100;
  return `${y}-${String(mo).padStart(2, "0")}-${String(da).padStart(2, "0")}`;
}

// --------------- column config: Settings sheet or defaults ----------
// The fields the script needs, each with its default letter (taken from HDR).
function settingsFields(): [string, string, string][] {
  return [
    ["pairCourse",   "Pairings - Course id (e.g. Course-Ext-ID)", HDR.pairCourse],
    ["pairTitle",    "Pairings - Course title (blank = none)",    HDR.pairTitle],
    ["pairJob",      "Pairings - Job key (COMBO)",                HDR.pairJob],
    ["pairState",    "Pairings - State (optional)",               HDR.pairState],
    ["takenEmp",     "Taken - Employee ID",                       HDR.takenEmp],
    ["takenCourse",  "Taken - Course id",                         HDR.takenCourse],
    ["takenTitle",   "Taken - Course title",                      HDR.takenTitle],
    ["takenDate",    "Taken - Date completed",                    HDR.takenDate],
    ["rosterEmp",    "Roster - Employee ID",                      HDR.rosterEmp],
    ["rosterDept",   "Roster - DEPT-CODE col (blank if COMBO is pre-glued)", HDR.rosterDept],
    ["rosterNewJob", "Roster - New job (COMBO) or JOB-CODE col", HDR.rosterNewJob],
  ];
}

// Returns the effective column spec for each field. With ASK_VIA_SETTINGS on,
// reads letters from the "Settings" sheet (auto-created with defaults if absent).
function resolveCols(workbook: ExcelScript.Workbook): { [k: string]: string } {
  const out: { [k: string]: string } = {
    pairJob: HDR.pairJob, pairCourse: HDR.pairCourse, pairTitle: HDR.pairTitle,
    pairState: HDR.pairState, takenEmp: HDR.takenEmp, takenCourse: HDR.takenCourse,
    takenTitle: HDR.takenTitle, takenDate: HDR.takenDate,
    rosterEmp: HDR.rosterEmp, rosterDept: HDR.rosterDept, rosterNewJob: HDR.rosterNewJob,
  };
  if (!USE_COLUMN_LETTERS || !ASK_VIA_SETTINGS) return out;

  const fields = settingsFields();
  let ws = workbook.getWorksheet(SHEET.settings);
  if (!ws) {
    // first run: create the Settings sheet prefilled with defaults
    ws = workbook.addWorksheet(SHEET.settings);
    const rows: string[][] = [["Setting  (put the column LETTER on the right, then re-run)", "Column"]];
    for (const [, label, def] of fields) rows.push([label, def]);
    ws.getRangeByIndexes(0, 0, rows.length, 2).setValues(rows);
    console.log("Created a 'Settings' sheet with default column letters. Edit it if needed, then re-run.");
    return out; // this run uses the defaults
  }
  // read what's there; blanks fall back to the default
  const used = ws.getUsedRange();
  const byLabel = new Map<string, string>();
  if (used) {
    for (const r of used.getTexts().slice(1)) {
      const label = String(r[0] || "").trim().toLowerCase();
      const val = String(r[1] || "").trim();
      if (label) byLabel.set(label, val);
    }
  }
  for (const [k, label] of fields) {
    const v = byLabel.get(label.toLowerCase());
    if (v !== undefined && v !== "") out[k] = v;
  }
  return out;
}

// ----------------------- sheet helpers ------------------------------
type Table = { header: string[]; rows: string[][] };

function readSheet(workbook: ExcelScript.Workbook, name: string, required: boolean): Table | null {
  const ws = workbook.getWorksheet(name);
  if (!ws) {
    if (required) throw new Error(`Missing sheet '${name}'. Check the sheet name in CONFIG.`);
    return null;
  }
  const used = ws.getUsedRange();
  if (!used) return { header: [], rows: [] };
  const texts = used.getTexts();
  const header = (texts[0] || []).map((v) => String(v));
  return { header, rows: texts.slice(1).map((r) => r.map((v) => String(v))) };
}

// for SOURCE tables (Pairings/Taken/Roster): respects USE_COLUMN_LETTERS
function col(t: Table | null, spec: string): number {
  if (spec === "") return -1;
  if (USE_COLUMN_LETTERS) return letterToIndex(spec);
  return colByHeader(t, spec);
}

// for sheets THIS script creates (Review): always matched by header name
function colByHeader(t: Table | null, headerName: string): number {
  if (!t) return -1;
  for (let i = 0; i < t.header.length; i++) {
    if (t.header[i].trim().toLowerCase() === headerName.trim().toLowerCase()) return i;
  }
  return -1;
}

function letterToIndex(s: string): number {
  s = s.trim().toUpperCase();
  let n = 0;
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1; // "A" -> 0
}

function writeSheet(workbook: ExcelScript.Workbook, name: string,
                    header: string[], rows: (string | number)[][]) {
  workbook.getWorksheet(name)?.delete();
  const ws = workbook.addWorksheet(name);
  const data: (string | number)[][] = [header, ...rows];
  ws.getRangeByIndexes(0, 0, data.length, header.length).setValues(data);
}

// helpers to coerce cell values
function str(v: unknown): string { return v === null || v === undefined ? "" : String(v); }
function key(v: unknown): string { return str(v).trim().toUpperCase(); }
