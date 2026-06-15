// Proof harness: runs the REAL Office Script (GapFinder.osts.ts) through a
// mock Excel workbook, so we can verify the matching logic actually works.
// Not for the end user -- just evidence the script is correct.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, "..");

// ---- sample data (header row first), the classic "Hazardous Materials" case
// Real 12-column layout (A..L). Letter config: course=A, title=B, Combo=H, State=L.
//   A Course-ext-id | B Full-course-name | C Course-notes | D Dept-code | E Dept-desc
//   F Job-code | G Job-title | H Combo | I Date-added | J Date-edited | K Added-by | L State
const P = (course, title, dept, job, combo, state) =>
  [course, title, "", dept, dept + " desc", job, job + " title", combo, "", "", "sys", state];
const PAIRINGS = [
  ["Course-Ext-ID", "Full-Course-Name", "Course-Notes", "DEPT-CODE", "DEPT-DESC",
   "JOB-CODE", "JOB-TITLE", "COMBO", "DATE-ADDED", "DATE-EDITED", "ADDED-BY", "STATE"],
  // COMBO is DEPT-CODE + JOB-CODE glued with no separator (4150 + 299 = 4150299)
  P("HZ-200", "Taking Care of Hazardous Materials", "4150", "299", "4150299", "Active"),
  P("FIRE-101", "Fire Safety in the Workplace", "4150", "299", "4150299", "Active"),
  P("BLS-500", "Basic Life Support for Healthcare Providers", "4150", "299", "4150299", "Active"),
  P("HIPAA-1", "Patient Privacy (HIPAA)", "4150", "299", "4150299", "Active"),
  // decoy: SAME JOB-CODE 299 but a different DEPT/COMBO -> must NOT be pulled in
  P("XRAY-9", "Radiation Safety", "9999", "299", "9999299", "Active"),
];
const TAKEN = [
  ["Employee ID", "Course ID", "Course Title", "Date Completed"],
  ["1001", "HAZ-99", "Hazardous Materials Handling", "2024-03-10"], // fuzzy -> HZ-200
  ["1001", "CPR-12", "CPR and AED Certification", "2023-11-01"],
  ["1001", "BLS-500", "Basic Life Support", "2018-01-15"],          // exact but EXPIRED
  ["1001", "HIPAA-1", "Patient Privacy (HIPAA)", "2024-05-01"],     // exact + recent -> satisfied, no gap
  ["1002", "XYZ-1", "Slip Trip and Fall Prevention", "2024-05-01"], // nothing relevant
];
const ROSTER = [
  ["Employee ID", "New Job Code"],     // holds the glued COMBO value (DEPT-CODE + JOB-CODE, e.g. 4150299)
  ["1001", "4150299"],
  ["1002", "4150299"],
];

// ---- tiny mock of the ExcelScript workbook API the script uses
function makeWorkbook(initial) {
  const sheets = new Map(Object.entries(initial).map(([n, v]) => [n, v.map((r) => r.slice())]));
  const range = (name, r0, c0) => ({
    getTexts: () => sheets.get(name).map((row) => row.map((v) => (v == null ? "" : String(v)))),
    getValues: () => sheets.get(name),
    setValues: (data) => {
      const g = sheets.get(name);
      for (let i = 0; i < data.length; i++) {
        const rr = r0 + i;
        if (!g[rr]) g[rr] = [];
        for (let j = 0; j < data[i].length; j++) g[rr][c0 + j] = data[i][j];
      }
    },
  });
  const sheet = (name) => ({
    getName: () => name,
    getUsedRange: () => { const g = sheets.get(name); return g && g.length ? range(name, 0, 0) : undefined; },
    getRangeByIndexes: (r, c) => range(name, r, c),
    delete: () => sheets.delete(name),
  });
  return {
    getWorksheet: (n) => (sheets.has(n) ? sheet(n) : undefined),
    addWorksheet: (n) => { sheets.set(n, []); return sheet(n); },
    _get: (n) => sheets.get(n),
  };
}

async function loadScript(mode) {
  // take the real file verbatim, force MODE, expose main via an export shim
  let src = fs.readFileSync(path.join(root, "GapFinder.osts.ts"), "utf8");
  src = src.replace(/const MODE:[^=]*=\s*"review";/, `const MODE = "${mode}";`);
  src += "\nexport { main };\n";
  const tmp = path.join(here, `_tmp_${mode}.mts`);
  fs.writeFileSync(tmp, src);
  const mod = await import("file://" + tmp);
  fs.unlinkSync(tmp);
  return mod.main;
}

function asObjects(grid) {
  const [head, ...rows] = grid;
  return rows.map((r) => Object.fromEntries(head.map((h, i) => [h, r[i]])));
}

let failures = 0;
function check(label, cond) {
  console.log(`  ${cond ? "PASS" : "FAIL"}  ${label}`);
  if (!cond) failures++;
}

console.log("Running REAL Office Script logic against sample data...\n");

const main = await loadScript("review");
const wb = makeWorkbook({ Pairings: PAIRINGS, Taken: TAKEN, Roster: ROSTER });
main(wb);

const gaps = asObjects(wb._get("Gaps"));
const review = asObjects(wb._get("Review"));

console.log("GAPS sheet:");
for (const g of gaps) console.log("   ", g["Employee ID"], g["Required Course ID"], "-", g["Status"]);
console.log("\nREVIEW sheet:");
for (const r of review) {
  console.log("   ", r["Employee ID"], r["Required Course ID"],
    "| best:", r["Best Match Taken Title"] || "(none)",
    "| " + r["Similarity %"] + "%",
    "| needs=" + r["Needs Course? (yes/no)"],
    "| " + r["Suggestion"]);
}

console.log("\nChecks:");
const g = (emp, c) => gaps.find((x) => x["Employee ID"] === emp && x["Required Course ID"] === c);
const r = (emp, c) => review.find((x) => x["Employee ID"] === emp && x["Required Course ID"] === c);

// 1001 took HIPAA-1 recently and exactly -> must NOT appear as a gap
check("1001 HIPAA-1 satisfied (no gap)", !g("1001", "HIPAA-1"));
// 1001 BLS-500 taken but in 2018 -> EXPIRED gap
check("1001 BLS-500 flagged EXPIRED", !!g("1001", "BLS-500") && g("1001", "BLS-500")["Status"].startsWith("EXPIRED"));
// 1001 HZ-200 never taken -> MISSING gap, but fuzzy finds Hazardous Materials Handling
check("1001 HZ-200 is a gap", !!g("1001", "HZ-200"));
check("1001 HZ-200 fuzzy-matched HAZ-99", r("1001", "HZ-200")["Best Match Course ID"] === "HAZ-99");
check("1001 HZ-200 similarity >= 60%", Number(r("1001", "HZ-200")["Similarity %"]) >= 60);
// 1001 FIRE-101 -> nothing related -> real gap, needs yes
check("1001 FIRE-101 needs=yes", r("1001", "FIRE-101")["Needs Course? (yes/no)"] === "yes");
// 1002 needs everything (took nothing relevant)
check("1002 HZ-200 needs=yes", r("1002", "HZ-200")["Needs Course? (yes/no)"] === "yes");
check("1002 has 4 gaps", gaps.filter((x) => x["Employee ID"] === "1002").length === 4);
// Combo scoping: XRAY-9 is RN in a DIFFERENT department -> must NOT be a gap for D5-RN nurses
check("Combo scoping excludes other-dept RN course (XRAY-9)", !g("1001", "XRAY-9") && !g("1002", "XRAY-9"));

// ---- finalize mode: feed an edited Review back in
console.log("\nFinalize mode (simulating reviewer keeping the real gaps):");
const edited = [wb._get("Review")[0], ...wb._get("Review").slice(1).map((row) => row.slice())];
// header indices
const head = edited[0];
const iNeeds = head.indexOf("Needs Course? (yes/no)");
const iEmp = head.indexOf("Employee ID");
const iCourse = head.indexOf("Required Course ID");
// reviewer decides 1001 HZ-200 IS covered by the fuzzy match -> set to no
for (let i = 1; i < edited.length; i++) {
  if (edited[i][iEmp] === "1001" && edited[i][iCourse] === "HZ-200") edited[i][iNeeds] = "no";
}
const wb2 = makeWorkbook({ Review: edited });
const mainF = await loadScript("finalize");
mainF(wb2);
const final = asObjects(wb2._get("Final"));
console.log("FINAL sheet:");
for (const f of final) console.log("   ", f["Employee ID"], f["Course ID"], "-", f["Course Title"]);
check("Final excludes the HZ-200 row reviewer marked 'no'",
  !final.find((x) => x["Employee ID"] === "1001" && x["Course ID"] === "HZ-200"));
check("Final keeps 1001 FIRE-101", !!final.find((x) => x["Employee ID"] === "1001" && x["Course ID"] === "FIRE-101"));

console.log("\n" + (failures === 0 ? "ALL CHECKS PASSED" : `${failures} CHECK(S) FAILED`));
process.exit(failures === 0 ? 0 : 1);
