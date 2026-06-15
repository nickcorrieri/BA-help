Attribute VB_Name = "GapFinder"
' =====================================================================
'  PHASE 1 - GAP FINDER  (pure Excel / VBA, no installs needed)
' =====================================================================
'  What it does:
'    For every transferring nurse, it lists the courses her NEW job code
'    requires, then checks what she has already taken. A required course
'    is counted as DONE only if she took the EXACT same Course ID AND the
'    completion date is within the last 5 years.
'
'    Anything left over (never taken, or taken too long ago) is written to
'    the "Gaps" sheet. Those are the only rows the fuzzy step has to look
'    at later.
'
'  How to use:
'    1. Put your data on three sheets named exactly: Pairings, Taken, Roster
'       (see the headers in the CONFIG section below - edit them to match
'        your real column names if they differ).
'    2. Press Alt+F11 to open the VBA editor, then File > Import File...
'       and pick this 01_GapFinder.bas. (Or paste it into a new Module.)
'    3. Back in Excel press Alt+F8, choose BuildGaps, click Run.
'    4. Read the "Gaps" sheet.
'
'  Then run the fuzzy step (fuzzy_match.py in VS Code, OR the VBA fallback
'  in 02_FuzzyFallback.bas) to catch "same topic, different course number"
'  matches. Finally run BuildFinal (bottom of this file) for the CSV.
' =====================================================================
Option Explicit

' ============================ CONFIG =================================
' Sheet names ---------------------------------------------------------
Private Const SHEET_PAIRINGS As String = "Pairings"   ' Table 1: job code + course it needs
Private Const SHEET_TAKEN    As String = "Taken"      ' Table 2: courses each nurse has taken
Private Const SHEET_ROSTER   As String = "Roster"     ' the 12 nurses + their new job code
Private Const SHEET_CATALOG  As String = "Catalog"    ' OPTIONAL: Course ID -> Title. Set to "" to skip.
Private Const SHEET_GAPS     As String = "Gaps"       ' output (created/overwritten automatically)
Private Const SHEET_REVIEW   As String = "Review"     ' used by BuildFinal (filled by fuzzy step)
Private Const SHEET_FINAL    As String = "Final"      ' final result (created by BuildFinal)
Private Const SHEET_SETTINGS As String = "Settings"   ' optional column map (shared with the Office Script)

' Columns -------------------------------------------------------------
' If TRUE, the values below are COLUMN LETTERS (A, B, C...) and the header
' row is ignored. If FALSE, they are header NAMES matched against row 1.
' (Letters are easy: read them off the top of your sheet. Data must start in row 1.)
Private Const USE_COLUMN_LETTERS As Boolean = True

' ASK-vs-DON'T-ASK (letter mode only):
'   True  = when you run BuildGaps it asks you to CLICK each column, then saves
'           your picks to a "Settings" sheet (shared with the Office Script).
'   False = don't prompt. A "Settings" sheet is still honored if present;
'           otherwise the prefilled letters below are used.
Private Const ASK_AT_RUNTIME As Boolean = False

' Table 1 - Pairings (prefilled from the column order you listed)
Private Const H_PAIR_JOB    As String = "H"   ' Combo (Dept-code + Job-code) -- the job key
Private Const H_PAIR_COURSE As String = "A"   ' Course-ext-id
Private Const H_PAIR_TITLE  As String = "B"   ' Full-course-name (leave "" if none)
Private Const H_PAIR_STATE  As String = "L"   ' State (optional; only used if ACTIVE_STATE set)
Private Const ACTIVE_STATE  As String = ""    ' "" = use all rows; else e.g. "Active"

' Table 2 - Taken (VERIFY these letters against your sheet)
Private Const H_TAKEN_EMP    As String = "A"  ' Employee ID
Private Const H_TAKEN_COURSE As String = "B"  ' course id (matches Course-ext-id)
Private Const H_TAKEN_TITLE  As String = "C"  ' course title
Private Const H_TAKEN_DATE   As String = "D"  ' completion date

' Roster (VERIFY). H_ROST_NEWJOB must hold the SAME kind of value as H_PAIR_JOB (a Combo)
Private Const H_ROST_EMP    As String = "A"   ' Employee ID
Private Const H_ROST_NEWJOB As String = "B"   ' new job (Combo: Dept-code + Job-code)

' Catalog (only used if SHEET_CATALOG <> "" and Pairings has no title)
Private Const H_CAT_COURSE As String = "A"
Private Const H_CAT_TITLE  As String = "B"

' Rule ----------------------------------------------------------------
Private Const YEARS_VALID As Integer = 5     ' a past course only counts if taken within this many years

' Settings-sheet row labels (MUST match the Office Script exactly, so all tools
' share one Settings sheet). Don't change these unless you change them everywhere.
Private Const LBL_PAIR_COURSE As String = "Pairings - Course id (e.g. Course-Ext-ID)"
Private Const LBL_PAIR_TITLE  As String = "Pairings - Course title (blank = none)"
Private Const LBL_PAIR_JOB    As String = "Pairings - Job key (COMBO)"
Private Const LBL_PAIR_STATE  As String = "Pairings - State (optional)"
Private Const LBL_TAKEN_EMP    As String = "Taken - Employee ID"
Private Const LBL_TAKEN_COURSE As String = "Taken - Course id"
Private Const LBL_TAKEN_TITLE  As String = "Taken - Course title"
Private Const LBL_TAKEN_DATE   As String = "Taken - Date completed"
Private Const LBL_ROST_EMP As String = "Roster - Employee ID"
Private Const LBL_ROST_JOB As String = "Roster - New job (COMBO)"
' ====================================================================


Public Sub BuildGaps()
    Dim wsP As Worksheet, wsT As Worksheet, wsR As Worksheet
    On Error Resume Next
    Set wsP = ThisWorkbook.Sheets(SHEET_PAIRINGS)
    Set wsT = ThisWorkbook.Sheets(SHEET_TAKEN)
    Set wsR = ThisWorkbook.Sheets(SHEET_ROSTER)
    On Error GoTo 0
    If wsP Is Nothing Or wsT Is Nothing Or wsR Is Nothing Then
        MsgBox "Missing a sheet. I need sheets named: " & SHEET_PAIRINGS & ", " & _
               SHEET_TAKEN & ", " & SHEET_ROSTER & ".", vbExclamation
        Exit Sub
    End If

    Dim cutoff As Date
    cutoff = DateSerial(Year(Date) - YEARS_VALID, Month(Date), Day(Date))

    ' ---- if asked, let the user click each column (saves to Settings) -
    If ASK_AT_RUNTIME And USE_COLUMN_LETTERS Then AskAllColumns

    ' ---- read everything into memory (fast), anchored at A1 ---------
    Dim arrP As Variant, arrT As Variant, arrR As Variant
    arrP = ReadA1(wsP)
    arrT = ReadA1(wsT)
    arrR = ReadA1(wsR)

    ' ---- resolve each column: Settings sheet if present, else default
    Dim pJob&, pCourse&, pTitle&, pState&
    pJob = SrcCol(arrP, ColSpec(LBL_PAIR_JOB, H_PAIR_JOB))
    pCourse = SrcCol(arrP, ColSpec(LBL_PAIR_COURSE, H_PAIR_COURSE))
    Dim sTitle$: sTitle = ColSpec(LBL_PAIR_TITLE, H_PAIR_TITLE)
    Dim sState$: sState = ColSpec(LBL_PAIR_STATE, H_PAIR_STATE)
    pTitle = IIf(Len(sTitle) > 0, SrcCol(arrP, sTitle), 0)
    pState = IIf(Len(sState) > 0, SrcCol(arrP, sState), 0)

    Dim tEmp&, tCourse&, tDate&
    tEmp = SrcCol(arrT, ColSpec(LBL_TAKEN_EMP, H_TAKEN_EMP))
    tCourse = SrcCol(arrT, ColSpec(LBL_TAKEN_COURSE, H_TAKEN_COURSE))
    tDate = SrcCol(arrT, ColSpec(LBL_TAKEN_DATE, H_TAKEN_DATE))

    Dim rEmp&, rNew&
    rEmp = SrcCol(arrR, ColSpec(LBL_ROST_EMP, H_ROST_EMP))
    rNew = SrcCol(arrR, ColSpec(LBL_ROST_JOB, H_ROST_NEWJOB))

    If pJob = 0 Or pCourse = 0 Or tEmp = 0 Or tCourse = 0 Or tDate = 0 _
       Or rEmp = 0 Or rNew = 0 Then
        MsgBox "Could not find one of the expected columns." & vbCrLf & _
               "Check the column letters (or header names) in the CONFIG section.", vbExclamation
        Exit Sub
    End If

    ' ---- optional catalog title lookup ------------------------------
    Dim catTitle As Object: Set catTitle = CreateObject("Scripting.Dictionary")
    If Len(SHEET_CATALOG) > 0 And pTitle = 0 Then
        Dim wsC As Worksheet
        On Error Resume Next
        Set wsC = ThisWorkbook.Sheets(SHEET_CATALOG)
        On Error GoTo 0
        If Not wsC Is Nothing Then
            Dim arrC As Variant: arrC = ReadA1(wsC)
            Dim cC&, cTi&
            cC = SrcCol(arrC, H_CAT_COURSE): cTi = SrcCol(arrC, H_CAT_TITLE)
            If cC > 0 And cTi > 0 Then
                Dim ci&
                For ci = 2 To UBound(arrC, 1)
                    Dim ck$: ck = NKey(arrC(ci, cC))
                    If ck <> "" Then catTitle(ck) = CStr2(arrC(ci, cTi))
                Next ci
            End If
        End If
    End If

    ' ---- build "taken" lookup: emp|course -> latest date ------------
    Dim taken As Object: Set taken = CreateObject("Scripting.Dictionary")
    Dim i As Long, emp$, course$, k$
    For i = 2 To UBound(arrT, 1)
        emp = NKey(arrT(i, tEmp)): course = NKey(arrT(i, tCourse))
        If emp <> "" And course <> "" Then
            k = emp & "|" & course
            Dim dt As Date: dt = CDate(0)   ' sentinel = taken but no/unknown date
            If IsDate(arrT(i, tDate)) Then dt = CDate(arrT(i, tDate))
            If Not taken.Exists(k) Then
                taken(k) = dt
            ElseIf dt > taken(k) Then
                taken(k) = dt               ' keep the most recent completion
            End If
        End If
    Next i

    ' ---- build job -> {course -> title} -----------------------------
    Dim jobCourses As Object: Set jobCourses = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(arrP, 1)
        If Len(ACTIVE_STATE) > 0 And pState > 0 Then
            If StrComp(Trim(CStr2(arrP(i, pState))), ACTIVE_STATE, vbTextCompare) <> 0 Then GoTo NextPair
        End If
        Dim job$: job = NKey(arrP(i, pJob))
        course = NKey(arrP(i, pCourse))
        If job <> "" And course <> "" Then
            If Not jobCourses.Exists(job) Then _
                Set jobCourses(job) = CreateObject("Scripting.Dictionary")
            Dim ttl$: ttl = ""
            If pTitle > 0 Then ttl = CStr2(arrP(i, pTitle))
            If ttl = "" And catTitle.Exists(course) Then ttl = CStr2(catTitle(course))
            jobCourses(job)(course) = ttl
        End If
NextPair:
    Next i

    ' ---- walk the roster and collect gaps ---------------------------
    Dim out() As String
    Dim n As Long: n = 0
    ReDim out(1 To 5, 1 To 1000)
    Dim noReq As String: noReq = ""

    For i = 2 To UBound(arrR, 1)
        emp = NKey(arrR(i, rEmp))
        Dim newJob$: newJob = NKey(arrR(i, rNew))
        If emp <> "" Then
            If Not jobCourses.Exists(newJob) Then
                noReq = noReq & vbCrLf & "  - Emp " & emp & " / job '" & newJob & "'"
            Else
                Dim courses As Object: Set courses = jobCourses(newJob)
                Dim crsKey As Variant
                For Each crsKey In courses.Keys
                    course = CStr(crsKey)
                    k = emp & "|" & course
                    Dim status$: status = ""
                    If taken.Exists(k) Then
                        If taken(k) >= cutoff Then
                            status = ""                       ' satisfied -> not a gap
                        Else
                            status = "EXPIRED (>" & YEARS_VALID & "yr or no date)"
                        End If
                    Else
                        status = "MISSING (never taken)"
                    End If

                    If status <> "" Then
                        n = n + 1
                        If n > UBound(out, 2) Then ReDim Preserve out(1 To 5, 1 To n + 1000)
                        out(1, n) = emp
                        out(2, n) = newJob
                        out(3, n) = course
                        out(4, n) = CStr2(courses(crsKey))
                        out(5, n) = status
                    End If
                Next crsKey
            End If
        End If
    Next i

    ' ---- write the Gaps sheet ---------------------------------------
    Dim wsG As Worksheet: Set wsG = FreshSheet(SHEET_GAPS)
    wsG.Range("A1:E1").Value = Array("Employee ID", "New Job Code", _
        "Required Course ID", "Required Course Title", "Status")
    If n > 0 Then
        Dim block() As String: ReDim block(1 To n, 1 To 5)
        Dim j As Long
        For j = 1 To n
            block(j, 1) = out(1, j): block(j, 2) = out(2, j)
            block(j, 3) = out(3, j): block(j, 4) = out(4, j)
            block(j, 5) = out(5, j)
        Next j
        wsG.Range("A2").Resize(n, 5).Value = block
    End If
    wsG.Rows(1).Font.Bold = True
    wsG.Columns("A:E").AutoFit

    Dim msg$
    msg = "Done. " & n & " gap row(s) written to '" & SHEET_GAPS & "'." & vbCrLf & _
          "Courses counted as DONE = exact Course ID taken on/after " & Format(cutoff, "yyyy-mm-dd") & "."
    If Len(noReq) > 0 Then _
        msg = msg & vbCrLf & vbCrLf & "Heads up - no required courses found for:" & noReq & _
              vbCrLf & "(Check that the New Job Code matches a Job Code in Pairings.)"
    MsgBox msg, vbInformation
End Sub


' =====================================================================
'  BuildFinal - after you (or the fuzzy step) fill in the "Review" sheet,
'  this turns the confirmed gaps into the final enrollment list.
'
'  It reads the Review sheet and keeps every row whose
'  "Needs Course? (yes/no)" column says yes. Output -> "Final" sheet,
'  which you can then Save As CSV.
'
'  If you skipped the fuzzy step entirely, you can point this at the Gaps
'  sheet instead by changing SOURCE below to SHEET_GAPS - but then every
'  gap is treated as needed.
' =====================================================================
Public Sub BuildFinal()
    Const NEEDS_HEADER As String = "Needs Course? (yes/no)"
    Dim SOURCE As String: SOURCE = SHEET_REVIEW   ' change to SHEET_GAPS to skip fuzzy review

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SOURCE)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No '" & SOURCE & "' sheet found. Run the fuzzy step first " & _
               "(it creates the Review sheet), or change SOURCE in BuildFinal.", vbExclamation
        Exit Sub
    End If

    Dim arr As Variant: arr = ws.UsedRange.Value
    Dim cEmp&, cCourse&, cTitle&, cNeeds&
    cEmp = ColIdx(arr, "Employee ID")
    cCourse = ColIdx(arr, "Required Course ID")
    cTitle = ColIdx(arr, "Required Course Title")
    cNeeds = ColIdx(arr, NEEDS_HEADER)

    If cEmp = 0 Or cCourse = 0 Then
        MsgBox "Couldn't find 'Employee ID' / 'Required Course ID' columns on '" & SOURCE & "'.", vbExclamation
        Exit Sub
    End If

    Dim wsF As Worksheet: Set wsF = FreshSheet(SHEET_FINAL)
    wsF.Range("A1:C1").Value = Array("Employee ID", "Course ID", "Course Title")
    wsF.Rows(1).Font.Bold = True

    Dim i&, r&: r = 1
    For i = 2 To UBound(arr, 1)
        Dim keep As Boolean
        If cNeeds = 0 Then
            keep = True                                   ' no decision column -> keep all
        Else
            keep = (LCase(Trim(CStr2(arr(i, cNeeds)))) = "yes")
        End If
        If keep And NKey(arr(i, cEmp)) <> "" And NKey(arr(i, cCourse)) <> "" Then
            r = r + 1
            wsF.Cells(r, 1).Value = CStr2(arr(i, cEmp))
            wsF.Cells(r, 2).Value = CStr2(arr(i, cCourse))
            If cTitle > 0 Then wsF.Cells(r, 3).Value = CStr2(arr(i, cTitle))
        End If
    Next i
    wsF.Columns("A:C").AutoFit
    MsgBox (r - 1) & " enrollment row(s) written to '" & SHEET_FINAL & "'." & vbCrLf & _
           "Use File > Save As > CSV to export it.", vbInformation
End Sub


' ===================== small helpers ================================
Private Function ReadA1(ws As Worksheet) As Variant
    ' read the used data as a 2D array ANCHORED at A1, so column letters line up
    Dim ur As Range: Set ur = ws.UsedRange
    Dim lastR As Long, lastC As Long
    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1
    ReadA1 = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
End Function

Private Function SrcCol(arr As Variant, spec As String) As Long
    ' column lookup for SOURCE tables: letter (A/B/C) or header name per USE_COLUMN_LETTERS
    If Len(spec) = 0 Then SrcCol = 0: Exit Function
    If USE_COLUMN_LETTERS Then SrcCol = LetterToCol(spec) Else SrcCol = ColIdx(arr, spec)
End Function

Private Function ColSpec(label As String, def As String) As String
    ' effective spec for a field: value from the Settings sheet (by label), else default.
    ' Only consulted in letter mode.
    ColSpec = def
    If Not USE_COLUMN_LETTERS Then Exit Function
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SETTINGS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim arr As Variant: arr = ws.UsedRange.Value
    If Not IsArray(arr) Then Exit Function
    If UBound(arr, 2) < 2 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        If StrComp(Trim$(CStr2(arr(i, 1))), label, vbTextCompare) = 0 Then
            Dim v As String: v = Trim$(CStr2(arr(i, 2)))
            If v <> "" Then ColSpec = v
            Exit Function
        End If
    Next i
End Function

Public Sub AskAllColumns()
    ' prompt the user to CLICK each column; save picks to the Settings sheet
    Dim labels As Variant, defs As Variant
    labels = Array(LBL_PAIR_COURSE, LBL_PAIR_TITLE, LBL_PAIR_JOB, LBL_PAIR_STATE, _
                   LBL_TAKEN_EMP, LBL_TAKEN_COURSE, LBL_TAKEN_TITLE, LBL_TAKEN_DATE, _
                   LBL_ROST_EMP, LBL_ROST_JOB)
    defs = Array(H_PAIR_COURSE, H_PAIR_TITLE, H_PAIR_JOB, H_PAIR_STATE, _
                 H_TAKEN_EMP, H_TAKEN_COURSE, H_TAKEN_TITLE, H_TAKEN_DATE, _
                 H_ROST_EMP, H_ROST_NEWJOB)
    Dim ws As Worksheet: Set ws = EnsureSettings(labels, defs)

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        Dim rng As Range: Set rng = Nothing
        On Error Resume Next
        Set rng = Application.InputBox( _
            Prompt:="Click any cell in the column for:" & vbCrLf & vbCrLf & labels(i) & _
                    vbCrLf & vbCrLf & "(Cancel = keep '" & CStr2(ws.Cells(i + 2, 2).Value) & "')", _
            Title:="Pick column " & (i + 1) & " of " & (UBound(labels) + 1), Type:=8)
        On Error GoTo 0
        If Not rng Is Nothing Then ws.Cells(i + 2, 2).Value = ColLetter(rng.Column)
    Next i
    MsgBox "Saved your column choices to the '" & SHEET_SETTINGS & "' sheet.", vbInformation
End Sub

Private Function EnsureSettings(labels As Variant, defs As Variant) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SETTINGS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_SETTINGS
        ws.Range("A1:B1").Value = Array("Setting  (put the column LETTER on the right)", "Column")
        Dim i As Long
        For i = LBound(labels) To UBound(labels)
            ws.Cells(i + 2, 1).Value = labels(i)
            ws.Cells(i + 2, 2).Value = defs(i)
        Next i
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:B").AutoFit
    End If
    Set EnsureSettings = ws
End Function

Private Function ColLetter(n As Long) As String
    ' 1 -> "A", 27 -> "AA"
    Dim s As String, r As Long
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr$(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function LetterToCol(s As String) As Long
    ' "A" -> 1, "B" -> 2, ... "AA" -> 27
    Dim t As String: t = UCase$(Trim$(s))
    Dim i As Long, n As Long
    For i = 1 To Len(t)
        n = n * 26 + (Asc(Mid$(t, i, 1)) - 64)
    Next i
    LetterToCol = n
End Function

Private Function ColIdx(arr As Variant, headerName As String) As Long
    ' returns the 1-based column number whose row-1 header matches headerName
    Dim c As Long
    For c = LBound(arr, 2) To UBound(arr, 2)
        If StrComp(Trim(CStr2(arr(1, c))), Trim(headerName), vbTextCompare) = 0 Then
            ColIdx = c: Exit Function
        End If
    Next c
    ColIdx = 0
End Function

Private Function NKey(v As Variant) As String
    ' normalize an ID for matching: trimmed + uppercase
    If IsError(v) Then NKey = "" Else NKey = UCase$(Trim$(CStr2(v)))
End Function

Private Function CStr2(v As Variant) As String
    ' safe CStr that tolerates empty/error cells
    If IsError(v) Then CStr2 = "" ElseIf IsNull(v) Then CStr2 = "" Else CStr2 = CStr(v)
End Function

Private Function FreshSheet(nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nm
    Else
        ws.Cells.Clear
    End If
    Set FreshSheet = ws
End Function
