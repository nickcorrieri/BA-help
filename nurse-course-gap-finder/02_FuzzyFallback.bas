Attribute VB_Name = "FuzzyFallback"
' =====================================================================
'  PHASE 2 (FALLBACK) - FUZZY MATCHER IN PURE EXCEL/VBA
' =====================================================================
'  Use this ONLY if Python / the terminal is not available. It does the
'  same job as fuzzy_match.py - finds "same topic, different course
'  number" matches - but entirely inside Excel, no installs.
'
'  It reads the "Gaps" sheet (made by 01_GapFinder.bas) and the "Taken"
'  sheet, and writes a "Review" sheet with a match % and a prefilled
'  "Needs Course? (yes/no)" guess for each gap. You then eyeball it and
'  run BuildFinal (in 01_GapFinder.bas) to get the final list.
'
'  HOW TO USE
'    1. Run BuildGaps first (01_GapFinder.bas).
'    2. Import this file (Alt+F11 > File > Import File...), or paste into
'       a new Module.
'    3. Alt+F8 > BuildReviewFuzzy > Run.
'    4. Read "Review", fix the Needs column, then run BuildFinal.
'
'  NOTE: VBA has no difflib, so the scoring here is a close cousin
'  (word overlap + edit-distance ratio). It is intentionally a little
'  generous - that's fine, because a human confirms every row.
' =====================================================================
Option Explicit

' ---- must match the CONFIG in 01_GapFinder.bas ----------------------
Private Const SHEET_TAKEN    As String = "Taken"
Private Const SHEET_GAPS     As String = "Gaps"
Private Const SHEET_REVIEW   As String = "Review"
Private Const SHEET_SETTINGS As String = "Settings"   ' honored if present (BuildGaps writes it in ASK mode)

' If TRUE these are COLUMN LETTERS (header row ignored); if FALSE, header names.
' Keep this the SAME as in 01_GapFinder.bas.
Private Const USE_COLUMN_LETTERS As Boolean = True

' Table 2 - Taken (VERIFY against your sheet; same as in 01_GapFinder.bas)
Private Const H_TAKEN_EMP    As String = "A"  ' Employee ID
Private Const H_TAKEN_COURSE As String = "B"  ' course id
Private Const H_TAKEN_TITLE  As String = "C"  ' course title
Private Const H_TAKEN_DATE   As String = "D"  ' completion date

' Settings-sheet row labels (MUST match 01_GapFinder.bas and the Office Script)
Private Const LBL_TAKEN_EMP    As String = "Taken - Employee ID"
Private Const LBL_TAKEN_COURSE As String = "Taken - Course id"
Private Const LBL_TAKEN_TITLE  As String = "Taken - Course title"
Private Const LBL_TAKEN_DATE   As String = "Taken - Date completed"

Private Const YEARS_VALID As Integer = 5

' ----- OPTIONS: how strict the fuzzy "same topic" match is -----------
'   Similarity runs 0.0 - 1.0. Two knobs set the label + the prefilled guess:
'     score >= STRONG (and recent)  -> "LIKELY already covered" (prefill: no)
'     WEAK <= score < STRONG        -> "MAYBE - check it"        (prefill: yes)
'     score <  WEAK                 -> "no close match"          (prefill: yes)
'   Raise STRONG to send MORE rows to manual review; lower it to auto-clear more.
Private Const STRONG As Double = 0.75     ' 0.0 - 1.0
Private Const WEAK   As Double = 0.45     ' 0.0 - 1.0
' --------------------------------------------------------------------

' filler words removed before comparing (keep in sync with fuzzy_match.py)
Private Const STOPWORDS As String = _
    " the of a an to for in and on with your course training module intro " & _
    "introduction basic basics annual required mandatory online elearning part "


Public Sub BuildReviewFuzzy()
    Dim wsT As Worksheet, wsG As Worksheet
    On Error Resume Next
    Set wsT = ThisWorkbook.Sheets(SHEET_TAKEN)
    Set wsG = ThisWorkbook.Sheets(SHEET_GAPS)
    On Error GoTo 0
    If wsT Is Nothing Or wsG Is Nothing Then
        MsgBox "Need both '" & SHEET_GAPS & "' (run BuildGaps first) and '" & _
               SHEET_TAKEN & "'.", vbExclamation
        Exit Sub
    End If

    Dim cutoff As Date
    cutoff = DateSerial(Year(Date) - YEARS_VALID, Month(Date), Day(Date))

    Dim arrT As Variant, arrG As Variant
    arrT = ReadA1(wsT)            ' Taken is a SOURCE table -> anchored at A1 for letters
    arrG = wsG.UsedRange.Value    ' Gaps is OUR output -> always read by header name

    ' Taken (source): resolve via Settings sheet if present, else defaults.
    Dim tEmp&, tCourse&, tTitle&, tDate&
    tEmp = SrcCol(arrT, ColSpec(LBL_TAKEN_EMP, H_TAKEN_EMP))
    tCourse = SrcCol(arrT, ColSpec(LBL_TAKEN_COURSE, H_TAKEN_COURSE))
    tTitle = SrcCol(arrT, ColSpec(LBL_TAKEN_TITLE, H_TAKEN_TITLE))
    tDate = SrcCol(arrT, ColSpec(LBL_TAKEN_DATE, H_TAKEN_DATE))

    Dim gEmp&, gJob&, gCourse&, gTitle&
    gEmp = ColIdx(arrG, "Employee ID"): gJob = ColIdx(arrG, "New Job Code")
    gCourse = ColIdx(arrG, "Required Course ID"): gTitle = ColIdx(arrG, "Required Course Title")

    If tEmp = 0 Or tTitle = 0 Or gEmp = 0 Or gCourse = 0 Then
        MsgBox "Missing an expected column on Taken or Gaps. Check headers.", vbExclamation
        Exit Sub
    End If

    ' group taken titles by employee: emp -> array of "recentFlag|date|course|title"
    Dim byEmp As Object: Set byEmp = CreateObject("Scripting.Dictionary")
    Dim i&
    For i = 2 To UBound(arrT, 1)
        Dim emp$: emp = NKey(arrT(i, tEmp))
        If emp <> "" Then
            Dim dt As Date, hasDate As Boolean, recent As Boolean
            hasDate = IsDate(arrT(i, tDate))
            If hasDate Then dt = CDate(arrT(i, tDate))
            recent = (hasDate And dt >= cutoff)
            Dim rec$
            rec = IIf(recent, "1", "0") & Chr(1) & _
                  IIf(hasDate, Format(dt, "yyyy-mm-dd"), "") & Chr(1) & _
                  CStr2(arrT(i, tCourse)) & Chr(1) & CStr2(arrT(i, tTitle))
            If Not byEmp.Exists(emp) Then byEmp(emp) = ""
            byEmp(emp) = byEmp(emp) & rec & Chr(2)
        End If
    Next i

    ' build Review rows
    Dim wsR As Worksheet: Set wsR = FreshSheet(SHEET_REVIEW)
    wsR.Range("A1:J1").Value = Array("Employee ID", "New Job Code", _
        "Required Course ID", "Required Course Title", "Best Match Taken Title", _
        "Best Match Course ID", "Similarity %", "Taken Date", "Suggestion", _
        "Needs Course? (yes/no)")
    wsR.Rows(1).Font.Bold = True

    Dim outRow&: outRow = 1
    For i = 2 To UBound(arrG, 1)
        emp = NKey(arrG(i, gEmp))
        If emp <> "" Then
            Dim reqTitle$: reqTitle = CStr2(arrG(i, gTitle))
            Dim bestScore As Double, bestTitle$, bestCourse$, bestDate$, bestRecent As Boolean
            Dim foundAny As Boolean: foundAny = False
            bestScore = 0

            If byEmp.Exists(emp) Then
                Dim recs() As String, recsRecent As Boolean
                ' pass 1: only recent; pass 2: all (if none recent)
                Dim pass As Integer
                For pass = 1 To 2
                    recs = Split(byEmp(emp), Chr(2))
                    Dim r As Long
                    For r = LBound(recs) To UBound(recs)
                        If Len(recs(r)) > 0 Then
                            Dim parts() As String: parts = Split(recs(r), Chr(1))
                            Dim isRecent As Boolean: isRecent = (parts(0) = "1")
                            If (pass = 1 And isRecent) Or (pass = 2 And Not foundAny) Then
                                Dim sc As Double: sc = Similarity(reqTitle, parts(3))
                                If sc > bestScore Then
                                    bestScore = sc
                                    bestDate = parts(1): bestCourse = parts(2)
                                    bestTitle = parts(3): bestRecent = isRecent
                                End If
                                foundAny = True
                            End If
                        End If
                    Next r
                    If foundAny And pass = 1 Then Exit For   ' had recent ones, don't fall back
                Next pass
            End If

            Dim suggestion$, needs$
            If Not foundAny Then
                suggestion = "NO PRIOR COURSES - needs it": needs = "yes"
                bestTitle = "": bestCourse = "": bestDate = ""
            ElseIf bestScore >= STRONG And bestRecent Then
                suggestion = "LIKELY same topic - review": needs = "no"
            ElseIf bestScore >= WEAK Then
                suggestion = "MAYBE - check it" & IIf(bestRecent, "", "  (match is OLD/expired)")
                needs = "yes"
            Else
                suggestion = "no close match - needs it": needs = "yes"
            End If

            outRow = outRow + 1
            wsR.Cells(outRow, 1).Value = CStr2(arrG(i, gEmp))
            If gJob > 0 Then wsR.Cells(outRow, 2).Value = CStr2(arrG(i, gJob))
            wsR.Cells(outRow, 3).Value = CStr2(arrG(i, gCourse))
            wsR.Cells(outRow, 4).Value = reqTitle
            wsR.Cells(outRow, 5).Value = bestTitle
            wsR.Cells(outRow, 6).Value = bestCourse
            wsR.Cells(outRow, 7).Value = Round(bestScore * 100)
            wsR.Cells(outRow, 8).Value = bestDate
            wsR.Cells(outRow, 9).Value = suggestion
            wsR.Cells(outRow, 10).Value = needs
        End If
    Next i

    wsR.Columns("A:J").AutoFit
    MsgBox (outRow - 1) & " rows written to '" & SHEET_REVIEW & "'." & vbCrLf & _
           "Check the 'Needs Course?' column, then run BuildFinal.", vbInformation
End Sub


' ===================== fuzzy scoring ================================
Private Function Similarity(a As String, b As String) As Double
    Dim ta() As String, tb() As String
    ta = Toks(a): tb = Toks(b)
    If NumTok(ta) = 0 Or NumTok(tb) = 0 Then Similarity = 0: Exit Function

    Dim inter&, ua&, ub&
    ua = NumTok(ta): ub = NumTok(tb)
    inter = Intersect(ta, tb)

    Dim jacc As Double, overlap As Double
    jacc = inter / (ua + ub - inter)
    overlap = inter / IIf(ua < ub, ua, ub)          ' subset-friendly

    Dim lev As Double
    lev = LevRatio(JoinSorted(ta), JoinSorted(tb))  ' typo / wording tolerance

    Dim m As Double
    m = jacc
    If lev > m Then m = lev
    If 0.9 * overlap > m Then m = 0.9 * overlap
    Similarity = m
End Function

Private Function Toks(s As String) As String()
    ' lowercase, punctuation -> space, drop stopwords; return word array
    Dim t$, i&, ch$, clean$
    t = LCase$(s)
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            clean = clean & ch
        Else
            clean = clean & " "
        End If
    Next i
    Dim parts() As String: parts = Split(Application.WorksheetFunction.Trim(clean), " ")
    Dim out() As String, n&: n = 0
    ReDim out(0 To 0)
    Dim p As Variant
    For Each p In parts
        If Len(p) > 0 And InStr(STOPWORDS, " " & p & " ") = 0 Then
            ReDim Preserve out(0 To n)
            out(n) = CStr(p): n = n + 1
        End If
    Next p
    If n = 0 Then ReDim out(0 To -1)   ' empty
    Toks = out
End Function

Private Function NumTok(a() As String) As Long
    On Error Resume Next
    NumTok = UBound(a) - LBound(a) + 1
    If Err.Number <> 0 Then NumTok = 0
    On Error GoTo 0
End Function

Private Function Intersect(a() As String, b() As String) As Long
    Dim i&, j&, c&
    For i = LBound(a) To UBound(a)
        For j = LBound(b) To UBound(b)
            If a(i) = b(j) Then c = c + 1: Exit For
        Next j
    Next i
    Intersect = c
End Function

Private Function JoinSorted(a() As String) As String
    Dim arr() As String, n&: n = NumTok(a)
    If n = 0 Then JoinSorted = "": Exit Function
    ReDim arr(0 To n - 1)
    Dim i&
    For i = 0 To n - 1: arr(i) = a(LBound(a) + i): Next i
    ' simple insertion sort (titles are short)
    Dim j&, key$
    For i = 1 To n - 1
        key = arr(i): j = i - 1
        Do While j >= 0
            If arr(j) <= key Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = key
    Next i
    JoinSorted = Join(arr, " ")
End Function

Private Function LevRatio(a As String, b As String) As Double
    ' 1 - (edit distance / longer length)
    Dim la&, lb&: la = Len(a): lb = Len(b)
    If la = 0 And lb = 0 Then LevRatio = 1: Exit Function
    If la = 0 Or lb = 0 Then LevRatio = 0: Exit Function
    Dim d() As Long: ReDim d(0 To la, 0 To lb)
    Dim i&, j&
    For i = 0 To la: d(i, 0) = i: Next i
    For j = 0 To lb: d(0, j) = j: Next j
    For i = 1 To la
        For j = 1 To lb
            Dim cost&: cost = IIf(Mid$(a, i, 1) = Mid$(b, j, 1), 0, 1)
            d(i, j) = Min3(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i
    Dim maxLen&: maxLen = IIf(la > lb, la, lb)
    LevRatio = 1# - (d(la, lb) / maxLen)
End Function

Private Function Min3(a As Long, b As Long, c As Long) As Long
    Min3 = a
    If b < Min3 Then Min3 = b
    If c < Min3 Then Min3 = c
End Function


' ===================== shared helpers ==============================
Private Function ReadA1(ws As Worksheet) As Variant
    Dim ur As Range: Set ur = ws.UsedRange
    ReadA1 = ws.Range(ws.Cells(1, 1), _
        ws.Cells(ur.Row + ur.Rows.Count - 1, ur.Column + ur.Columns.Count - 1)).Value
End Function

Private Function SrcCol(arr As Variant, spec As String) As Long
    If Len(spec) = 0 Then SrcCol = 0: Exit Function
    If USE_COLUMN_LETTERS Then SrcCol = LetterToCol(spec) Else SrcCol = ColIdx(arr, spec)
End Function

Private Function ColSpec(label As String, def As String) As String
    ' effective spec: value from the Settings sheet (by label) if present, else default
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

Private Function LetterToCol(s As String) As Long
    Dim t As String: t = UCase$(Trim$(s))
    Dim i As Long, n As Long
    For i = 1 To Len(t): n = n * 26 + (Asc(Mid$(t, i, 1)) - 64): Next i
    LetterToCol = n
End Function

Private Function ColIdx(arr As Variant, headerName As String) As Long
    Dim c As Long
    For c = LBound(arr, 2) To UBound(arr, 2)
        If StrComp(Trim(CStr2(arr(1, c))), Trim(headerName), vbTextCompare) = 0 Then
            ColIdx = c: Exit Function
        End If
    Next c
    ColIdx = 0
End Function

Private Function NKey(v As Variant) As String
    If IsError(v) Then NKey = "" Else NKey = UCase$(Trim$(CStr2(v)))
End Function

Private Function CStr2(v As Variant) As String
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
