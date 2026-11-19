Attribute VB_Name = "PerfMacros"
'===============================================================================
' PerfMacros - the macros behind the buttons. Thin wrappers around PerfService
' and PerfOptimiser.
'===============================================================================
Option Explicit

Private mCalc As Long

Private Sub BeginWork()
    mCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = -4135          ' xlCalculationManual
    Application.StatusBar = APP_NAME & ": working..."
End Sub

Private Sub EndWork()
    If mCalc = 0 Then mCalc = -4105          ' xlCalculationAutomatic
    Application.Calculation = mCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Private Sub Fail(ByVal what As String)
    Dim d As String: d = Err.Description
    EndWork
    MsgBox what & " failed:" & vbCrLf & vbCrLf & d, vbExclamation, APP_NAME
End Sub

Private Sub Report(ByVal m As String)
    If InStr(1, m, "nothing was saved", vbTextCompare) > 0 Then
        MsgBox m, vbExclamation, APP_NAME
    Else
        MsgBox m, vbInformation, APP_NAME
    End If
End Sub

Public Sub Setup_Database()
    Dim p As String, msg As String
    On Error GoTo EH
    p = DbPath()
    If Len(Dir$(p)) = 0 Then
        If MsgBox("Create a new Access database at:" & vbCrLf & p & vbCrLf & vbCrLf & _
                  "The reference lists, cedants, treaties and stats in this workbook will be loaded into it.", _
                  vbOKCancel + vbQuestion, APP_NAME) <> vbOK Then Exit Sub
    End If
    BeginWork
    msg = PerfService.SetupDatabase()
    EndWork
    Report msg
    Exit Sub
EH:
    Fail "Setup"
End Sub

Public Sub Pull_All()
    Dim m As String
    On Error GoTo EH
    If PendingEdits(SH_TREATIES) Or PendingEdits(SH_STATS) Then
        If MsgBox("Treaties or Stats have rows that are not saved to Access (generated rows waiting for " & _
                  "figures count too)." & vbCrLf & "Pulling will discard them. Continue?", _
                  vbOKCancel + vbExclamation, APP_NAME) <> vbOK Then Exit Sub
    End If
    BeginWork
    m = PerfService.PullAll()
    EndWork
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Pull"
End Sub

Public Sub Push_Masters()
    Dim m As String
    On Error GoTo EH
    BeginWork
    m = PerfService.PushMasters()
    EndWork
    Report m
    Exit Sub
EH:
    Fail "Push lists + cedants"
End Sub

Public Sub Pull_Masters()
    Dim m As String
    On Error GoTo EH
    BeginWork
    m = PerfService.PullMasters()
    EndWork
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Pull lists + cedants"
End Sub

Public Sub Push_Treaties()
    Dim m As String
    On Error GoTo EH
    BeginWork
    m = PerfService.PushTreaties()
    EndWork
    Report m
    Exit Sub
EH:
    Fail "Push treaties"
End Sub

Public Sub Pull_Treaties()
    Dim m As String
    On Error GoTo EH
    If PendingEdits(SH_TREATIES) Then
        If MsgBox("Treaties has rows that are not saved to Access. Pulling will discard them. Continue?", _
                  vbOKCancel + vbExclamation, APP_NAME) <> vbOK Then Exit Sub
    End If
    BeginWork
    m = PerfService.PullTreaties()
    EndWork
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Pull treaties"
End Sub

Public Sub Generate_Stat_Rows()
    Dim m As String
    On Error GoTo EH
    BeginWork
    m = PerfService.GenerateStatRows()
    EndWork
    MsgBox m, vbInformation, APP_NAME
    ThisWorkbook.Worksheets(SH_STATS).Activate
    Exit Sub
EH:
    Fail "Generate stat rows"
End Sub

Public Sub Push_Stats()
    Dim m As String
    On Error GoTo EH
    BeginWork
    m = PerfService.PushStats()
    EndWork
    Report m
    Exit Sub
EH:
    Fail "Push stats"
End Sub

Public Sub Pull_Stats()
    Dim m As String
    On Error GoTo EH
    If PendingEdits(SH_STATS) Then
        If MsgBox("Stats has rows that are not saved to Access (generated rows waiting for figures count too)." & _
                  vbCrLf & "Pulling will discard them. Continue?", vbOKCancel + vbExclamation, APP_NAME) <> vbOK Then Exit Sub
    End If
    BeginWork
    m = PerfService.PullStats()
    EndWork
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Pull stats"
End Sub

Public Sub Run_Optimiser()
    Dim m As String
    On Error GoTo EH
    Application.Calculate                      ' make sure Stats / Shares formulas are current
    Application.ScreenUpdating = False
    m = PerfOptimiser.RunOptimiser()
    Application.ScreenUpdating = True
    Application.Calculate
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Optimiser"
End Sub

Public Sub Use_Optimised_As_Scenario()
    Dim m As String
    On Error GoTo EH
    m = PerfOptimiser.UseOptimisedAsScenario()
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Copy optimised shares"
End Sub

Public Sub Clear_Scenario()
    If MsgBox("Clear the Scenario share column? Blank scenario = current share.", _
              vbOKCancel + vbQuestion, APP_NAME) <> vbOK Then Exit Sub
    PerfOptimiser.ClearScenario
End Sub

Public Sub Add_Cedant()
    Dim nm As Variant, ctry As Variant, m As String
    On Error GoTo EH
    nm = Application.InputBox("Cedant name:", APP_NAME & " - add cedant", Type:=2)
    If VarType(nm) = vbBoolean Then Exit Sub
    If Len(Trim$(CStr(nm))) = 0 Then Exit Sub
    ctry = Application.InputBox("Home country of '" & Trim$(CStr(nm)) & "' (required)" & vbCrLf & _
        "Country name or ISO2 code, e.g. Zambia or ZM.", APP_NAME & " - add cedant", Type:=2)
    If VarType(ctry) = vbBoolean Then Exit Sub
    BeginWork
    m = PerfService.AddCedant(CStr(nm), CStr(ctry))
    EndWork
    MsgBox m, vbInformation, APP_NAME
    Exit Sub
EH:
    Fail "Add cedant"
End Sub

Public Sub Open_Database()
    On Error GoTo EH
    Dim p As String: p = DbPath()
    If Len(Dir$(p)) = 0 Then
        MsgBox "No database yet at" & vbCrLf & p & vbCrLf & "Run Setup first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    ThisWorkbook.FollowHyperlink p
    Exit Sub
EH:
    Fail "Open database"
End Sub

' Unsaved rows on Treaties or Stats: new, marked for delete, or edited since the pull.
Private Function PendingEdits(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet, vals As Variant, i As Long, c As Long, lastRow As Long, hasData As Boolean
    Dim c1 As Long, c2 As Long, idCol As Long, delCol As Long, fpCol As Long, isTrt As Boolean
    isTrt = (sheetName = SH_TREATIES)
    If isTrt Then
        lastRow = LAST_ROW_TRT: c1 = T_CEDANT: c2 = T_DELETE: idCol = T_ID: delCol = T_DELETE: fpCol = T_FP
    Else
        lastRow = LAST_ROW_STAT: c1 = S_CEDANT: c2 = S_DELETE: idCol = S_ID: delCol = S_DELETE: fpCol = S_FP
    End If
    Set ws = ThisWorkbook.Worksheets(sheetName)
    vals = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastRow, fpCol)).Value
    For i = 1 To UBound(vals, 1)
        hasData = False
        For c = c1 To c2
            If IsError(vals(i, c)) Then hasData = True: Exit For
            If Len(PerfService.Txt(vals(i, c))) > 0 Then hasData = True: Exit For
        Next c
        If hasData Then
            If Len(PerfService.Txt(vals(i, idCol))) = 0 Then PendingEdits = True: Exit Function
            If UCase$(PerfService.Txt(vals(i, delCol))) = "X" Then PendingEdits = True: Exit Function
            If isTrt Then
                If PerfService.TreatyFingerprint(vals, i) <> PerfService.Txt(vals(i, fpCol)) Then PendingEdits = True: Exit Function
            Else
                If PerfService.StatFingerprint(vals, i) <> PerfService.Txt(vals(i, fpCol)) Then PendingEdits = True: Exit Function
            End If
        End If
    Next i
End Function

' ------------------------------------------------------------------------------
' Adds the buttons. Run once after importing the modules. Safe to run again.
' ------------------------------------------------------------------------------
Public Sub InstallButtons()
    Dim ws As Worksheet, nm As Variant
    Set ws = ThisWorkbook.Worksheets(SH_START)
    ClearButtons ws
    AddButton ws, "C13", "1. Setup / create database", "Setup_Database"
    AddButton ws, "C15", "Pull everything from Access", "Pull_All"
    AddButton ws, "C17", "Push reference lists + cedants", "Push_Masters"
    AddButton ws, "C19", "Push treaties", "Push_Treaties"
    AddButton ws, "C21", "Generate stat rows", "Generate_Stat_Rows"
    AddButton ws, "C23", "Push stats", "Push_Stats"
    AddButton ws, "C25", "Run share optimiser", "Run_Optimiser"
    AddButton ws, "C27", "Add a cedant", "Add_Cedant"
    AddButton ws, "C29", "Open database in Access", "Open_Database"

    Set ws = ThisWorkbook.Worksheets(SH_TREATIES)
    ClearButtons ws
    AddButton ws, "H1", "Push treaties", "Push_Treaties", 1, 120
    AddButton ws, "J1", "Pull treaties", "Pull_Treaties", 1, 110
    AddButton ws, "L1", "Generate stat rows", "Generate_Stat_Rows", 1, 140
    AddButton ws, "O1", "Add cedant", "Add_Cedant", 1, 100

    Set ws = ThisWorkbook.Worksheets(SH_STATS)
    ClearButtons ws
    AddButton ws, "I1", "Push stats", "Push_Stats", 1, 110
    AddButton ws, "K1", "Pull stats", "Pull_Stats", 1, 110
    AddButton ws, "M1", "Generate stat rows", "Generate_Stat_Rows", 1, 140

    Set ws = ThisWorkbook.Worksheets(SH_SHARES)
    ClearButtons ws
    AddButton ws, "N1", "Run optimiser", "Run_Optimiser", 1, 120
    AddButton ws, "P1", "Use optimised as scenario", "Use_Optimised_As_Scenario", 1, 170
    AddButton ws, "S1", "Clear scenario", "Clear_Scenario", 1, 110

    For Each nm In Array(SH_CEDANTS, SH_COUNTRIES, SH_RATES, SH_LISTS)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        ClearButtons ws
        AddButton ws, "H1", "Push lists + cedants", "Push_Masters", 1, 150
        AddButton ws, "K1", "Pull from Access", "Pull_Masters", 1, 140
    Next nm
    AddButton ThisWorkbook.Worksheets(SH_CEDANTS), "N1", "Add cedant", "Add_Cedant", 1, 110
    ThisWorkbook.Worksheets(SH_START).Activate
End Sub

Private Sub ClearButtons(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal anchor As String, ByVal caption As String, _
                      ByVal macroName As String, Optional ByVal rowsHigh As Long = 2, _
                      Optional ByVal widthPt As Double = 210)
    Dim r As Range, b As Object
    Set r = ws.Range(anchor)
    Set b = ws.Buttons.Add(r.Left + 2, r.Top + 2, widthPt, r.Resize(rowsHigh, 1).Height * 0.9 + 12)
    b.Name = "btn_" & macroName & "_" & Replace(anchor, "$", "")
    b.Caption = caption
    b.OnAction = macroName
    b.Font.Bold = True
End Sub
