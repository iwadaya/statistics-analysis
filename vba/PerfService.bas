Attribute VB_Name = "PerfService"
'===============================================================================
' PerfService - validation and orchestration between the sheets and Access
' (Universe: server/src/services + validation/*). Same rules as the Aggregate
' Tracker:
'   * Access is the system of record; the sheets are a working copy.
'   * A push is all-or-nothing: one transaction, rolled back on the first failure.
'   * Updates and deletes are optimistic-locked on row_version.
'   * Every write is recorded in audit_log.
' Statistics are captured at 100% of the cedant's figures, one row per UW year.
'===============================================================================
Option Explicit

Private Const A_NONE As Long = 0
Private Const A_INS As Long = 1
Private Const A_UPD As Long = 2
Private Const A_DEL As Long = 3
Private Const A_SAME As Long = 4
Private Const A_EMPTY As Long = 5

Public Const EMPTY_STATUS As String = "EMPTY - enter figures"

' ============================== small helpers =================================
Private Function Sh(ByVal nm As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(nm)
End Function

Public Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

Private Function IsBlankV(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    IsBlankV = (Len(Txt(v)) = 0)
End Function

Private Function NullIfBlank(ByVal v As Variant) As Variant
    If Len(Txt(v)) = 0 Then NullIfBlank = Null Else NullIfBlank = Txt(v)
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If IsBlankV(v) Then Exit Function
    IsNum = IsNumeric(v)
End Function

' Normalised text for fingerprints: numbers via Str$, dates as ISO, blank stays blank.
Private Function Norm(ByVal v As Variant, Optional ByVal asDate As Boolean = False, _
                      Optional ByVal upper As Boolean = False) As String
    If IsError(v) Then Norm = "#ERR": Exit Function
    If IsBlankV(v) Then Exit Function
    If asDate And IsDate(v) Then
        Norm = Format$(CDate(v), "yyyy-mm-dd")
    ElseIf IsNum(v) And Not asDate Then
        Norm = Trim$(Str$(CDbl(v)))
    ElseIf upper Then
        Norm = UCase$(Txt(v))
    Else
        Norm = Txt(v)
    End If
End Function

Private Function RefText(ByVal v As Variant) As String
    RefText = UCase$(Txt(v))
    If Len(RefText) = 0 Then RefText = DEFAULT_REF
End Function

Private Function GridValues(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal nCols As Long) As Variant
    GridValues = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastRow, nCols)).Value
End Function

Private Function SheetRow(ByVal i As Long) As Long
    SheetRow = FIRST_ROW + i - 1
End Function

Private Function AnyContent(ByVal vals As Variant, ByVal i As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim c As Long
    For c = c1 To c2
        If IsError(vals(i, c)) Then AnyContent = True: Exit Function
        If Not IsBlankV(vals(i, c)) Then AnyContent = True: Exit Function
    Next c
End Function

Private Function ShareValue(ByVal v As Variant, ByVal label As String, ByRef out As Variant) As String
    If IsBlankV(v) Then
        out = Null
    ElseIf IsNum(v) Then
        If CDbl(v) < 0 Or CDbl(v) > 1 Then ShareValue = label & " must be between 0% and 100%": Exit Function
        out = CDbl(v)
    Else
        ShareValue = label & " must be a percentage, e.g. 7.5%"
    End If
End Function

Private Function AmountValue(ByVal v As Variant, ByVal label As String, ByRef out As Variant, _
                             ByVal blankAsZero As Boolean) As String
    If IsBlankV(v) Then
        If blankAsZero Then out = 0# Else out = Null
    ElseIf IsNum(v) Then
        out = CDbl(v)
    Else
        AmountValue = label & " must be a number"
    End If
End Function

Private Sub WriteStatusCol(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal stat As Variant)
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastRow, 1)).Value = stat
End Sub


' ================================ master data ==================================
' Push order: countries -> currencies/rates -> treaty types/classes -> cedants.
' Validation runs over every sheet first; if anything fails nothing is written.
Public Function PushMasters() As String
    Dim cn As Object, inTx As Boolean, errs As String
    Dim vC As Variant, vR As Variant, vL As Variant, vK As Variant
    Dim sC() As Variant, sR() As Variant, sK() As Variant
    Dim nCtry As Long, nCcy As Long, nRate As Long, nList As Long, nCed As Long

    vC = GridValues(Sh(SH_COUNTRIES), LAST_ROW_CTRY, CTRY_COLS)
    vR = GridValues(Sh(SH_RATES), LAST_ROW_RATE, RATE_COLS)
    vL = GridValues(Sh(SH_LISTS), LAST_ROW_LIST, 5)
    vK = GridValues(Sh(SH_CEDANTS), LAST_ROW_CED, CED_COLS)
    ReDim sC(1 To UBound(vC, 1), 1 To 1)
    ReDim sR(1 To UBound(vR, 1), 1 To 1)
    ReDim sK(1 To UBound(vK, 1), 1 To 1)

    errs = ValidateCountries(vC, sC) & ValidateRates(vR, sR) & ValidateLists(vL) & ValidateCedants(vK, vC, sK)
    If Len(errs) > 0 Then
        WriteCol Sh(SH_COUNTRIES), 6, sC
        WriteCol Sh(SH_RATES), 5, sR
        WriteCol Sh(SH_CEDANTS), 5, sK
        PushMasters = "Nothing was saved - fix these first:" & vbCrLf & Left$(errs, 1500)
        Exit Function
    End If

    On Error GoTo EH
    Set cn = PerfDb.Connect()
    cn.BeginTrans
    inTx = True
    nCtry = WriteCountries(cn, vC)
    WriteRates cn, vR, nCcy, nRate
    nList = WriteLists(cn, vL)
    nCed = WriteCedants(cn, vK)
    cn.CommitTrans
    inTx = False
    PerfDb.CloseQuietly cn
    PullMasters
    PushMasters = "Saved to Access: " & nCtry & " countries, " & nCcy & " currencies, " & nRate & _
        " FX rate change(s), " & nList & " treaty types/classes, " & nCed & " cedants."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    If inTx Then PerfDb.RollbackQuietly cn
    PerfDb.CloseQuietly cn
    Err.Raise vbObjectError + 1400, APP_NAME, "Nothing was saved: " & d
End Function

Private Sub WriteCol(ByVal ws As Worksheet, ByVal col As Long, ByVal arr As Variant)
    ws.Range(ws.Cells(FIRST_ROW, col), ws.Cells(FIRST_ROW + UBound(arr, 1) - 1, col)).Value = arr
End Sub

Private Function RowBlank(ByVal v As Variant, ByVal i As Long, ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim c As Long
    For c = c1 To c2
        If IsError(v(i, c)) Then Exit Function
        If Not IsBlankV(v(i, c)) Then Exit Function
    Next c
    RowBlank = True
End Function

Private Function ErrLine(ByVal sheetName As String, ByVal i As Long, ByVal m As String) As String
    ErrLine = "  " & sheetName & " row " & SheetRow(i) & ": " & m & vbCrLf
End Function

' Countries: A name | B ISO2 | C ISO3 | D region | E id | F status
Private Function ValidateCountries(ByVal v As Variant, ByRef st() As Variant) As String
    Dim i As Long, e As String, m As String, seenC As Object, seenN As Object, cd As String, nm As String
    Set seenC = CreateObject("Scripting.Dictionary"): seenC.CompareMode = vbTextCompare
    Set seenN = CreateObject("Scripting.Dictionary"): seenN.CompareMode = vbTextCompare
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            nm = Txt(v(i, 1)): cd = UCase$(Txt(v(i, 2))): m = ""
            If Len(nm) = 0 Then
                m = "country name is required"
            ElseIf Len(cd) <> 2 Then
                m = "ISO2 code must be 2 letters"
            ElseIf Len(Txt(v(i, 3))) > 0 And Len(Txt(v(i, 3))) <> 3 Then
                m = "ISO3 code must be 3 letters"
            ElseIf seenC.Exists(cd) Then
                m = "ISO2 " & cd & " repeats row " & seenC(cd)
            ElseIf seenN.Exists(nm) Then
                m = "name repeats row " & seenN(nm)
            End If
            If Len(m) > 0 Then
                st(i, 1) = "ERROR: " & m: e = e & ErrLine(SH_COUNTRIES, i, m)
            Else
                seenC(cd) = SheetRow(i): seenN(nm) = SheetRow(i): st(i, 1) = v(i, 6)
            End If
        End If
    Next i
    ValidateCountries = e
End Function

Private Function WriteCountries(ByVal cn As Object, ByVal v As Variant) As Long
    Dim i As Long, byCode As Object, byName As Object, cd As String, nm As String, id As Variant
    Set byCode = PerfRepo.CountryCodeMap(cn)
    Set byName = PerfRepo.CountryMap(cn)
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            nm = Txt(v(i, 1)): cd = UCase$(Txt(v(i, 2))): id = Empty
            If byCode.Exists(cd) Then
                id = byCode(cd)
            ElseIf byName.Exists(nm) Then
                id = byName(nm)
            End If
            PerfRepo.UpsertCountry cn, id, nm, cd, NullIfBlank(UCase$(Txt(v(i, 3)))), NullIfBlank(v(i, 4))
            WriteCountries = WriteCountries + 1
        End If
    Next i
End Function

' Rates: A code | B name | C USD per 1 unit | D effective date | E status
Private Function ValidateRates(ByVal v As Variant, ByRef st() As Variant) As String
    Dim i As Long, e As String, m As String, seen As Object, cd As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            cd = UCase$(Txt(v(i, 1))): m = ""
            If IsError(v(i, 3)) Or IsError(v(i, 4)) Then
                m = "rate or date cell contains an error value"
            ElseIf Len(cd) <> 3 Then
                m = "currency code must be 3 letters"
            ElseIf seen.Exists(cd) Then
                m = cd & " repeats row " & seen(cd)
            ElseIf Not IsBlankV(v(i, 3)) Then
                If Not IsNumeric(v(i, 3)) Then
                    m = "rate must be a number (USD per 1 " & cd & ")"
                ElseIf CDbl(v(i, 3)) <= 0 Then
                    m = "rate must be greater than zero"
                ElseIf IsBlankV(v(i, 4)) Or Not IsDate(v(i, 4)) Then
                    m = "effective date is required when a rate is entered"
                End If
            End If
            If Len(m) > 0 Then
                st(i, 1) = "ERROR: " & m: e = e & ErrLine(SH_RATES, i, m)
            Else
                seen(cd) = SheetRow(i): st(i, 1) = v(i, 5)
            End If
        End If
    Next i
    ValidateRates = e
End Function

Private Sub WriteRates(ByVal cn As Object, ByVal v As Variant, ByRef nCcy As Long, ByRef nRate As Long)
    Dim i As Long, have As Object, cd As String, nm As String
    Set have = PerfRepo.CurrencyMap(cn)
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            cd = UCase$(Txt(v(i, 1))): nm = Txt(v(i, 2))
            If Len(nm) = 0 Then nm = cd
            PerfRepo.UpsertCurrency cn, have.Exists(cd), cd, Left$(nm, 60)
            nCcy = nCcy + 1
            If Not IsBlankV(v(i, 3)) Then
                If PerfRepo.UpsertRate(cn, cd, CDbl(v(i, 3)), DateValue(CDate(v(i, 4)))) Then nRate = nRate + 1
            End If
        End If
    Next i
End Sub

' Lists: A treaty type | B category | D class of business | E code
Private Function ValidateLists(ByVal v As Variant) As String
    Dim i As Long, e As String, seenT As Object, seenB As Object, nm As String
    Set seenT = CreateObject("Scripting.Dictionary"): seenT.CompareMode = vbTextCompare
    Set seenB = CreateObject("Scripting.Dictionary"): seenB.CompareMode = vbTextCompare
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 2) Then
            nm = Txt(v(i, 1))
            If Len(nm) = 0 Then
                e = e & ErrLine(SH_LISTS, i, "treaty type name is required")
            ElseIf seenT.Exists(nm) Then
                e = e & ErrLine(SH_LISTS, i, "treaty type '" & nm & "' repeats")
            Else
                seenT(nm) = True
            End If
        End If
        If Not RowBlank(v, i, 4, 5) Then
            nm = Txt(v(i, 4))
            If Len(nm) = 0 Then
                e = e & ErrLine(SH_LISTS, i, "class of business name is required")
            ElseIf seenB.Exists(nm) Then
                e = e & ErrLine(SH_LISTS, i, "class of business '" & nm & "' repeats")
            Else
                seenB(nm) = True
            End If
        End If
    Next i
    ValidateLists = e
End Function

Private Function WriteLists(ByVal cn As Object, ByVal v As Variant) As Long
    Dim i As Long, tt As Object, cob As Object, nm As String
    Set tt = PerfRepo.TreatyTypeMap(cn)
    Set cob = PerfRepo.CobMap(cn)
    For i = 1 To UBound(v, 1)
        nm = Txt(v(i, 1))
        If Len(nm) > 0 Then
            PerfRepo.UpsertTreatyType cn, tt.Exists(nm), Left$(nm, 80), NullIfBlank(v(i, 2))
            WriteLists = WriteLists + 1
        End If
        nm = Txt(v(i, 4))
        If Len(nm) > 0 Then
            PerfRepo.UpsertCob cn, cob.Exists(nm), Left$(nm, 80), NullIfBlank(v(i, 5))
            WriteLists = WriteLists + 1
        End If
    Next i
End Function

' Cedants: A name | B home country | C active Y/N | D id | E status
Private Function ValidateCedants(ByVal v As Variant, ByVal vCountries As Variant, ByRef st() As Variant) As String
    Dim i As Long, e As String, m As String, seen As Object, ctry As Object, nm As String, hc As String, act As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    Set ctry = CreateObject("Scripting.Dictionary"): ctry.CompareMode = vbTextCompare
    For i = 1 To UBound(vCountries, 1)
        If Len(Txt(vCountries(i, 1))) > 0 Then ctry(Txt(vCountries(i, 1))) = True
    Next i
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            nm = Txt(v(i, 1)): hc = Txt(v(i, 2)): act = UCase$(Txt(v(i, 3))): m = ""
            If Len(nm) = 0 Then
                m = "cedant name is required"
            ElseIf Len(nm) > 150 Then
                m = "cedant name is longer than 150 characters"
            ElseIf seen.Exists(nm) Then
                m = "name repeats row " & seen(nm)
            ElseIf Len(hc) = 0 Then
                m = "home country is required"
            ElseIf Not ctry.Exists(hc) Then
                m = "home country '" & hc & "' is not on the Countries sheet"
            ElseIf Len(act) > 0 And act <> "Y" And act <> "N" Then
                m = "Active must be Y or N"
            ElseIf Not IsBlankV(v(i, 4)) And Not IsNumeric(v(i, 4)) Then
                m = "ID column was edited - Pull masters again"
            End If
            If Len(m) > 0 Then
                st(i, 1) = "ERROR: " & m: e = e & ErrLine(SH_CEDANTS, i, m)
            Else
                seen(nm) = SheetRow(i): st(i, 1) = v(i, 5)
            End If
        End If
    Next i
    ValidateCedants = e
End Function

Private Function WriteCedants(ByVal cn As Object, ByVal v As Variant) As Long
    Dim i As Long, byName As Object, ctry As Object, nm As String, hc As Variant, active As Boolean
    Dim newId As Long, cid As Long
    Set byName = PerfRepo.CedantMap(cn)
    Set ctry = PerfRepo.CountryMap(cn)          ' includes countries written earlier in this transaction
    For i = 1 To UBound(v, 1)
        If Not RowBlank(v, i, 1, 4) Then
            nm = Txt(v(i, 1))
            hc = ctry(Txt(v(i, 2)))
            active = (UCase$(Txt(v(i, 3))) <> "N")
            If Not IsBlankV(v(i, 4)) Then
                cid = CLng(v(i, 4))
                If PerfRepo.UpdateCedant(cn, cid, nm, hc, active) = 0 Then
                    Err.Raise vbObjectError + 1410, APP_NAME, "Cedants row " & SheetRow(i) & ": ID " & cid & _
                        " is not in the database - Pull masters again"
                End If
                PerfRepo.WriteAudit cn, "cedant", cid, "CEDANT_SAVED", nm
            ElseIf byName.Exists(nm) Then
                cid = CLng(byName(nm))
                PerfRepo.UpdateCedant cn, cid, nm, hc, active
                PerfRepo.WriteAudit cn, "cedant", cid, "CEDANT_SAVED", nm
            Else
                newId = PerfRepo.InsertCedant(cn, nm, hc, active)
                PerfRepo.WriteAudit cn, "cedant", newId, "CEDANT_CREATED", nm
            End If
            WriteCedants = WriteCedants + 1
        End If
    Next i
End Function

Public Function PullMasters() As String
    Dim cn As Object, rows As Variant, out() As Variant, r As Long, nK As Long
    On Error GoTo EH
    Set cn = PerfDb.Connect()

    ' Countries
    rows = PerfRepo.SelectCountries(cn)
    ClearGrid Sh(SH_COUNTRIES), LAST_ROW_CTRY, 1, CTRY_COLS
    If IsArray(rows) Then
        ReDim out(1 To UBound(rows, 1), 1 To CTRY_COLS)
        For r = 1 To UBound(rows, 1)
            out(r, 1) = rows(r, 1): out(r, 2) = rows(r, 2): out(r, 3) = rows(r, 3)
            out(r, 4) = rows(r, 4): out(r, 5) = rows(r, 5): out(r, 6) = "OK"
        Next r
        PutGrid Sh(SH_COUNTRIES), out
    End If

    ' Currencies with their latest rate
    rows = PerfRepo.SelectRates(cn)
    ClearGrid Sh(SH_RATES), LAST_ROW_RATE, 1, RATE_COLS
    If IsArray(rows) Then
        ReDim out(1 To UBound(rows, 1), 1 To RATE_COLS)
        For r = 1 To UBound(rows, 1)
            out(r, 1) = rows(r, 1): out(r, 2) = rows(r, 2): out(r, 3) = rows(r, 3)
            out(r, 4) = rows(r, 4)
            If IsEmpty(rows(r, 3)) Then out(r, 5) = "NO RATE - USD figures exclude this currency" Else out(r, 5) = "OK"
        Next r
        PutGrid Sh(SH_RATES), out
    End If

    ' Treaty types (A:B) and classes (D:E)
    ClearGrid Sh(SH_LISTS), LAST_ROW_LIST, 1, 2
    ClearGrid Sh(SH_LISTS), LAST_ROW_LIST, 4, 5
    rows = PerfRepo.SelectTreatyTypes(cn)
    If IsArray(rows) Then PutGrid Sh(SH_LISTS), rows, 1
    rows = PerfRepo.SelectCobs(cn)
    If IsArray(rows) Then PutGrid Sh(SH_LISTS), rows, 4

    ' Cedants
    rows = PerfRepo.SelectCedants(cn)
    ClearGrid Sh(SH_CEDANTS), LAST_ROW_CED, 1, CED_COLS
    If IsArray(rows) Then
        nK = UBound(rows, 1)
        ReDim out(1 To nK, 1 To CED_COLS)
        For r = 1 To nK
            out(r, 1) = rows(r, 1): out(r, 2) = rows(r, 2)
            If CBool(rows(r, 3)) Then out(r, 3) = "Y" Else out(r, 3) = "N"
            out(r, 4) = rows(r, 4): out(r, 5) = "OK"
        Next r
        PutGrid Sh(SH_CEDANTS), out
    End If

    PerfDb.CloseQuietly cn
    SetConfigValue "cfg_LastPull", Now
    PullMasters = nK & " cedant(s) and reference lists loaded from Access."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    PerfDb.CloseQuietly cn
    Err.Raise vbObjectError + 1420, APP_NAME, d
End Function

Private Sub ClearGrid(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal c1 As Long, ByVal c2 As Long)
    ws.Range(ws.Cells(FIRST_ROW, c1), ws.Cells(lastRow, c2)).ClearContents
End Sub

Private Sub PutGrid(ByVal ws As Worksheet, ByVal arr As Variant, Optional ByVal firstCol As Long = 1)
    ws.Range(ws.Cells(FIRST_ROW, firstCol), _
             ws.Cells(FIRST_ROW + UBound(arr, 1) - 1, firstCol + UBound(arr, 2) - 1)).Value = arr
End Sub


' ================================ add cedant ===================================
' Adds one cedant straight to Access and appends it to the Cedants sheet, so it
' is in the dropdowns at once. Unsaved edits elsewhere on the sheet are kept.
Public Function AddCedant(ByVal nm As String, ByVal countryText As String) As String
    Dim cn As Object, inTx As Boolean, hc As Variant, hcName As Variant, newId As Long
    Dim byName As Object, byCode As Object, ws As Worksheet, r As Long, freeRow As Long

    nm = Trim$(nm): countryText = Trim$(countryText)
    If Len(nm) = 0 Then Err.Raise vbObjectError + 1700, APP_NAME, "Cedant name is required."
    If Len(countryText) = 0 Then Err.Raise vbObjectError + 1706, APP_NAME, "Home country is required."
    If Len(nm) > 150 Then Err.Raise vbObjectError + 1701, APP_NAME, "Cedant name is longer than 150 characters."

    Set ws = Sh(SH_CEDANTS)
    For r = FIRST_ROW To LAST_ROW_CED
        If Len(Txt(ws.Cells(r, 1).Value)) = 0 And Len(Txt(ws.Cells(r, 4).Value)) = 0 Then freeRow = r: Exit For
    Next r
    If freeRow = 0 Then Err.Raise vbObjectError + 1702, APP_NAME, "The Cedants sheet is full (" & _
        (LAST_ROW_CED - FIRST_ROW + 1) & " rows)."

    On Error GoTo EH
    Set cn = PerfDb.Connect()
    If PerfRepo.CedantMap(cn).Exists(nm) Then
        Err.Raise vbObjectError + 1703, APP_NAME, "A cedant called '" & nm & "' already exists."
    End If
    hc = Null: hcName = Empty
    If Len(countryText) > 0 Then
        Set byName = PerfRepo.CountryMap(cn)
        Set byCode = PerfRepo.CountryCodeMap(cn)
        If byName.Exists(countryText) Then
            hc = byName(countryText)
        ElseIf byCode.Exists(countryText) Then
            hc = byCode(countryText)
        Else
            Err.Raise vbObjectError + 1704, APP_NAME, "Country '" & countryText & "' is not in the database." & _
                vbCrLf & "Use a name or ISO2 code from the Countries sheet, or add the country there and Push masters."
        End If
        hcName = PerfDb.Scalar(cn, "SELECT country_name FROM ref_country WHERE country_id = ?", CLng(hc))
    End If
    cn.BeginTrans
    inTx = True
    newId = PerfRepo.InsertCedant(cn, nm, hc, True)
    PerfRepo.WriteAudit cn, "cedant", newId, "CEDANT_CREATED", nm
    cn.CommitTrans
    inTx = False
    PerfDb.CloseQuietly cn

    ws.Cells(freeRow, 1).Value = nm
    If Not IsEmpty(hcName) And Not IsNull(hcName) Then ws.Cells(freeRow, 2).Value = hcName
    ws.Cells(freeRow, 3).Value = "Y"
    ws.Cells(freeRow, 4).Value = newId
    ws.Cells(freeRow, 5).Value = "OK"
    AddCedant = "Added cedant '" & nm & "' (ID " & newId & ")." & vbCrLf & _
        "It is now in the Cedant dropdown on the Treaties sheet."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    If inTx Then PerfDb.RollbackQuietly cn
    PerfDb.CloseQuietly cn
    Err.Raise vbObjectError + 1705, APP_NAME, d
End Function


' ================================= treaties ====================================
Public Function TreatyFingerprint(ByVal vals As Variant, ByVal i As Long) As String
    Dim p(T_CEDANT To T_NOTES) As String, c As Long
    For c = T_CEDANT To T_NOTES
        Select Case c
            Case T_REF: p(c) = RefText(vals(i, c))
            Case T_CCY: p(c) = Norm(vals(i, c), upper:=True)
            Case T_START, T_END: p(c) = Norm(vals(i, c), asDate:=True)
            Case Else: p(c) = Norm(vals(i, c))
        End Select
    Next c
    TreatyFingerprint = Join(p, "|")
End Function

' Validates one Treaties row; builds the record and the treaty key. "" = valid.
Private Function ValidateTreaty(ByVal vals As Variant, ByVal i As Long, ByVal ced As Object, ByVal tt As Object, _
        ByVal cob As Object, ByVal ccy As Object, ByRef rec As Variant, ByRef key As String) As String
    Dim r(0 To 10) As Variant, s As String, c As Long, m As String, n As Long
    For c = T_CEDANT To T_NOTES
        If IsError(vals(i, c)) Then ValidateTreaty = "a cell contains an error value": Exit Function
    Next c
    s = Txt(vals(i, T_CEDANT))
    If Len(s) = 0 Then ValidateTreaty = "Cedant is required": Exit Function
    If Not ced.Exists(s) Then ValidateTreaty = "Cedant '" & s & "' is not in the database (add it on Cedants)": Exit Function
    r(PerfRepo.TR_CEDANT) = ced(s)
    s = Txt(vals(i, T_TYPE))
    If Len(s) = 0 Then ValidateTreaty = "Treaty type is required": Exit Function
    If Not tt.Exists(s) Then ValidateTreaty = "Treaty type '" & s & "' is not in the database": Exit Function
    r(PerfRepo.TR_TYPE) = tt(s)
    s = Txt(vals(i, T_CLASS))
    If Len(s) = 0 Then ValidateTreaty = "Class of business is required": Exit Function
    If Not cob.Exists(s) Then ValidateTreaty = "Class '" & s & "' is not in the database": Exit Function
    r(PerfRepo.TR_COB) = cob(s)
    s = UCase$(Txt(vals(i, T_CCY)))
    If Len(s) = 0 Then ValidateTreaty = "Currency is required": Exit Function
    If Not ccy.Exists(s) Then ValidateTreaty = "Currency '" & s & "' is not in the database (add it on Rates)": Exit Function
    r(PerfRepo.TR_CCY) = s
    s = RefText(vals(i, T_REF))
    If Len(s) > 50 Then ValidateTreaty = "Reference is longer than 50 characters": Exit Function
    r(PerfRepo.TR_REF) = s
    If IsBlankV(vals(i, T_START)) Or Not IsDate(vals(i, T_START)) Then ValidateTreaty = "Stats start date is required": Exit Function
    If IsBlankV(vals(i, T_END)) Or Not IsDate(vals(i, T_END)) Then ValidateTreaty = "Stats end date is required": Exit Function
    r(PerfRepo.TR_START) = DateValue(CDate(vals(i, T_START)))
    r(PerfRepo.TR_END) = DateValue(CDate(vals(i, T_END)))
    If r(PerfRepo.TR_END) <= r(PerfRepo.TR_START) Then ValidateTreaty = "End date must be after the start date": Exit Function
    If Year(r(PerfRepo.TR_START)) < MIN_UW_YEAR Or Year(r(PerfRepo.TR_START)) > MAX_UW_YEAR Then
        ValidateTreaty = "Start date year must be between " & MIN_UW_YEAR & " and " & MAX_UW_YEAR: Exit Function
    End If
    n = PeriodYears(r(PerfRepo.TR_START), r(PerfRepo.TR_END))
    If n > MAX_TREATY_YEARS Then ValidateTreaty = "Period covers " & n & " years; the maximum is " & MAX_TREATY_YEARS: Exit Function
    m = ShareValue(vals(i, T_SHARE), "Our share", r(PerfRepo.TR_SHARE)): If Len(m) > 0 Then ValidateTreaty = m: Exit Function
    m = ShareValue(vals(i, T_MIN), "Min share", r(PerfRepo.TR_MIN)): If Len(m) > 0 Then ValidateTreaty = m: Exit Function
    m = ShareValue(vals(i, T_MAX), "Max share", r(PerfRepo.TR_MAX)): If Len(m) > 0 Then ValidateTreaty = m: Exit Function
    If Not IsNull(r(PerfRepo.TR_MIN)) And Not IsNull(r(PerfRepo.TR_MAX)) Then
        If r(PerfRepo.TR_MIN) > r(PerfRepo.TR_MAX) Then ValidateTreaty = "Min share is above max share": Exit Function
    End If
    r(PerfRepo.TR_NOTES) = NullIfBlank(Left$(Txt(vals(i, T_NOTES)), 255))
    key = PerfRepo.TreatyKey(vals(i, T_CEDANT), vals(i, T_TYPE), vals(i, T_CLASS), r(PerfRepo.TR_CCY), r(PerfRepo.TR_REF))
    rec = r
End Function

Public Function PullTreaties() As String
    Dim cn As Object, rows As Variant, n As Long
    On Error GoTo EH
    Set cn = PerfDb.Connect()
    rows = PerfRepo.SelectTreaties(cn)
    PerfDb.CloseQuietly cn
    n = WriteSynced(Sh(SH_TREATIES), LAST_ROW_TRT, rows, T_NOTES, T_VERSION, T_FP, True)
    PullTreaties = n & " treaty(ies) loaded."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    PerfDb.CloseQuietly cn
    Err.Raise vbObjectError + 2100, APP_NAME, d
End Function

' Writes pulled rows (B .. lastDataCol, then version) to a synced sheet and
' fingerprints them. Returns the row count.
Private Function WriteSynced(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal rows As Variant, _
        ByVal lastDataCol As Long, ByVal verCol As Long, ByVal fpCol As Long, ByVal isTreaty As Boolean) As Long
    Dim nr As Long, r As Long, c As Long, nData As Long, out() As Variant, vals As Variant, fp() As Variant
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastRow, fpCol)).ClearContents
    If Not IsArray(rows) Then Exit Function
    nr = UBound(rows, 1)
    If nr > lastRow - FIRST_ROW + 1 Then
        Err.Raise vbObjectError + 2101, APP_NAME, "Access holds " & nr & " rows for " & ws.Name & " but the sheet holds " & _
            (lastRow - FIRST_ROW + 1) & ". Report from the Access views (vw_perf_*) or extend the sheet."
    End If
    nData = lastDataCol - 2 + 1               ' B .. lastDataCol
    ReDim out(1 To nr, 1 To fpCol)
    For r = 1 To nr
        out(r, 1) = "OK"
        For c = 1 To nData
            out(r, 1 + c) = rows(r, c)
        Next c
        out(r, verCol) = rows(r, nData + 1)
    Next r
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + nr - 1, fpCol)).Value = out
    vals = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + nr - 1, fpCol)).Value
    ReDim fp(1 To nr, 1 To 1)
    For r = 1 To nr
        If isTreaty Then fp(r, 1) = TreatyFingerprint(vals, r) Else fp(r, 1) = StatFingerprint(vals, r)
    Next r
    ws.Range(ws.Cells(FIRST_ROW, fpCol), ws.Cells(FIRST_ROW + nr - 1, fpCol)).Value = fp
    WriteSynced = nr
End Function

Public Function PushTreaties() As String
    Dim ws As Worksheet, vals As Variant, n As Long, i As Long, pass As Long
    Dim cn As Object, inTx As Boolean, curRow As Long
    Dim ced As Object, tt As Object, cob As Object, ccy As Object, keys As Object
    Dim stat() As Variant, act() As Long, recs() As Variant, rec As Variant, key As String
    Dim msg As String, errCount As Long, isDel As Boolean, hasId As Boolean
    Dim nIns As Long, nUpd As Long, nDel As Long, nSame As Long, id As Long, ver As Long, k As Long
    Dim yr As Variant, y1 As Long, y2 As Long

    Set ws = Sh(SH_TREATIES)
    n = LAST_ROW_TRT - FIRST_ROW + 1
    vals = GridValues(ws, LAST_ROW_TRT, T_INPUT_COLS)
    ReDim stat(1 To n, 1 To 1): ReDim act(1 To n): ReDim recs(1 To n)

    On Error GoTo EH
    Set cn = PerfDb.Connect()
    Set ced = PerfRepo.CedantMap(cn): Set tt = PerfRepo.TreatyTypeMap(cn)
    Set cob = PerfRepo.CobMap(cn): Set ccy = PerfRepo.CurrencyMap(cn)
    Set keys = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        If AnyContent(vals, i, T_ID, T_DELETE) Then
            isDel = (UCase$(Txt(vals(i, T_DELETE))) = "X")
            hasId = Not IsBlankV(vals(i, T_ID))
            If hasId And (Not IsNum(vals(i, T_ID)) Or Not IsNum(vals(i, T_VERSION))) Then
                stat(i, 1) = "ERROR: ID / version were edited - Pull again": errCount = errCount + 1
            ElseIf isDel Then
                If hasId Then act(i) = A_DEL Else stat(i, 1) = "SKIPPED: marked for delete but never saved - clear the row"
            Else
                msg = ValidateTreaty(vals, i, ced, tt, cob, ccy, rec, key)
                If Len(msg) = 0 And keys.Exists(key) Then msg = "same cedant / type / class / currency / reference as row " & keys(key)
                If Len(msg) > 0 Then
                    stat(i, 1) = "ERROR: " & msg: errCount = errCount + 1
                Else
                    keys.Add key, SheetRow(i)
                    recs(i) = rec
                    If Not hasId Then
                        act(i) = A_INS
                    ElseIf TreatyFingerprint(vals, i) <> Txt(vals(i, T_FP)) Then
                        act(i) = A_UPD
                    Else
                        act(i) = A_SAME: stat(i, 1) = "OK"
                    End If
                End If
            End If
        End If
    Next i

    If errCount > 0 Then
        For i = 1 To n
            If act(i) = A_INS Or act(i) = A_UPD Or act(i) = A_DEL Then stat(i, 1) = "NOT SAVED - fix the errors first"
        Next i
        PerfDb.CloseQuietly cn
        WriteStatusCol ws, LAST_ROW_TRT, stat
        PushTreaties = errCount & " treaty row(s) have errors - nothing was saved." & vbCrLf & _
            "Filter column A (Status) on 'ERROR' to see them."
        Exit Function
    End If

    cn.BeginTrans: inTx = True
    For pass = 1 To 3
        For i = 1 To n
            curRow = i
            If act(i) = A_DEL And pass = 1 Then
                id = CLng(vals(i, T_ID)): ver = CLng(vals(i, T_VERSION))
                k = PerfRepo.StatCount(cn, id)
                If k > 0 Then Err.Raise vbObjectError + 2110, APP_NAME, "this treaty has " & k & _
                    " saved stat row(s) - delete them on Stats first (X in Delete), push, then delete the treaty"
                If PerfRepo.DeleteTreaty(cn, id, ver) = 0 Then Err.Raise vbObjectError + 2111, APP_NAME, _
                    "CONFLICT: changed or deleted in Access since your last pull - Pull again"
                PerfRepo.WriteAudit cn, "treaty", id, "TREATY_DELETED", Txt(vals(i, T_FP))
                nDel = nDel + 1
            ElseIf act(i) = A_UPD And pass = 2 Then
                id = CLng(vals(i, T_ID)): ver = CLng(vals(i, T_VERSION))
                yr = PerfRepo.StatYearRange(cn, id)
                If IsArray(yr) Then
                    If Not IsEmpty(yr(1, 1)) Then
                        y1 = Year(recs(i)(PerfRepo.TR_START))
                        y2 = y1 + PeriodYears(recs(i)(PerfRepo.TR_START), recs(i)(PerfRepo.TR_END)) - 1
                        If CLng(yr(1, 1)) < y1 Or CLng(yr(1, 2)) > y2 Then Err.Raise vbObjectError + 2112, APP_NAME, _
                            "saved stats cover UW " & yr(1, 1) & "-" & yr(1, 2) & " but the new period only covers " & _
                            y1 & "-" & y2 & ". Delete those stat rows first or widen the period"
                    End If
                End If
                If PerfRepo.UpdateTreaty(cn, id, ver, recs(i)) = 0 Then Err.Raise vbObjectError + 2113, APP_NAME, _
                    "CONFLICT: changed or deleted in Access since your last pull - Pull again"
                PerfRepo.WriteAudit cn, "treaty", id, "TREATY_UPDATED", TreatyFingerprint(vals, i), "was: " & Txt(vals(i, T_FP))
                nUpd = nUpd + 1
            ElseIf act(i) = A_INS And pass = 3 Then
                id = PerfRepo.InsertTreaty(cn, recs(i))
                PerfRepo.WriteAudit cn, "treaty", id, "TREATY_CREATED", TreatyFingerprint(vals, i)
                nIns = nIns + 1
            ElseIf act(i) = A_SAME And pass = 1 Then
                nSame = nSame + 1
            End If
        Next i
    Next pass
    curRow = 0
    cn.CommitTrans: inTx = False
    PerfDb.CloseQuietly cn
    PullTreaties
    PushTreaties = "Treaties saved: " & nIns & " added, " & nUpd & " updated, " & nDel & " deleted, " & nSame & " unchanged."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    If inTx Then PerfDb.RollbackQuietly cn
    PerfDb.CloseQuietly cn
    If curRow > 0 Then
        For i = 1 To n
            If act(i) = A_INS Or act(i) = A_UPD Or act(i) = A_DEL Then stat(i, 1) = "NOT SAVED (rolled back)"
        Next i
        stat(curRow, 1) = "ERROR: " & d
        WriteStatusCol ws, LAST_ROW_TRT, stat
        Err.Raise vbObjectError + 2114, APP_NAME, "Nothing was saved. Treaties row " & SheetRow(curRow) & ": " & d
    End If
    Err.Raise vbObjectError + 2115, APP_NAME, d
End Function

' =================================== stats =====================================
Public Function StatFingerprint(ByVal vals As Variant, ByVal i As Long) As String
    Dim p(S_CEDANT To S_NOTES) As String, c As Long
    For c = S_CEDANT To S_NOTES
        Select Case c
            Case S_REF: p(c) = RefText(vals(i, c))
            Case S_CCY: p(c) = Norm(vals(i, c), upper:=True)
            Case Else: p(c) = Norm(vals(i, c))
        End Select
    Next c
    StatFingerprint = Join(p, "|")
End Function

Private Function StatKey(ByVal vals As Variant, ByVal i As Long) As String
    StatKey = PerfRepo.TreatyKey(vals(i, S_CEDANT), vals(i, S_TYPE), vals(i, S_CLASS), _
        UCase$(Txt(vals(i, S_CCY))), RefText(vals(i, S_REF)))
End Function

Private Function HasFigures(ByVal vals As Variant, ByVal i As Long) As Boolean
    HasFigures = AnyContent(vals, i, S_PREM, S_BROK)
End Function

' Validates one Stats row against the treaty map. "" = valid.
Private Function ValidateStat(ByVal vals As Variant, ByVal i As Long, ByVal trt As Object, _
                              ByRef rec As Variant, ByRef key As String) As String
    Dim r(0 To 6) As Variant, c As Long, t As Variant, v As Variant, y1 As Long, y2 As Long, cat As String, m As String
    For c = S_CEDANT To S_NOTES
        If IsError(vals(i, c)) Then ValidateStat = "a cell contains an error value": Exit Function
    Next c
    If Len(Txt(vals(i, S_CEDANT))) = 0 Or Len(Txt(vals(i, S_TYPE))) = 0 Or Len(Txt(vals(i, S_CLASS))) = 0 _
       Or Len(Txt(vals(i, S_CCY))) = 0 Then
        ValidateStat = "cedant, treaty type, class and currency are required": Exit Function
    End If
    key = StatKey(vals, i)
    If Not trt.Exists(key) Then
        ValidateStat = "no saved treaty matches this cedant / type / class / currency / reference - " & _
            "add it on Treaties and Push treaties first": Exit Function
    End If
    t = trt(key)                                   ' Array(id, start, end, category)
    r(PerfRepo.SR_TREATY) = t(0)
    v = vals(i, S_YEAR)
    If Not IsNum(v) Then ValidateStat = "UW year is required": Exit Function
    If CDbl(v) <> Int(CDbl(v)) Then ValidateStat = "UW year must be a whole year": Exit Function
    y1 = Year(t(1)): y2 = y1 + PeriodYears(t(1), t(2)) - 1
    If CLng(v) < y1 Or CLng(v) > y2 Then
        ValidateStat = "UW year " & CLng(v) & " is outside the treaty's stats period (" & y1 & "-" & y2 & ")": Exit Function
    End If
    r(PerfRepo.SR_YEAR) = CLng(v)
    If IsBlankV(vals(i, S_PREM)) Then ValidateStat = "Premium is required (enter 0 if none)": Exit Function
    m = AmountValue(vals(i, S_PREM), "Premium", r(PerfRepo.SR_PREM), True): If Len(m) > 0 Then ValidateStat = m: Exit Function
    m = AmountValue(vals(i, S_CLAIMS), "Incurred claims", r(PerfRepo.SR_CLAIMS), True): If Len(m) > 0 Then ValidateStat = m: Exit Function
    m = AmountValue(vals(i, S_COMM), "Commission", r(PerfRepo.SR_COMM), False): If Len(m) > 0 Then ValidateStat = m: Exit Function
    m = AmountValue(vals(i, S_BROK), "Brokerage", r(PerfRepo.SR_BROK), False): If Len(m) > 0 Then ValidateStat = m: Exit Function
    cat = UCase$(Txt(t(3)))
    If cat = NP_CAT Then
        If Not IsNull(r(PerfRepo.SR_COMM)) Then
            If r(PerfRepo.SR_COMM) <> 0 Then ValidateStat = "non-proportional treaty: use Brokerage, not Commission": Exit Function
        End If
    Else
        If Not IsNull(r(PerfRepo.SR_BROK)) Then
            If r(PerfRepo.SR_BROK) <> 0 Then ValidateStat = "proportional treaty: use Commission, not Brokerage": Exit Function
        End If
    End If
    r(PerfRepo.SR_NOTES) = NullIfBlank(Left$(Txt(vals(i, S_NOTES)), 255))
    rec = r
End Function

Public Function PullStats() As String
    Dim cn As Object, rows As Variant, n As Long
    On Error GoTo EH
    Set cn = PerfDb.Connect()
    rows = PerfRepo.SelectStats(cn)
    PerfDb.CloseQuietly cn
    n = WriteSynced(Sh(SH_STATS), LAST_ROW_STAT, rows, S_NOTES, S_VERSION, S_FP, False)
    SetConfigValue "cfg_LastPull", Now
    PullStats = n & " stat row(s) loaded."
    Exit Function
EH:
    Dim d As String: d = Err.Description
    PerfDb.CloseQuietly cn
    Err.Raise vbObjectError + 2200, APP_NAME, d
End Function

Public Function PushStats() As String
    Dim ws As Worksheet, vals As Variant, n As Long, i As Long, pass As Long
    Dim cn As Object, inTx As Boolean, curRow As Long, trt As Object, keys As Object
    Dim stat() As Variant, act() As Long, recs() As Variant, rec As Variant, key As String, yk As String
    Dim msg As String, errCount As Long, isDel As Boolean, hasId As Boolean
    Dim nIns As Long, nUpd As Long, nDel As Long, nSame As Long, nEmpty As Long, id As Long, ver As Long
    Dim empties As Collection, e As Variant

    Set ws = Sh(SH_STATS)
    n = LAST_ROW_STAT - FIRST_ROW + 1
    vals = GridValues(ws, LAST_ROW_STAT, S_INPUT_COLS)
    ReDim stat(1 To n, 1 To 1): ReDim act(1 To n): ReDim recs(1 To n)
    Set empties = New Collection

    On Error GoTo EH
    Set cn = PerfDb.Connect()
    Set trt = PerfRepo.TreatyKeyMap(cn)
    Set keys = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        If AnyContent(vals, i, S_ID, S_DELETE) Then
            isDel = (UCase$(Txt(vals(i, S_DELETE))) = "X")
            hasId = Not IsBlankV(vals(i, S_ID))
            If hasId And (Not IsNum(vals(i, S_ID)) Or Not IsNum(vals(i, S_VERSION))) Then
                stat(i, 1) = "ERROR: ID / version were edited - Pull again": errCount = errCount + 1
            ElseIf isDel Then
                If hasId Then act(i) = A_DEL Else stat(i, 1) = "SKIPPED: marked for delete but never saved - clear the row"
            ElseIf Not hasId And Not HasFigures(vals, i) Then
                ' generated row still waiting for figures: kept on the sheet, not saved
                act(i) = A_EMPTY: stat(i, 1) = EMPTY_STATUS: nEmpty = nEmpty + 1
                empties.Add Array(vals(i, S_CEDANT), vals(i, S_TYPE), vals(i, S_CLASS), vals(i, S_CCY), _
                                  vals(i, S_REF), vals(i, S_YEAR), vals(i, S_NOTES))
            Else
                msg = ValidateStat(vals, i, trt, rec, key)
                If Len(msg) = 0 Then
                    yk = key & "|" & CLng(vals(i, S_YEAR))
                    If keys.Exists(yk) Then msg = "same treaty and UW year as row " & keys(yk)
                End If
                If Len(msg) > 0 Then
                    stat(i, 1) = "ERROR: " & msg: errCount = errCount + 1
                Else
                    keys.Add yk, SheetRow(i)
                    recs(i) = rec
                    If Not hasId Then
                        act(i) = A_INS
                    ElseIf StatFingerprint(vals, i) <> Txt(vals(i, S_FP)) Then
                        act(i) = A_UPD
                    Else
                        act(i) = A_SAME: stat(i, 1) = "OK"
                    End If
                End If
            End If
        End If
    Next i

    If errCount > 0 Then
        For i = 1 To n
            If act(i) = A_INS Or act(i) = A_UPD Or act(i) = A_DEL Then stat(i, 1) = "NOT SAVED - fix the errors first"
        Next i
        PerfDb.CloseQuietly cn
        WriteStatusCol ws, LAST_ROW_STAT, stat
        PushStats = errCount & " stat row(s) have errors - nothing was saved." & vbCrLf & _
            "Filter column A (Status) on 'ERROR' to see them."
        Exit Function
    End If

    cn.BeginTrans: inTx = True
    For pass = 1 To 3
        For i = 1 To n
            curRow = i
            If act(i) = A_DEL And pass = 1 Then
                id = CLng(vals(i, S_ID)): ver = CLng(vals(i, S_VERSION))
                If PerfRepo.DeleteStat(cn, id, ver) = 0 Then Err.Raise vbObjectError + 2210, APP_NAME, _
                    "CONFLICT: changed or deleted in Access since your last pull - Pull again"
                PerfRepo.WriteAudit cn, "treaty_stat", id, "STAT_DELETED", Txt(vals(i, S_FP))
                nDel = nDel + 1
            ElseIf act(i) = A_UPD And pass = 2 Then
                id = CLng(vals(i, S_ID)): ver = CLng(vals(i, S_VERSION))
                If PerfRepo.UpdateStat(cn, id, ver, recs(i)) = 0 Then Err.Raise vbObjectError + 2211, APP_NAME, _
                    "CONFLICT: changed or deleted in Access since your last pull - Pull again"
                PerfRepo.WriteAudit cn, "treaty_stat", id, "STAT_UPDATED", StatFingerprint(vals, i), "was: " & Txt(vals(i, S_FP))
                nUpd = nUpd + 1
            ElseIf act(i) = A_INS And pass = 3 Then
                id = PerfRepo.InsertStat(cn, recs(i))
                PerfRepo.WriteAudit cn, "treaty_stat", id, "STAT_CREATED", StatFingerprint(vals, i)
                nIns = nIns + 1
            ElseIf act(i) = A_SAME And pass = 1 Then
                nSame = nSame + 1
            End If
        Next i
    Next pass
    curRow = 0
    cn.CommitTrans: inTx = False
    PerfDb.CloseQuietly cn

    PullStats
    For Each e In empties                          ' put the still-empty generated rows back
        AppendStatRow e(0), e(1), e(2), e(3), e(4), e(5), e(6)
    Next e
    SetConfigValue "cfg_LastPush", Now
    PushStats = "Stats saved: " & nIns & " added, " & nUpd & " updated, " & nDel & " deleted, " & nSame & " unchanged." & _
        IIf(nEmpty > 0, vbCrLf & nEmpty & " generated row(s) still have no figures - kept on the sheet, not saved.", "")
    Exit Function
EH:
    Dim d As String: d = Err.Description
    If inTx Then PerfDb.RollbackQuietly cn
    PerfDb.CloseQuietly cn
    If curRow > 0 Then
        For i = 1 To n
            If act(i) = A_INS Or act(i) = A_UPD Or act(i) = A_DEL Then stat(i, 1) = "NOT SAVED (rolled back)"
        Next i
        stat(curRow, 1) = "ERROR: " & d
        WriteStatusCol ws, LAST_ROW_STAT, stat
        Err.Raise vbObjectError + 2212, APP_NAME, "Nothing was saved. Stats row " & SheetRow(curRow) & ": " & d
    End If
    Err.Raise vbObjectError + 2213, APP_NAME, d
End Function

' First row on Stats with nothing in B..M.
Private Function NextFreeStatRow(ByVal ws As Worksheet, ByVal vals As Variant, ByVal startAt As Long) As Long
    Dim i As Long
    For i = startAt To UBound(vals, 1)
        If Not AnyContent(vals, i, S_ID, S_NOTES) Then NextFreeStatRow = i: Exit Function
    Next i
End Function

Private Sub AppendStatRow(ByVal ced As Variant, ByVal tt As Variant, ByVal cob As Variant, ByVal ccy As Variant, _
                          ByVal ref As Variant, ByVal yr As Variant, ByVal notes As Variant)
    Dim ws As Worksheet, vals As Variant, i As Long, r As Long
    Set ws = Sh(SH_STATS)
    vals = GridValues(ws, LAST_ROW_STAT, S_INPUT_COLS)
    i = NextFreeStatRow(ws, vals, 1)
    If i = 0 Then Exit Sub
    r = SheetRow(i)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, S_NOTES)).Value = _
        Array(EMPTY_STATUS, Empty, ced, tt, cob, ccy, ref, yr, Empty, Empty, Empty, Empty, notes)
End Sub

' For every treaty on the Treaties sheet, add one Stats row per UW year in its
' stats period that is not on the Stats sheet yet. Works from the sheet, so it
' also covers treaties not pushed yet.
Public Function GenerateStatRows() As String
    Dim wt As Worksheet, ws As Worksheet, tv As Variant, sv As Variant, have As Object
    Dim i As Long, y As Long, y1 As Long, nYears As Long, k As String, ref As String
    Dim out() As Variant, nOut As Long, cap As Long, freeIdx As Long, r As Long, c As Long
    Dim nTrt As Long, nSkip As Long, firstFree As Long

    Set wt = Sh(SH_TREATIES): Set ws = Sh(SH_STATS)
    tv = GridValues(wt, LAST_ROW_TRT, T_INPUT_COLS)
    sv = GridValues(ws, LAST_ROW_STAT, S_INPUT_COLS)
    Set have = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(sv, 1)
        If Len(Txt(sv(i, S_CEDANT))) > 0 And IsNum(sv(i, S_YEAR)) Then
            have(StatKey(sv, i) & "|" & CLng(sv(i, S_YEAR))) = True
        End If
    Next i

    cap = LAST_ROW_STAT - FIRST_ROW + 1
    ReDim out(1 To cap, 1 To S_NOTES)
    For i = 1 To UBound(tv, 1)
        If Len(Txt(tv(i, T_CEDANT))) > 0 And UCase$(Txt(tv(i, T_DELETE))) <> "X" Then
            If IsDate(tv(i, T_START)) And IsDate(tv(i, T_END)) And Not IsBlankV(tv(i, T_START)) And Not IsBlankV(tv(i, T_END)) Then
                If CDate(tv(i, T_END)) > CDate(tv(i, T_START)) Then
                    nTrt = nTrt + 1
                    ref = RefText(tv(i, T_REF))
                    y1 = Year(CDate(tv(i, T_START)))
                    nYears = PeriodYears(CDate(tv(i, T_START)), CDate(tv(i, T_END)))
                    If nYears > MAX_TREATY_YEARS Then nYears = MAX_TREATY_YEARS
                    For y = y1 To y1 + nYears - 1
                        k = PerfRepo.TreatyKey(tv(i, T_CEDANT), tv(i, T_TYPE), tv(i, T_CLASS), UCase$(Txt(tv(i, T_CCY))), ref) & "|" & y
                        If Not have.Exists(k) Then
                            nOut = nOut + 1
                            If nOut > cap Then Exit For
                            out(nOut, S_STATUS) = EMPTY_STATUS
                            out(nOut, S_CEDANT) = Txt(tv(i, T_CEDANT))
                            out(nOut, S_TYPE) = Txt(tv(i, T_TYPE))
                            out(nOut, S_CLASS) = Txt(tv(i, T_CLASS))
                            out(nOut, S_CCY) = UCase$(Txt(tv(i, T_CCY)))
                            out(nOut, S_REF) = ref
                            out(nOut, S_YEAR) = y
                            have.Add k, True
                        End If
                    Next y
                Else
                    nSkip = nSkip + 1
                End If
            Else
                nSkip = nSkip + 1
            End If
        End If
    Next i
    If nOut = 0 Then
        GenerateStatRows = "Nothing to add: every UW year of the " & nTrt & " treaty(ies) already has a row on Stats." & _
            IIf(nSkip > 0, vbCrLf & nSkip & " treaty row(s) skipped - they need a valid start and end date.", "")
        Exit Function
    End If

    ' Write after the last used row (rows are appended, never inserted between saved rows)
    firstFree = 0
    For i = UBound(sv, 1) To 1 Step -1
        If AnyContent(sv, i, S_ID, S_NOTES) Then firstFree = i + 1: Exit For
    Next i
    If firstFree = 0 Then firstFree = 1
    If firstFree + nOut - 1 > cap Then
        Err.Raise vbObjectError + 2300, APP_NAME, "Stats sheet is full: " & nOut & " new rows needed, " & _
            (cap - firstFree + 1) & " free."
    End If
    Dim block() As Variant
    ReDim block(1 To nOut, 1 To S_NOTES)
    For r = 1 To nOut
        For c = 1 To S_NOTES
            block(r, c) = out(r, c)
        Next c
    Next r
    r = SheetRow(firstFree)
    ws.Range(ws.Cells(r, 1), ws.Cells(r + nOut - 1, S_NOTES)).Value = block
    GenerateStatRows = nOut & " stat row(s) added for " & nTrt & " treaty(ies), starting at Stats row " & r & "." & vbCrLf & _
        "Fill in Premium, Incurred claims and Commission (proportional) or Brokerage (non-proportional), then Push stats." & _
        IIf(nSkip > 0, vbCrLf & nSkip & " treaty row(s) skipped - they need a valid start and end date.", "")
End Function


' ================================== setup ======================================
Public Function PullAll() As String
    PullAll = PullMasters() & vbCrLf & PullTreaties() & vbCrLf & PullStats()
End Function

' Create the .accdb if needed, bring the schema up to date, seed Universe lists and,
' for a database with no schema yet, load this workbook's lists, treaties and stats.
Public Function SetupDatabase() As String
    Dim path As String, existed As Boolean, isNew As Boolean, cn As Object, m As String, seedMsg As String
    path = DbPath()
    existed = (Len(Dir$(path)) > 0)
    If Not existed Then PerfDb.CreateDatabaseFile path
    On Error GoTo EH
    Set cn = PerfDb.Connect()
    isNew = Not PerfDb.ObjectExists(cn, "schema_migrations")
    m = PerfMigrate.RunMigrations(cn)
    m = m & vbCrLf & PerfSeed.EnsureReferenceData(cn)
    PerfDb.CloseQuietly cn
    On Error GoTo 0
    If isNew Then
        seedMsg = PushMasters()
        If Left$(seedMsg, 7) = "Nothing" Then
            SetupDatabase = "Database at" & vbCrLf & path & vbCrLf & m & vbCrLf & vbCrLf & seedMsg
            Exit Function
        End If
        seedMsg = seedMsg & vbCrLf & PushTreaties()
        If InStr(seedMsg, "nothing was saved") = 0 Then seedMsg = seedMsg & vbCrLf & PushStats()
    Else
        seedMsg = PullAll()
    End If
    SetupDatabase = IIf(existed, "Connected to", "Created") & vbCrLf & path & vbCrLf & vbCrLf & m & vbCrLf & seedMsg
    Exit Function
EH:
    Dim d As String: d = Err.Description
    PerfDb.CloseQuietly cn
    Err.Raise vbObjectError + 1500, APP_NAME, d
End Function
