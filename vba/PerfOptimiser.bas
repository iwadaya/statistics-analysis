Attribute VB_Name = "PerfOptimiser"
'===============================================================================
' PerfOptimiser - suggests a share for each treaty (Shares sheet, column M).
'
' Objective (USD, at our share, over the UW years chosen on Shares):
'     maximise   mean(annual result) - lambda * stdev(annual result)
'     subject to min share <= share <= max share for every treaty
'                optional: min <= average annual premium <= max
' annual result = sum over treaties of share x (premium - claims - commission - brokerage)
'
' lambda = 0 maximises profit only; a higher lambda prefers a steadier result
' year to year. Method: coordinate search - each treaty's share is tried on a
' grid between its min and max (then refined) while the others are held, and
' passes repeat until no share changes. When a premium limit binds, a price on
' premium (Lagrange multiplier) is found by bisection first so that premium is
' traded between treaties, then the solution is polished against the limit.
' Inputs are read from the Stats and Shares sheets as they are now (100% figures
' in USD).
'===============================================================================
Option Explicit

Private Const GRID_STEPS As Long = 40        ' coarse grid points between min and max
Private Const FINE_STEPS As Long = 20        ' refinement points around the best grid point
Private Const MAX_PASSES As Long = 60
Private Const PENALTY As Double = 100#       ' per USD outside the premium range

Private Const BISECT_STEPS As Long = 30

Private nY As Long, nT As Long
Private lam As Double, pMin As Double, pMax As Double, hasPMin As Boolean, hasPMax As Boolean
Private mu As Double                         ' price on average premium (Lagrange phase)
Private usePen As Boolean                    ' apply the premium-range penalty
Private P() As Double, R() As Double, lo() As Double, hi() As Double, active() As Boolean

Private Function CfgValue(ByVal nm As String) As Variant
    On Error Resume Next
    CfgValue = ThisWorkbook.Names(nm).RefersToRange.Value
End Function

Private Function NumOrBlank(ByVal v As Variant, ByRef out As Double) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(v)) = 0 Then Exit Function
    End If
    If Not IsNumeric(v) Then Exit Function
    out = CDbl(v)
    NumOrBlank = True
End Function

' Objective for annual results A() and premiums Q()
Private Function Objective(ByRef A() As Double, ByRef Q() As Double) As Double
    Dim y As Long, mA As Double, mQ As Double, v As Double, f As Double
    For y = 1 To nY
        mA = mA + A(y): mQ = mQ + Q(y)
    Next y
    mA = mA / nY: mQ = mQ / nY
    For y = 1 To nY
        v = v + (A(y) - mA) ^ 2
    Next y
    f = mA - lam * Sqr(v / nY) - mu * mQ
    If usePen Then
        If hasPMin And mQ < pMin Then f = f - PENALTY * (pMin - mQ)
        If hasPMax And mQ > pMax Then f = f - PENALTY * (mQ - pMax)
    End If
    Objective = f
End Function

' Objective if treaty i's share moves by delta (A, Q left unchanged)
Private Function ObjectiveWith(ByRef A() As Double, ByRef Q() As Double, _
                               ByVal i As Long, ByVal delta As Double, ByRef tA() As Double, ByRef tQ() As Double) As Double
    Dim y As Long
    For y = 1 To nY
        tA(y) = A(y) + delta * R(i, y)
        tQ(y) = Q(y) + delta * P(i, y)
    Next y
    ObjectiveWith = Objective(tA, tQ)
End Function

' Coordinate search from shares s(); returns the average annual premium reached.
Private Function SearchFrom(ByRef s() As Double, ByRef passes As Long) As Double
    Dim A() As Double, Q() As Double, tA() As Double, tQ() As Double
    Dim i As Long, j As Long, y As Long, pass As Long, changed As Boolean
    Dim f As Double, bestF As Double, bestS As Double, cand As Double, fc As Double, stp As Double, mQ As Double
    ReDim A(1 To nY): ReDim Q(1 To nY): ReDim tA(1 To nY): ReDim tQ(1 To nY)
    For i = 1 To nT
        For y = 1 To nY
            A(y) = A(y) + s(i) * R(i, y): Q(y) = Q(y) + s(i) * P(i, y)
        Next y
    Next i
    f = Objective(A, Q)
    For pass = 1 To MAX_PASSES
        changed = False
        For i = 1 To nT
            If active(i) Then
                bestS = s(i): bestF = f
                stp = (hi(i) - lo(i)) / GRID_STEPS
                For j = 0 To GRID_STEPS
                    cand = lo(i) + j * stp
                    fc = ObjectiveWith(A, Q, i, cand - s(i), tA, tQ)
                    If fc > bestF + 0.000000001 * (1 + Abs(bestF)) Then bestF = fc: bestS = cand
                Next j
                For j = -FINE_STEPS To FINE_STEPS          ' refine around the best point
                    cand = bestS + j * stp / FINE_STEPS
                    If cand >= lo(i) And cand <= hi(i) Then
                        fc = ObjectiveWith(A, Q, i, cand - s(i), tA, tQ)
                        If fc > bestF + 0.000000001 * (1 + Abs(bestF)) Then bestF = fc: bestS = cand
                    End If
                Next j
                If bestS <> s(i) Then
                    For y = 1 To nY
                        A(y) = A(y) + (bestS - s(i)) * R(i, y): Q(y) = Q(y) + (bestS - s(i)) * P(i, y)
                    Next y
                    s(i) = bestS: f = bestF: changed = True
                End If
            End If
        Next i
        passes = passes + 1
        If Not changed Then Exit For
    Next pass
    For y = 1 To nY
        mQ = mQ + Q(y)
    Next y
    SearchFrom = mQ / nY
End Function

Private Function Violates(ByVal mQ As Double, ByVal over As Boolean) As Boolean
    If over Then Violates = (mQ > pMax) Else Violates = (mQ < pMin)
End Function

Private Sub CopyShares(ByRef src() As Double, ByRef dst() As Double)
    Dim i As Long
    ReDim dst(1 To nT)
    For i = 1 To nT
        dst(i) = src(i)
    Next i
End Sub

Public Function RunOptimiser() As String
    Dim wsS As Worksheet, wsT As Worksheet, sv As Variant, tv As Variant
    Dim i As Long, j As Long, y As Long, yr As Long, k As Variant, idx As Object
    Dim fromY As Long, toY As Long, d As Double, hasFrom As Boolean, hasTo As Boolean
    Dim s() As Double, s2() As Double, feas() As Double
    Dim nActive As Long, out() As Variant, lowY As Long, highY As Long, passes As Long
    Dim cur As Double, x As Double, msg As String, mQ As Double, over As Boolean, sgn As Double
    Dim muLo As Double, muHi As Double, muMid As Double, bisect As Long, infeasible As Boolean

    Set wsS = ThisWorkbook.Worksheets(SH_SHARES)
    Set wsT = ThisWorkbook.Worksheets(SH_STATS)
    nT = SH_LAST - SH_FIRST + 1
    tv = wsS.Range(wsS.Cells(SH_FIRST, 1), wsS.Cells(SH_LAST, SH_OPT)).Value
    sv = wsT.Range(wsT.Cells(FIRST_ROW, 1), wsT.Cells(LAST_ROW_STAT, S_KEY)).Value

    ' --- settings ---------------------------------------------------------------
    hasFrom = NumOrBlank(CfgValue("opt_FromYear"), d): If hasFrom Then fromY = CLng(d)
    hasTo = NumOrBlank(CfgValue("opt_ToYear"), d): If hasTo Then toY = CLng(d)
    lam = 0: If NumOrBlank(CfgValue("opt_Lambda"), d) Then lam = d
    If lam < 0 Then Err.Raise vbObjectError + 3000, APP_NAME, "Risk weight (lambda) cannot be negative."
    hasPMin = NumOrBlank(CfgValue("opt_PremMin"), pMin)
    hasPMax = NumOrBlank(CfgValue("opt_PremMax"), pMax)
    If hasPMin And hasPMax Then
        If pMin > pMax Then Err.Raise vbObjectError + 3001, APP_NAME, "Minimum premium is above maximum premium."
    End If

    ' --- treaties -----------------------------------------------------------------
    Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare
    ReDim lo(1 To nT): ReDim hi(1 To nT): ReDim s(1 To nT): ReDim active(1 To nT)
    For i = 1 To nT
        k = tv(i, SH_KEY)
        If Not IsError(k) Then
            If Len(Trim$(CStr(k))) > 0 Then
                If Not idx.Exists(CStr(k)) Then idx.Add CStr(k), i
                lo(i) = 0: hi(i) = 0: cur = 0
                If NumOrBlank(tv(i, SH_MIN), x) Then lo(i) = x
                If NumOrBlank(tv(i, SH_MAX), x) Then hi(i) = x
                If NumOrBlank(tv(i, SH_CURRENT), x) Then cur = x
                If lo(i) < 0 Then lo(i) = 0
                If hi(i) > 1 Then hi(i) = 1
                If hi(i) < lo(i) Then hi(i) = lo(i)
                s(i) = cur
                If s(i) < lo(i) Then s(i) = lo(i)
                If s(i) > hi(i) Then s(i) = hi(i)
            End If
        End If
    Next i
    If idx.Count = 0 Then Err.Raise vbObjectError + 3002, APP_NAME, "No treaties on the Shares sheet - add treaties and stats first."

    ' --- year range ---------------------------------------------------------------
    lowY = 99999: highY = -99999
    For j = 1 To UBound(sv, 1)
        If Not IsError(sv(j, S_KEY)) And Not IsError(sv(j, S_YEAR)) Then
            If idx.Exists(CStr(sv(j, S_KEY))) And NumOrBlank(sv(j, S_YEAR), d) Then
                If d < lowY Then lowY = CLng(d)
                If d > highY Then highY = CLng(d)
            End If
        End If
    Next j
    If highY < lowY Then Err.Raise vbObjectError + 3003, APP_NAME, "No stats rows match the treaties on Shares."
    If Not hasFrom Then fromY = lowY
    If Not hasTo Then toY = highY
    If toY < fromY Then Err.Raise vbObjectError + 3004, APP_NAME, "'To' UW year is before 'From' UW year."
    nY = toY - fromY + 1

    ' --- yearly figures per treaty (100%, USD) ------------------------------------------
    ReDim P(1 To nT, 1 To nY): ReDim R(1 To nT, 1 To nY)
    For j = 1 To UBound(sv, 1)
        If Not IsError(sv(j, S_KEY)) And Not IsError(sv(j, S_YEAR)) And Not IsError(sv(j, S_PREM_USD)) _
           And Not IsError(sv(j, S_RESULT_USD)) Then
            If idx.Exists(CStr(sv(j, S_KEY))) And NumOrBlank(sv(j, S_YEAR), d) Then
                yr = CLng(d)
                If yr >= fromY And yr <= toY Then
                    i = idx(CStr(sv(j, S_KEY)))
                    If NumOrBlank(sv(j, S_PREM_USD), x) Then P(i, yr - fromY + 1) = P(i, yr - fromY + 1) + x
                    If NumOrBlank(sv(j, S_RESULT_USD), x) Then R(i, yr - fromY + 1) = R(i, yr - fromY + 1) + x
                End If
            End If
        End If
    Next j
    For i = 1 To nT
        If hi(i) > lo(i) Then
            For y = 1 To nY
                If P(i, y) <> 0 Or R(i, y) <> 0 Then active(i) = True: Exit For
            Next y
        End If
        If active(i) Then nActive = nActive + 1
    Next i
    If nActive = 0 Then Err.Raise vbObjectError + 3005, APP_NAME, _
        "Nothing to optimise: no treaty has both stats in the chosen years and a max share above its min share."

    ' --- 1. unconstrained search ----------------------------------------------------
    Application.StatusBar = APP_NAME & ": optimising..."
    mu = 0: usePen = False
    mQ = SearchFrom(s, passes)

    ' --- 2. premium range binds: find the price on premium by bisection --------------
    If (hasPMax And mQ > pMax) Or (hasPMin And mQ < pMin) Then
        over = (hasPMax And mQ > pMax)
        If over Then sgn = 1 Else sgn = -1
        muLo = 0: muHi = 1
        CopyShares s, s2
        mu = sgn * muHi: mQ = SearchFrom(s2, passes)
        Do While Violates(mQ, over) And muHi < 1000000#
            muLo = muHi: muHi = muHi * 2
            mu = sgn * muHi: mQ = SearchFrom(s2, passes)
        Loop
        infeasible = Violates(mQ, over)
        CopyShares s2, feas
        For bisect = 1 To BISECT_STEPS
            Application.StatusBar = APP_NAME & ": fitting the premium range, step " & bisect & " of " & BISECT_STEPS & "..."
            muMid = (muLo + muHi) / 2
            CopyShares feas, s2
            mu = sgn * muMid: mQ = SearchFrom(s2, passes)
            If Violates(mQ, over) Then
                muLo = muMid
            Else
                muHi = muMid: CopyShares s2, feas
            End If
        Next bisect
        CopyShares feas, s
    End If

    ' --- 3. polish against the premium range -----------------------------------------
    mu = 0: usePen = True
    mQ = SearchFrom(s, passes)

    ' --- write results (rounded to 0.01%) ---------------------------------------------
    ReDim out(1 To nT, 1 To 1)
    For i = 1 To nT
        k = tv(i, SH_KEY)
        If Not IsError(k) Then
            If Len(Trim$(CStr(k))) > 0 Then out(i, 1) = Round(s(i), 4)
        End If
    Next i
    wsS.Range(wsS.Cells(SH_FIRST, SH_OPT), wsS.Cells(SH_LAST, SH_OPT)).Value = out
    msg = "Optimised " & Format$(Now, "dd-mmm-yyyy hh:nn") & " over UW " & fromY & "-" & toY & _
          ", risk weight " & lam & ", " & nActive & " treaty(ies) free to move, " & passes & " search pass(es); " & _
          "average annual premium at our share USD " & Format$(mQ, "#,##0") & "."
    If infeasible Then msg = msg & " The premium range cannot be reached within the min / max shares - " & _
          "this is the closest the limits allow."
    On Error Resume Next
    ThisWorkbook.Names("opt_Status").RefersToRange.Value = msg
    On Error GoTo 0
    Application.StatusBar = False
    RunOptimiser = msg & vbCrLf & vbCrLf & "Suggested shares are in column M of Shares. Compare them with current and " & _
        "scenario shares on the Scenarios sheet."
End Function

' Copy the optimised shares into the Scenario column (J) so you can tweak them.
Public Function UseOptimisedAsScenario() As String
    Dim ws As Worksheet, v As Variant, i As Long, out() As Variant, n As Long, cnt As Long
    Set ws = ThisWorkbook.Worksheets(SH_SHARES)
    v = ws.Range(ws.Cells(SH_FIRST, 1), ws.Cells(SH_LAST, SH_OPT)).Value
    n = UBound(v, 1)
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        If Not IsError(v(i, SH_OPT)) Then
            If Not IsEmpty(v(i, SH_OPT)) Then out(i, 1) = v(i, SH_OPT): cnt = cnt + 1
        End If
    Next i
    If cnt = 0 Then Err.Raise vbObjectError + 3010, APP_NAME, "Run the optimiser first - column M is empty."
    ws.Range(ws.Cells(SH_FIRST, SH_SCENARIO), ws.Cells(SH_LAST, SH_SCENARIO)).Value = out
    UseOptimisedAsScenario = cnt & " optimised share(s) copied into the Scenario column."
End Function

Public Sub ClearScenario()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SHARES)
    ws.Range(ws.Cells(SH_FIRST, SH_SCENARIO), ws.Cells(SH_LAST, SH_SCENARIO)).ClearContents
End Sub
