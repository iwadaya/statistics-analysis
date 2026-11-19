Attribute VB_Name = "PerfConfig"
'===============================================================================
' PerfConfig - constants and configuration for the Treaty Performance Tracker
'
' Layering (same shape as Universe and the Aggregate Tracker):
'   PerfMacros    -> button entry points            (Universe: client screens)
'   PerfService   -> validation + orchestration     (Universe: server/src/services)
'   PerfOptimiser -> share optimiser                (Universe: services/* analytics)
'   PerfRepo      -> all SQL against Access         (Universe: modules/*/repositories)
'   PerfDb        -> connection / command helpers   (Universe: server/src/db/pool.js)
'   PerfMigrate   -> numbered schema migrations     (Universe: db/migrations + runMigrations.js)
'   PerfSeed      -> predefined reference data      (Universe: startup/ensureReferenceData.js)
'   PerfConfig    -> constants                      (Universe: config/env.js)
'===============================================================================
Option Explicit

Public Const APP_NAME As String = "Treaty Performance Tracker"
Public Const DB_FILENAME As String = "TreatyPerformance.accdb"

' --- Sheet names --------------------------------------------------------------
Public Const SH_START As String = "Start"
Public Const SH_TREATIES As String = "Treaties"
Public Const SH_STATS As String = "Stats"
Public Const SH_SHARES As String = "Shares"
Public Const SH_CEDANTS As String = "Cedants"
Public Const SH_COUNTRIES As String = "Countries"
Public Const SH_RATES As String = "Rates"
Public Const SH_LISTS As String = "Lists"

' --- Grid geometry ------------------------------------------------------------
Public Const FIRST_ROW As Long = 6           ' first data row on input / master sheets
Public Const LAST_ROW_TRT As Long = 405      ' 400 treaties
Public Const LAST_ROW_STAT As Long = 6005    ' 6,000 yearly stat rows
Public Const LAST_ROW_CED As Long = 505      ' 500 cedants
Public Const LAST_ROW_CTRY As Long = 205     ' 200 countries
Public Const LAST_ROW_RATE As Long = 105     ' 100 currencies
Public Const LAST_ROW_LIST As Long = 105     ' 100 treaty types / classes

' --- Treaties sheet: A..P inputs + sync state, Q.. formulas ---------------------
Public Const T_STATUS As Long = 1
Public Const T_ID As Long = 2
Public Const T_CEDANT As Long = 3
Public Const T_TYPE As Long = 4
Public Const T_CLASS As Long = 5
Public Const T_CCY As Long = 6
Public Const T_REF As Long = 7
Public Const T_START As Long = 8
Public Const T_END As Long = 9
Public Const T_SHARE As Long = 10
Public Const T_MIN As Long = 11
Public Const T_MAX As Long = 12
Public Const T_NOTES As Long = 13
Public Const T_DELETE As Long = 14
Public Const T_VERSION As Long = 15
Public Const T_FP As Long = 16
Public Const T_INPUT_COLS As Long = 16

' --- Stats sheet: A..P inputs + sync state, Q.. formulas -----------------------
Public Const S_STATUS As Long = 1
Public Const S_ID As Long = 2
Public Const S_CEDANT As Long = 3
Public Const S_TYPE As Long = 4
Public Const S_CLASS As Long = 5
Public Const S_CCY As Long = 6
Public Const S_REF As Long = 7
Public Const S_YEAR As Long = 8
Public Const S_PREM As Long = 9
Public Const S_CLAIMS As Long = 10
Public Const S_COMM As Long = 11
Public Const S_BROK As Long = 12
Public Const S_NOTES As Long = 13
Public Const S_DELETE As Long = 14
Public Const S_VERSION As Long = 15
Public Const S_FP As Long = 16
Public Const S_INPUT_COLS As Long = 16
' calculated columns the optimiser reads
Public Const S_PREM_USD As Long = 20         ' T
Public Const S_RESULT_USD As Long = 23       ' W
Public Const S_KEY As Long = 24              ' X

' --- Shares sheet --------------------------------------------------------------
Public Const SH_FIRST As Long = 10           ' first treaty row on Shares
Public Const SH_LAST As Long = 409
Public Const SH_KEY As Long = 8              ' H  treaty key
Public Const SH_CURRENT As Long = 9          ' I  current share
Public Const SH_SCENARIO As Long = 10        ' J  scenario share (input)
Public Const SH_MIN As Long = 11             ' K  effective min
Public Const SH_MAX As Long = 12             ' L  effective max
Public Const SH_OPT As Long = 13             ' M  optimised share (written by the optimiser)

' --- Master sheet widths --------------------------------------------------------
Public Const CED_COLS As Long = 5            ' name | home country | active | id | status
Public Const CTRY_COLS As Long = 6            ' name | ISO2 | ISO3 | region | id | status
Public Const RATE_COLS As Long = 5            ' code | name | USD per 1 | effective date | status

' --- Defaults -----------------------------------------------------------------
Public Const DEFAULT_REF As String = "MAIN"  ' blank treaty reference
Public Const MIN_UW_YEAR As Long = 1980
Public Const MAX_UW_YEAR As Long = 2100
Public Const MAX_TREATY_YEARS As Long = 50
Public Const PROP_CAT As String = "PROPORTIONAL"
Public Const NP_CAT As String = "NON_PROPORTIONAL"

' Path to the Access database. Start!C7 (named cfg_DbPath) overrides the default
' of "same folder as this workbook".
Public Function DbPath() As String
    Dim p As String
    On Error Resume Next
    p = Trim$(CStr(ThisWorkbook.Names("cfg_DbPath").RefersToRange.Value))
    On Error GoTo 0
    If Len(p) = 0 Then
        If Len(ThisWorkbook.Path) = 0 Then
            Err.Raise vbObjectError + 1000, APP_NAME, _
                "Save the workbook first, or enter a database path in Start!C7."
        End If
        p = ThisWorkbook.Path & Application.PathSeparator & DB_FILENAME
    End If
    If LCase$(Left$(p, 4)) = "http" Then
        Err.Raise vbObjectError + 1001, APP_NAME, _
            "The workbook is opened from OneDrive/SharePoint (" & p & ")." & vbCrLf & _
            "Access needs a local or network-drive path. Enter one in Start!C7, e.g. C:\Data\" & DB_FILENAME
    End If
    DbPath = p
End Function

Public Function Actor() As String
    Dim u As String
    u = Environ$("USERNAME")
    If Len(u) = 0 Then u = Application.UserName
    If Len(u) = 0 Then u = "EXCEL"
    Actor = Left$(u, 100)
End Function

Public Sub SetConfigValue(ByVal nm As String, ByVal v As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Value = v
End Sub

' Number of underwriting years covered by a stats period: 01/01/2016-31/12/2025
' = 10, 01/07/2016-30/06/2025 = 9. UW years run from Year(start) onwards.
Public Function PeriodYears(ByVal startDate As Date, ByVal endDate As Date) As Long
    Dim n As Long
    n = CLng(Round((CDbl(endDate) - CDbl(startDate) + 1) / 365.25, 0))
    If n < 1 Then n = 1
    PeriodYears = n
End Function
