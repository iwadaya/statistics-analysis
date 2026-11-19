Attribute VB_Name = "PerfRepo"
'===============================================================================
' PerfRepo - every SQL statement against the Access database lives here
' (Universe: modules/*/repositories). No validation and no sheet access.
'===============================================================================
Option Explicit

' Treaty record array positions (PerfService.BuildTreaty)
Public Const TR_CEDANT As Long = 0
Public Const TR_TYPE As Long = 1
Public Const TR_COB As Long = 2
Public Const TR_CCY As Long = 3
Public Const TR_REF As Long = 4
Public Const TR_START As Long = 5
Public Const TR_END As Long = 6
Public Const TR_SHARE As Long = 7
Public Const TR_MIN As Long = 8
Public Const TR_MAX As Long = 9
Public Const TR_NOTES As Long = 10

' Stat record array positions
Public Const SR_TREATY As Long = 0
Public Const SR_YEAR As Long = 1
Public Const SR_PREM As Long = 2
Public Const SR_CLAIMS As Long = 3
Public Const SR_COMM As Long = 4
Public Const SR_BROK As Long = 5
Public Const SR_NOTES As Long = 6

' --- Lookups (name, case-insensitive -> id) ------------------------------------
Public Function NameIdMap(ByVal cn As Object, ByVal tableName As String, _
                          ByVal nameCol As String, ByVal idCol As String) As Object
    Dim d As Object, rows As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    rows = PerfDb.QueryRows(cn, "SELECT " & nameCol & ", " & idCol & " FROM " & tableName)
    If IsArray(rows) Then
        For i = 1 To UBound(rows, 1)
            d(Trim$(CStr(rows(i, 1)))) = rows(i, 2)
        Next i
    End If
    Set NameIdMap = d
End Function

Public Function CedantMap(ByVal cn As Object) As Object
    Set CedantMap = NameIdMap(cn, "cedant", "cedant_name", "cedant_id")
End Function
Public Function CountryMap(ByVal cn As Object) As Object
    Set CountryMap = NameIdMap(cn, "ref_country", "country_name", "country_id")
End Function
Public Function CountryCodeMap(ByVal cn As Object) As Object
    Set CountryCodeMap = NameIdMap(cn, "ref_country", "country_code", "country_id")
End Function
Public Function TreatyTypeMap(ByVal cn As Object) As Object
    Set TreatyTypeMap = NameIdMap(cn, "ref_treaty_type", "treaty_type_name", "treaty_type_id")
End Function
' treaty type name -> category (PROPORTIONAL / NON_PROPORTIONAL)
Public Function TreatyCategoryMap(ByVal cn As Object) As Object
    Set TreatyCategoryMap = NameIdMap(cn, "ref_treaty_type", "treaty_type_name", "treaty_category")
End Function
Public Function CobMap(ByVal cn As Object) As Object
    Set CobMap = NameIdMap(cn, "ref_class_of_business", "cob_name", "cob_id")
End Function
Public Function CurrencyMap(ByVal cn As Object) As Object
    Set CurrencyMap = NameIdMap(cn, "ref_currency", "currency_code", "currency_name")
End Function

' "cedant|type|class|CCY|REF" (names, upper-cased) -> treaty row: Array(id, startDate, endDate, category)
Public Function TreatyKeyMap(ByVal cn As Object) As Object
    Dim d As Object, rows As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    rows = PerfDb.QueryRows(cn, "SELECT cedant_name, treaty_type_name, cob_name, currency_code, treaty_ref, " & _
        "treaty_id, stats_start, stats_end, treaty_category FROM vw_treaty_detail")
    If IsArray(rows) Then
        For i = 1 To UBound(rows, 1)
            d(TreatyKey(rows(i, 1), rows(i, 2), rows(i, 3), rows(i, 4), rows(i, 5))) = _
                Array(rows(i, 6), rows(i, 7), rows(i, 8), rows(i, 9))
        Next i
    End If
    Set TreatyKeyMap = d
End Function

Public Function TreatyKey(ByVal ced As Variant, ByVal tt As Variant, ByVal cob As Variant, _
                          ByVal ccy As Variant, ByVal ref As Variant) As String
    TreatyKey = UCase$(Trim$(CStr(ced)) & "|" & Trim$(CStr(tt)) & "|" & Trim$(CStr(cob)) & "|" & _
                Trim$(CStr(ccy)) & "|" & Trim$(CStr(ref)))
End Function

' --- Audit ---------------------------------------------------------------------
Public Sub WriteAudit(ByVal cn As Object, ByVal entityType As String, ByVal entityId As Long, _
                      ByVal eventType As String, ByVal payload As String, Optional ByVal comment As String = "")
    Dim c As Variant
    If Len(comment) = 0 Then c = Null Else c = Left$(comment, 255)
    PerfDb.ExecParams cn, "INSERT INTO audit_log (entity_type, entity_id, event_type, actor, payload, audit_comment, created_at) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?)", entityType, entityId, eventType, Actor(), payload, c, Now
End Sub

' --- Treaties ------------------------------------------------------------------
' Column order = Treaties sheet B (ID) .. M (Notes), then row_version.
Public Function SelectTreaties(ByVal cn As Object) As Variant
    SelectTreaties = PerfDb.QueryRows(cn, _
        "SELECT treaty_id, cedant_name, treaty_type_name, cob_name, currency_code, treaty_ref, stats_start, " & _
        "stats_end, our_share, min_share, max_share, treaty_notes, row_version FROM vw_treaty_detail " & _
        "ORDER BY cedant_name, treaty_type_name, cob_name, currency_code, treaty_ref")
End Function

Public Function InsertTreaty(ByVal cn As Object, ByVal r As Variant) As Long
    Dim ts As Date, who As String
    ts = Now: who = Actor()
    PerfDb.ExecParams cn, "INSERT INTO treaty (cedant_id, treaty_type_id, cob_id, currency_code, treaty_ref, " & _
        "stats_start, stats_end, our_share, min_share, max_share, treaty_notes, row_version, created_at, created_by, " & _
        "updated_at, updated_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", _
        r(TR_CEDANT), r(TR_TYPE), r(TR_COB), r(TR_CCY), r(TR_REF), r(TR_START), r(TR_END), r(TR_SHARE), _
        r(TR_MIN), r(TR_MAX), r(TR_NOTES), 1&, ts, who, ts, who
    InsertTreaty = PerfDb.LastId(cn)
End Function

Public Function UpdateTreaty(ByVal cn As Object, ByVal id As Long, ByVal version As Long, ByVal r As Variant) As Long
    UpdateTreaty = PerfDb.ExecParams(cn, "UPDATE treaty SET cedant_id = ?, treaty_type_id = ?, cob_id = ?, " & _
        "currency_code = ?, treaty_ref = ?, stats_start = ?, stats_end = ?, our_share = ?, min_share = ?, " & _
        "max_share = ?, treaty_notes = ?, row_version = row_version + 1, updated_at = ?, updated_by = ? " & _
        "WHERE treaty_id = ? AND row_version = ?", _
        r(TR_CEDANT), r(TR_TYPE), r(TR_COB), r(TR_CCY), r(TR_REF), r(TR_START), r(TR_END), r(TR_SHARE), _
        r(TR_MIN), r(TR_MAX), r(TR_NOTES), Now, Actor(), id, version)
End Function

Public Function DeleteTreaty(ByVal cn As Object, ByVal id As Long, ByVal version As Long) As Long
    DeleteTreaty = PerfDb.ExecParams(cn, "DELETE FROM treaty WHERE treaty_id = ? AND row_version = ?", id, version)
End Function

Public Function StatCount(ByVal cn As Object, ByVal treatyId As Long) As Long
    StatCount = CLng(PerfDb.Scalar(cn, "SELECT COUNT(*) FROM treaty_stat WHERE treaty_id = ?", treatyId))
End Function

' Stat years that fall outside a treaty's new period (blocks shrinking the period under saved stats)
Public Function StatYearRange(ByVal cn As Object, ByVal treatyId As Long) As Variant
    StatYearRange = PerfDb.QueryRows(cn, "SELECT MIN(uw_year), MAX(uw_year) FROM treaty_stat WHERE treaty_id = ?", treatyId)
End Function

' --- Stats ---------------------------------------------------------------------
' Column order = Stats sheet B (ID) .. M (Notes), then row_version.
Public Function SelectStats(ByVal cn As Object) As Variant
    SelectStats = PerfDb.QueryRows(cn, _
        "SELECT stat_id, cedant_name, treaty_type_name, cob_name, currency_code, treaty_ref, uw_year, premium, " & _
        "incurred_claims, commission, brokerage, stat_notes, row_version FROM vw_stat_detail " & _
        "ORDER BY cedant_name, treaty_type_name, cob_name, currency_code, treaty_ref, uw_year")
End Function

Public Function InsertStat(ByVal cn As Object, ByVal r As Variant) As Long
    Dim ts As Date, who As String
    ts = Now: who = Actor()
    PerfDb.ExecParams cn, "INSERT INTO treaty_stat (treaty_id, uw_year, premium, incurred_claims, commission, " & _
        "brokerage, stat_notes, row_version, created_at, created_by, updated_at, updated_by) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", _
        r(SR_TREATY), r(SR_YEAR), r(SR_PREM), r(SR_CLAIMS), r(SR_COMM), r(SR_BROK), r(SR_NOTES), 1&, ts, who, ts, who
    InsertStat = PerfDb.LastId(cn)
End Function

Public Function UpdateStat(ByVal cn As Object, ByVal id As Long, ByVal version As Long, ByVal r As Variant) As Long
    UpdateStat = PerfDb.ExecParams(cn, "UPDATE treaty_stat SET treaty_id = ?, uw_year = ?, premium = ?, " & _
        "incurred_claims = ?, commission = ?, brokerage = ?, stat_notes = ?, row_version = row_version + 1, " & _
        "updated_at = ?, updated_by = ? WHERE stat_id = ? AND row_version = ?", _
        r(SR_TREATY), r(SR_YEAR), r(SR_PREM), r(SR_CLAIMS), r(SR_COMM), r(SR_BROK), r(SR_NOTES), Now, Actor(), id, version)
End Function

Public Function DeleteStat(ByVal cn As Object, ByVal id As Long, ByVal version As Long) As Long
    DeleteStat = PerfDb.ExecParams(cn, "DELETE FROM treaty_stat WHERE stat_id = ? AND row_version = ?", id, version)
End Function

' --- Cedants -------------------------------------------------------------------
Public Function SelectCedants(ByVal cn As Object) As Variant
    SelectCedants = PerfDb.QueryRows(cn, _
        "SELECT c.cedant_name, k.country_name, c.is_active, c.cedant_id " & _
        "FROM cedant AS c INNER JOIN ref_country AS k ON c.home_country_id = k.country_id " & _
        "ORDER BY c.cedant_name")
End Function

Public Function InsertCedant(ByVal cn As Object, ByVal nm As String, ByVal homeCountryId As Variant, _
                             ByVal active As Boolean) As Long
    Dim ts As Date: ts = Now
    PerfDb.ExecParams cn, "INSERT INTO cedant (cedant_name, home_country_id, is_active, created_at, created_by, updated_at, updated_by) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?)", nm, homeCountryId, active, ts, Actor(), ts, Actor()
    InsertCedant = PerfDb.LastId(cn)
End Function

Public Function UpdateCedant(ByVal cn As Object, ByVal cedantId As Long, ByVal nm As String, _
                             ByVal homeCountryId As Variant, ByVal active As Boolean) As Long
    UpdateCedant = PerfDb.ExecParams(cn, "UPDATE cedant SET cedant_name = ?, home_country_id = ?, is_active = ?, " & _
        "updated_at = ?, updated_by = ? WHERE cedant_id = ?", nm, homeCountryId, active, Now, Actor(), cedantId)
End Function

' --- Countries -----------------------------------------------------------------
Public Function SelectCountries(ByVal cn As Object) As Variant
    SelectCountries = PerfDb.QueryRows(cn, _
        "SELECT country_name, country_code, country_iso3, region, country_id FROM ref_country ORDER BY country_name")
End Function

Public Sub UpsertCountry(ByVal cn As Object, ByVal existingId As Variant, ByVal nm As String, _
                         ByVal iso2 As String, ByVal iso3 As Variant, ByVal region As Variant)
    If IsEmpty(existingId) Or IsNull(existingId) Then
        PerfDb.ExecParams cn, "INSERT INTO ref_country (country_code, country_iso3, country_name, region) VALUES (?, ?, ?, ?)", _
            iso2, iso3, nm, region
    Else
        PerfDb.ExecParams cn, "UPDATE ref_country SET country_code = ?, country_iso3 = ?, country_name = ?, region = ? " & _
            "WHERE country_id = ?", iso2, iso3, nm, region, CLng(existingId)
    End If
End Sub

' --- Currencies and FX ---------------------------------------------------------
Public Function SelectRates(ByVal cn As Object) As Variant
    SelectRates = PerfDb.QueryRows(cn, _
        "SELECT c.currency_code, c.currency_name, f.rate_to_usd, f.effective_date " & _
        "FROM ref_currency AS c LEFT JOIN vw_fx_latest AS f ON c.currency_code = f.currency_code " & _
        "ORDER BY c.currency_code")
End Function

Public Sub UpsertCurrency(ByVal cn As Object, ByVal exists As Boolean, ByVal code As String, ByVal nm As String)
    If exists Then
        PerfDb.ExecParams cn, "UPDATE ref_currency SET currency_name = ? WHERE currency_code = ?", nm, code
    Else
        PerfDb.ExecParams cn, "INSERT INTO ref_currency (currency_code, currency_name) VALUES (?, ?)", code, nm
    End If
End Sub

Public Function UpsertRate(ByVal cn As Object, ByVal code As String, ByVal rateToUsd As Double, _
                           ByVal effective As Date) As Boolean
    Dim cur As Variant
    cur = PerfDb.Scalar(cn, "SELECT rate_to_usd FROM ref_exchange_rate WHERE currency_code = ? AND effective_date = ?", _
        code, effective)
    If IsNull(cur) Then
        PerfDb.ExecParams cn, "INSERT INTO ref_exchange_rate (currency_code, rate_to_usd, effective_date, created_at, created_by) " & _
            "VALUES (?, ?, ?, ?, ?)", code, rateToUsd, effective, Now, Actor()
        UpsertRate = True
    ElseIf Abs(CDbl(cur) - rateToUsd) > 0.0000000001 Then
        PerfDb.ExecParams cn, "UPDATE ref_exchange_rate SET rate_to_usd = ? WHERE currency_code = ? AND effective_date = ?", _
            rateToUsd, code, effective
        UpsertRate = True
    End If
End Function

' --- Treaty types and classes of business ---------------------------------------
Public Function SelectTreatyTypes(ByVal cn As Object) As Variant
    SelectTreatyTypes = PerfDb.QueryRows(cn, _
        "SELECT treaty_type_name, treaty_category FROM ref_treaty_type ORDER BY treaty_category DESC, treaty_type_name")
End Function

Public Function SelectCobs(ByVal cn As Object) As Variant
    SelectCobs = PerfDb.QueryRows(cn, "SELECT cob_name, cob_code FROM ref_class_of_business ORDER BY cob_name")
End Function

Public Sub UpsertTreatyType(ByVal cn As Object, ByVal exists As Boolean, ByVal nm As String, ByVal category As Variant)
    If exists Then
        PerfDb.ExecParams cn, "UPDATE ref_treaty_type SET treaty_category = ? WHERE treaty_type_name = ?", category, nm
    Else
        PerfDb.ExecParams cn, "INSERT INTO ref_treaty_type (treaty_type_name, treaty_category) VALUES (?, ?)", nm, category
    End If
End Sub

Public Sub UpsertCob(ByVal cn As Object, ByVal exists As Boolean, ByVal nm As String, ByVal code As Variant)
    If exists Then
        PerfDb.ExecParams cn, "UPDATE ref_class_of_business SET cob_code = ? WHERE cob_name = ?", code, nm
    Else
        PerfDb.ExecParams cn, "INSERT INTO ref_class_of_business (cob_name, cob_code) VALUES (?, ?)", nm, code
    End If
End Sub
