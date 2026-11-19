Attribute VB_Name = "PerfSeed"
'===============================================================================
' PerfSeed - Universe's predefined reference data (Universe:
' server/src/startup/ensureReferenceData.js + db/seeds/002_reference_data.sql).
'
' Same rule as Universe's boot seed: INSERT IF MISSING ONLY. It never renames,
' deactivates or deletes anything, so edits made by users survive a re-run.
' Called by Setup every time it runs.
'===============================================================================
Option Explicit

' Universe canonical treaty types  (name | category)
Private Function TreatyTypes() As Variant
    TreatyTypes = Array( _
        "Quota Share|PROPORTIONAL", _
        "Quota Share & Surplus|PROPORTIONAL", _
        "First Surplus|PROPORTIONAL", _
        "Second Surplus|PROPORTIONAL", _
        "Third Surplus|PROPORTIONAL", _
        "Fac Oblig|PROPORTIONAL", _
        "Risk XL|NON_PROPORTIONAL", _
        "CAT XL|NON_PROPORTIONAL", _
        "Risk & CAT XL|NON_PROPORTIONAL", _
        "Stop Loss|NON_PROPORTIONAL", _
        "Aggregate XL|NON_PROPORTIONAL")
End Function

' Universe canonical classes of business  (name | code)
Private Function Classes() As Variant
    Classes = Array( _
        "Property|PROP", "Motor|MOT", "Marine|MAR", "Engineering|ENG", _
        "Liability|LIA", "Medical|MED", "Aviation|AVI", "Energy|ENE", _
        "Agriculture|AGR", "Credit & Surety|CS", "Miscellaneous|MISC", _
        "Life|LIFE", "Group Life|GL", "Workers Compensation|WC")
End Function

' Home countries the seeded cedants need  (ISO2 | name | ISO3 | region)
Private Function SeedCountries() As Variant
    SeedCountries = Array( _
        "SA|Saudi Arabia|SAU|Middle East", "AE|United Arab Emirates|ARE|Middle East", _
        "BH|Bahrain|BHR|Middle East", "KW|Kuwait|KWT|Middle East", _
        "OM|Oman|OMN|Middle East", "EG|Egypt|EGY|Africa", _
        "GB|United Kingdom|GBR|Europe")
End Function

' Universe cedant seed  (name | home country ISO2). The two Universe lists are
' merged; ADNIC, Oman Insurance and Kuwait Insurance are left out because the
' other list already carries them as Abu Dhabi National Insurance,
' Oman Insurance Co and Kuwait Insurance Co.
Private Function Cedants() As Collection
    Dim c As New Collection
    c.Add "Tawuniya|SA"
    c.Add "MEDGULF|SA"
    c.Add "Al Rajhi Takaful|SA"
    c.Add "Bupa Arabia|SA"
    c.Add "Walaa Insurance|SA"
    c.Add "SALAMA|SA"
    c.Add "Malath Insurance|SA"
    c.Add "Gulf Union Insurance|SA"
    c.Add "Orient Insurance|AE"
    c.Add "Oman Insurance Co|AE"
    c.Add "Abu Dhabi National Insurance|AE"
    c.Add "Sukoon Insurance|AE"
    c.Add "Dubai Insurance|AE"
    c.Add "Salama Islamic Insurance|AE"
    c.Add "GIG Bahrain|BH"
    c.Add "Solidarity Bahrain|BH"
    c.Add "Kuwait Insurance Co|KW"
    c.Add "Gulf Insurance Group|KW"
    c.Add "Warba Insurance|KW"
    c.Add "Misr Insurance|EG"
    c.Add "GIG Egypt|EG"
    c.Add "Allianz Egypt|EG"
    c.Add "Dhofar Insurance|OM"
    c.Add "National Life & General|OM"
    c.Add "Aviva|GB"
    c.Add "AXA UK|GB"
    c.Add "RSA Insurance|GB"
    c.Add "Zurich UK|GB"
    Set Cedants = c
End Function

' Adds whatever is missing, in one transaction. Returns a one-line summary.
Public Function EnsureReferenceData(ByVal cn As Object) As String
    Dim inTx As Boolean, nT As Long, nB As Long, nK As Long, nC As Long
    On Error GoTo EH
    cn.BeginTrans
    inTx = True
    nT = SeedTreatyTypes(cn)
    nB = SeedClasses(cn)
    nC = SeedHomeCountries(cn)
    nK = SeedCedants(cn)
    cn.CommitTrans
    inTx = False
    If nT + nB + nC + nK = 0 Then
        EnsureReferenceData = "Universe reference data already present."
    Else
        EnsureReferenceData = "Universe reference data added: " & nT & " treaty type(s), " & nB & _
            " class(es) of business, " & nK & " cedant(s)" & IIf(nC > 0, ", " & nC & " country(ies)", "") & "."
    End If
    Exit Function
EH:
    Dim d As String: d = Err.Description
    If inTx Then PerfDb.RollbackQuietly cn
    Err.Raise vbObjectError + 1600, APP_NAME, "Seeding Universe reference data failed: " & d
End Function

Private Function SeedTreatyTypes(ByVal cn As Object) As Long
    Dim have As Object, items As Variant, i As Long, p() As String
    Set have = PerfRepo.TreatyTypeMap(cn)
    items = TreatyTypes()
    For i = LBound(items) To UBound(items)
        p = Split(items(i), "|")
        If Not have.Exists(p(0)) Then
            PerfRepo.UpsertTreatyType cn, False, p(0), p(1)
            SeedTreatyTypes = SeedTreatyTypes + 1
        End If
    Next i
End Function

Private Function SeedClasses(ByVal cn As Object) As Long
    Dim have As Object, items As Variant, i As Long, p() As String
    Set have = PerfRepo.CobMap(cn)
    items = Classes()
    For i = LBound(items) To UBound(items)
        p = Split(items(i), "|")
        If Not have.Exists(p(0)) Then
            PerfRepo.UpsertCob cn, False, p(0), p(1)
            SeedClasses = SeedClasses + 1
        End If
    Next i
End Function

Private Function SeedHomeCountries(ByVal cn As Object) As Long
    Dim byCode As Object, byName As Object, items As Variant, i As Long, p() As String
    Set byCode = PerfRepo.CountryCodeMap(cn)
    Set byName = PerfRepo.CountryMap(cn)
    items = SeedCountries()
    For i = LBound(items) To UBound(items)
        p = Split(items(i), "|")
        If Not byCode.Exists(p(0)) And Not byName.Exists(p(1)) Then
            PerfRepo.UpsertCountry cn, Empty, p(1), p(0), p(2), p(3)
            SeedHomeCountries = SeedHomeCountries + 1
        End If
    Next i
End Function

Private Function SeedCedants(ByVal cn As Object) As Long
    Dim have As Object, byCode As Object, items As Collection, i As Long, p() As String
    Dim hc As Variant, newId As Long
    Set have = PerfRepo.CedantMap(cn)
    Set byCode = PerfRepo.CountryCodeMap(cn)
    Set items = Cedants()
    For i = 1 To items.Count
        p = Split(items(i), "|")
        If Not have.Exists(p(0)) Then
            hc = Null
            If byCode.Exists(p(1)) Then hc = byCode(p(1))
            newId = PerfRepo.InsertCedant(cn, p(0), hc, True)
            PerfRepo.WriteAudit cn, "cedant", newId, "CEDANT_SEEDED", p(0), "Universe reference seed"
            SeedCedants = SeedCedants + 1
        End If
    Next i
End Function
