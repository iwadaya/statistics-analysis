Attribute VB_Name = "PerfMigrate"
'===============================================================================
' PerfMigrate - numbered, forward-only schema migrations, recorded in
' schema_migrations (Universe: db/migrations/NNN_*.sql + runMigrations.js).
' To change the schema later, ADD a migration - never edit an applied one.
'===============================================================================
Option Explicit

Public Const MIG_COUNT As Long = 5

Private Function MigrationName(ByVal v As Long) As String
    Select Case v
        Case 1: MigrationName = "001_foundation"
        Case 2: MigrationName = "002_reference_data"
        Case 3: MigrationName = "003_treaties_and_stats"
        Case 4: MigrationName = "004_reporting_views"
        Case 5: MigrationName = "005_seed_usd"
    End Select
End Function

Private Function MigrationSql(ByVal v As Long) As Collection
    Dim c As New Collection, s As String
    Select Case v

    Case 1 ' --- migration ledger + audit log
        s = "CREATE TABLE schema_migrations ("
        s = s & "mig_version LONG NOT NULL, mig_name VARCHAR(100) NOT NULL, "
        s = s & "applied_at DATETIME NOT NULL, applied_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_schema_migrations PRIMARY KEY (mig_version))"
        c.Add s
        s = "CREATE TABLE audit_log ("
        s = s & "event_id COUNTER NOT NULL, entity_type VARCHAR(50) NOT NULL, entity_id LONG NOT NULL, "
        s = s & "event_type VARCHAR(50) NOT NULL, actor VARCHAR(100) NOT NULL, payload LONGTEXT, "
        s = s & "audit_comment VARCHAR(255), created_at DATETIME NOT NULL, "
        s = s & "CONSTRAINT pk_audit_log PRIMARY KEY (event_id))"
        c.Add s
        c.Add "CREATE INDEX ix_audit_log_entity ON audit_log (entity_type, entity_id, created_at)"

    Case 2 ' --- reference data (same tables as Universe / Aggregate Tracker)
        s = "CREATE TABLE ref_country ("
        s = s & "country_id COUNTER NOT NULL, country_code VARCHAR(2) NOT NULL, country_iso3 VARCHAR(3), "
        s = s & "country_name VARCHAR(100) NOT NULL, region VARCHAR(50), "
        s = s & "CONSTRAINT pk_ref_country PRIMARY KEY (country_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_ref_country_code ON ref_country (country_code)"
        c.Add "CREATE UNIQUE INDEX ux_ref_country_name ON ref_country (country_name)"
        s = "CREATE TABLE ref_currency ("
        s = s & "currency_code VARCHAR(3) NOT NULL, currency_name VARCHAR(60) NOT NULL, "
        s = s & "CONSTRAINT pk_ref_currency PRIMARY KEY (currency_code))"
        c.Add s
        ' rate_to_usd = USD per 1 unit (Universe ref_exchange_rate convention)
        s = "CREATE TABLE ref_exchange_rate ("
        s = s & "rate_id COUNTER NOT NULL, currency_code VARCHAR(3) NOT NULL, rate_to_usd DOUBLE NOT NULL, "
        s = s & "effective_date DATETIME NOT NULL, created_at DATETIME NOT NULL, created_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_ref_exchange_rate PRIMARY KEY (rate_id), "
        s = s & "CONSTRAINT fk_fx_currency FOREIGN KEY (currency_code) REFERENCES ref_currency (currency_code))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_fx_ccy_date ON ref_exchange_rate (currency_code, effective_date)"
        s = "CREATE TABLE ref_treaty_type ("
        s = s & "treaty_type_id COUNTER NOT NULL, treaty_type_name VARCHAR(80) NOT NULL, treaty_category VARCHAR(30), "
        s = s & "CONSTRAINT pk_ref_treaty_type PRIMARY KEY (treaty_type_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_treaty_type_name ON ref_treaty_type (treaty_type_name)"
        s = "CREATE TABLE ref_class_of_business ("
        s = s & "cob_id COUNTER NOT NULL, cob_name VARCHAR(80) NOT NULL, cob_code VARCHAR(10), "
        s = s & "CONSTRAINT pk_ref_cob PRIMARY KEY (cob_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_cob_name ON ref_class_of_business (cob_name)"

    Case 3 ' --- cedants (country required), treaties, yearly statistics
        s = "CREATE TABLE cedant ("
        s = s & "cedant_id COUNTER NOT NULL, cedant_name VARCHAR(150) NOT NULL, home_country_id LONG NOT NULL, "
        s = s & "is_active BIT NOT NULL, created_at DATETIME NOT NULL, created_by VARCHAR(100), "
        s = s & "updated_at DATETIME NOT NULL, updated_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_cedant PRIMARY KEY (cedant_id), "
        s = s & "CONSTRAINT fk_cedant_country FOREIGN KEY (home_country_id) REFERENCES ref_country (country_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_cedant_name ON cedant (cedant_name)"

        ' One treaty = cedant x treaty type x class x currency x reference (e.g. MAIN, LAYER 1).
        ' stats_start / stats_end decide which UW years carry statistics.
        ' Shares are fractions (7.5% = 0.075).
        s = "CREATE TABLE treaty ("
        s = s & "treaty_id COUNTER NOT NULL, cedant_id LONG NOT NULL, treaty_type_id LONG NOT NULL, "
        s = s & "cob_id LONG NOT NULL, currency_code VARCHAR(3) NOT NULL, treaty_ref VARCHAR(50) NOT NULL, "
        s = s & "stats_start DATETIME NOT NULL, stats_end DATETIME NOT NULL, "
        s = s & "our_share DOUBLE, min_share DOUBLE, max_share DOUBLE, treaty_notes VARCHAR(255), "
        s = s & "row_version LONG NOT NULL, created_at DATETIME NOT NULL, created_by VARCHAR(100), "
        s = s & "updated_at DATETIME NOT NULL, updated_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_treaty PRIMARY KEY (treaty_id), "
        s = s & "CONSTRAINT fk_treaty_cedant FOREIGN KEY (cedant_id) REFERENCES cedant (cedant_id), "
        s = s & "CONSTRAINT fk_treaty_type FOREIGN KEY (treaty_type_id) REFERENCES ref_treaty_type (treaty_type_id), "
        s = s & "CONSTRAINT fk_treaty_cob FOREIGN KEY (cob_id) REFERENCES ref_class_of_business (cob_id), "
        s = s & "CONSTRAINT fk_treaty_currency FOREIGN KEY (currency_code) REFERENCES ref_currency (currency_code))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_treaty_key ON treaty (cedant_id, treaty_type_id, cob_id, currency_code, treaty_ref)"

        ' One row per treaty per UW year, at 100%, original currency.
        ' Proportional: commission. Non-proportional: brokerage.
        s = "CREATE TABLE treaty_stat ("
        s = s & "stat_id COUNTER NOT NULL, treaty_id LONG NOT NULL, uw_year LONG NOT NULL, "
        s = s & "premium DOUBLE NOT NULL, incurred_claims DOUBLE NOT NULL, commission DOUBLE, brokerage DOUBLE, "
        s = s & "stat_notes VARCHAR(255), row_version LONG NOT NULL, "
        s = s & "created_at DATETIME NOT NULL, created_by VARCHAR(100), updated_at DATETIME NOT NULL, updated_by VARCHAR(100), "
        s = s & "CONSTRAINT pk_treaty_stat PRIMARY KEY (stat_id), "
        s = s & "CONSTRAINT fk_stat_treaty FOREIGN KEY (treaty_id) REFERENCES treaty (treaty_id))"
        c.Add s
        c.Add "CREATE UNIQUE INDEX ux_treaty_stat_year ON treaty_stat (treaty_id, uw_year)"
        c.Add "CREATE INDEX ix_treaty_stat_year ON treaty_stat (uw_year)"

    Case 4 ' --- reporting views
        c.Add "CREATE VIEW vw_fx_latest_date AS SELECT currency_code, MAX(effective_date) AS max_date " & _
              "FROM ref_exchange_rate GROUP BY currency_code"
        s = "CREATE VIEW vw_fx_latest AS SELECT r.currency_code, r.rate_to_usd, r.effective_date "
        s = s & "FROM ref_exchange_rate AS r INNER JOIN vw_fx_latest_date AS m "
        s = s & "ON (r.currency_code = m.currency_code) AND (r.effective_date = m.max_date)"
        c.Add s

        s = "CREATE VIEW vw_treaty_detail AS SELECT t.treaty_id, c.cedant_name, k.country_name, k.region, "
        s = s & "y.treaty_type_name, y.treaty_category, b.cob_name, t.currency_code, t.treaty_ref, "
        s = s & "t.stats_start, t.stats_end, t.our_share, t.min_share, t.max_share, t.treaty_notes, t.row_version "
        s = s & "FROM (((treaty AS t "
        s = s & "INNER JOIN cedant AS c ON t.cedant_id = c.cedant_id) "
        s = s & "INNER JOIN ref_country AS k ON c.home_country_id = k.country_id) "
        s = s & "INNER JOIN ref_treaty_type AS y ON t.treaty_type_id = y.treaty_type_id) "
        s = s & "INNER JOIN ref_class_of_business AS b ON t.cob_id = b.cob_id"
        c.Add s

        s = "CREATE VIEW vw_stat_detail AS SELECT s.stat_id, d.cedant_name, d.treaty_type_name, d.cob_name, "
        s = s & "d.currency_code, d.treaty_ref, s.uw_year, s.premium, s.incurred_claims, s.commission, s.brokerage, "
        s = s & "s.stat_notes, s.row_version, d.country_name, d.region, d.treaty_category, d.our_share, s.treaty_id "
        s = s & "FROM treaty_stat AS s INNER JOIN vw_treaty_detail AS d ON s.treaty_id = d.treaty_id"
        c.Add s

        s = "CREATE VIEW vw_stat_usd AS SELECT d.stat_id, d.treaty_id, d.cedant_name, d.country_name, d.region, "
        s = s & "d.treaty_type_name, d.treaty_category, d.cob_name, d.currency_code, d.treaty_ref, d.uw_year, "
        s = s & "d.our_share, f.rate_to_usd, "
        s = s & "d.premium * f.rate_to_usd AS premium_usd, "
        s = s & "d.incurred_claims * f.rate_to_usd AS claims_usd, "
        s = s & "IIf(d.commission Is Null, 0, d.commission) * f.rate_to_usd AS commission_usd, "
        s = s & "IIf(d.brokerage Is Null, 0, d.brokerage) * f.rate_to_usd AS brokerage_usd, "
        s = s & "(d.premium - d.incurred_claims - IIf(d.commission Is Null, 0, d.commission) "
        s = s & "- IIf(d.brokerage Is Null, 0, d.brokerage)) * f.rate_to_usd AS result_usd "
        s = s & "FROM vw_stat_detail AS d LEFT JOIN vw_fx_latest AS f ON d.currency_code = f.currency_code"
        c.Add s

        c.Add PerfView("vw_perf_by_year", "")
        c.Add PerfView("vw_perf_by_cedant_year", "cedant_name, country_name")
        c.Add PerfView("vw_perf_by_country_year", "country_name, region")
        c.Add PerfView("vw_perf_by_class_year", "cob_name")
        c.Add PerfView("vw_perf_by_treaty_type_year", "treaty_type_name, treaty_category")
        c.Add PerfView("vw_perf_by_treaty_year", "treaty_id, cedant_name, treaty_type_name, cob_name, currency_code, treaty_ref")

        s = "CREATE VIEW vw_stat_missing_fx AS SELECT stat_id, cedant_name, uw_year, currency_code "
        s = s & "FROM vw_stat_usd WHERE rate_to_usd IS NULL"
        c.Add s

    Case 5
        ' data seed, handled in ApplySeed
    End Select
    Set MigrationSql = c
End Function

' Performance by UW year (100% USD and at our current share), grouped by groupCols.
Private Function PerfView(ByVal viewName As String, ByVal groupCols As String) As String
    Dim s As String, g As String, sh As String
    If Len(groupCols) > 0 Then g = groupCols & ", "
    sh = "IIf(our_share Is Null, 0, our_share)"
    s = "CREATE VIEW " & viewName & " AS SELECT " & g & "uw_year, COUNT(*) AS stat_rows, "
    s = s & "SUM(premium_usd) AS premium_usd, SUM(claims_usd) AS claims_usd, "
    s = s & "SUM(commission_usd) AS commission_usd, SUM(brokerage_usd) AS brokerage_usd, "
    s = s & "SUM(result_usd) AS result_usd, "
    s = s & "IIf(SUM(premium_usd) = 0, Null, SUM(claims_usd) / SUM(premium_usd)) AS loss_ratio, "
    s = s & "IIf(SUM(premium_usd) = 0, Null, (SUM(claims_usd) + SUM(commission_usd) + SUM(brokerage_usd)) / SUM(premium_usd)) AS combined_ratio, "
    s = s & "SUM(premium_usd * " & sh & ") AS premium_share_usd, "
    s = s & "SUM(result_usd * " & sh & ") AS result_share_usd "
    s = s & "FROM vw_stat_usd GROUP BY " & g & "uw_year"
    PerfView = s
End Function

Private Sub ApplySeed(ByVal cn As Object)
    If IsNull(PerfDb.Scalar(cn, "SELECT currency_code FROM ref_currency WHERE currency_code = ?", "USD")) Then
        PerfDb.ExecParams cn, "INSERT INTO ref_currency (currency_code, currency_name) VALUES (?, ?)", "USD", "US Dollar"
    End If
End Sub

Private Function CreatedObjectName(ByVal sql As String) As String
    Dim u As String, rest As String, p As Long
    u = UCase$(Trim$(sql))
    If Left$(u, 13) = "CREATE TABLE " Then
        rest = Trim$(Mid$(Trim$(sql), 14))
    ElseIf Left$(u, 12) = "CREATE VIEW " Then
        rest = Trim$(Mid$(Trim$(sql), 13))
    Else
        Exit Function
    End If
    p = InStr(rest, " ")
    If p = 0 Then p = InStr(rest, "(")
    If p > 0 Then rest = Left$(rest, p - 1)
    CreatedObjectName = Replace(rest, "(", "")
End Function

Private Function AppliedVersions(ByVal cn As Object) As Object
    Dim d As Object, rows As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    If PerfDb.ObjectExists(cn, "schema_migrations") Then
        rows = PerfDb.QueryRows(cn, "SELECT mig_version FROM schema_migrations")
        If IsArray(rows) Then
            For i = 1 To UBound(rows, 1)
                d(CLng(rows(i, 1))) = True
            Next i
        End If
    End If
    Set AppliedVersions = d
End Function

' Apply pending migrations in order. Re-runnable: existing objects are skipped.
Public Function RunMigrations(ByVal cn As Object) As String
    Dim applied As Object, v As Long, stmts As Collection, i As Long
    Dim sql As String, objName As String, ran As Long
    Set applied = AppliedVersions(cn)
    For v = 1 To MIG_COUNT
        If Not applied.Exists(v) Then
            Set stmts = MigrationSql(v)
            For i = 1 To stmts.Count
                sql = stmts(i)
                objName = CreatedObjectName(sql)
                If Len(objName) > 0 Then
                    If Not PerfDb.ObjectExists(cn, objName) Then ExecDdl cn, sql, v
                Else
                    ExecOtherDdl cn, sql, v
                End If
            Next i
            If v = 5 Then ApplySeed cn
            PerfDb.ExecParams cn, "INSERT INTO schema_migrations (mig_version, mig_name, applied_at, applied_by) VALUES (?, ?, ?, ?)", _
                v, MigrationName(v), Now, Actor()
            ran = ran + 1
        End If
    Next v
    If ran = 0 Then
        RunMigrations = "Schema up to date (" & MIG_COUNT & " migrations)."
    Else
        RunMigrations = "Applied " & ran & " migration(s); schema at " & MigrationName(MIG_COUNT) & "."
    End If
End Function

Private Sub ExecDdl(ByVal cn As Object, ByVal sql As String, ByVal v As Long)
    On Error GoTo EH
    PerfDb.ExecSql cn, sql
    Exit Sub
EH:
    Err.Raise Err.Number, APP_NAME, "Migration " & MigrationName(v) & " failed:" & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & sql
End Sub

' CREATE INDEX / ALTER: ignore "already exists" so re-runs are safe.
Private Sub ExecOtherDdl(ByVal cn As Object, ByVal sql As String, ByVal v As Long)
    Dim d As String, n As Long
    On Error Resume Next
    PerfDb.ExecSql cn, sql
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    If n <> 0 Then
        If InStr(1, d, "already", vbTextCompare) = 0 Then
            Err.Raise n, APP_NAME, "Migration " & MigrationName(v) & " failed:" & vbCrLf & d & vbCrLf & vbCrLf & sql
        End If
    End If
End Sub
