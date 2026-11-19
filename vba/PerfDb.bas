Attribute VB_Name = "PerfDb"
'===============================================================================
' PerfDb - ADO connection and command helpers (Universe: server/src/db/pool.js,
' withTransaction.js). Late-bound, so no Tools > References are needed.
' Every statement with user data goes through a parameterised ADODB.Command.
'===============================================================================
Option Explicit

' ADO constants (late binding)
Private Const adCmdText As Long = 1
Private Const adExecuteNoRecords As Long = 128
Private Const adParamInput As Long = 1
Private Const adInteger As Long = 3
Private Const adDouble As Long = 5
Private Const adBoolean As Long = 11
Private Const adDBTimeStamp As Long = 135
Private Const adVarWChar As Long = 202
Private Const adLongVarWChar As Long = 203
Private Const adStateOpen As Long = 1

Private mProvider As String

Private Function Providers() As Variant
    ' Office 2016+ / 365 ship ACE 16; Office 2010/2013 ship ACE 12.
    Providers = Array("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0")
End Function

Private Function ConnString(ByVal provider As String, ByVal path As String) As String
    ConnString = "Provider=" & provider & ";Data Source=" & path & ";"
End Function

' Open a connection to the tracker database.
Public Function Connect() As Object
    Dim path As String
    path = DbPath()
    If Len(Dir$(path)) = 0 Then
        Err.Raise vbObjectError + 1101, APP_NAME, _
            "Database not found:" & vbCrLf & path & vbCrLf & vbCrLf & _
            "Run 'Setup / Create database' first, or correct the path in Start!C7."
    End If
    Set Connect = OpenWithProvider(path)
End Function

Private Function OpenWithProvider(ByVal path As String) As Object
    Dim provs As Variant, i As Long, cn As Object, lastErr As String
    provs = Providers()
    If Len(mProvider) > 0 Then provs = Array(mProvider)
    For i = LBound(provs) To UBound(provs)
        Set cn = CreateObject("ADODB.Connection")
        On Error Resume Next
        cn.Open ConnString(CStr(provs(i)), path)
        If Err.Number = 0 Then
            On Error GoTo 0
            mProvider = CStr(provs(i))
            Set OpenWithProvider = cn
            Exit Function
        End If
        lastErr = Err.Description
        Err.Clear
        On Error GoTo 0
    Next i
    Err.Raise vbObjectError + 1102, APP_NAME, _
        "Could not open the database with the Access (ACE) OLE DB provider." & vbCrLf & _
        "Last error: " & lastErr & vbCrLf & vbCrLf & _
        "If Access is not installed, install the free 'Microsoft Access Database Engine' " & _
        "with the same bitness (32/64-bit) as Excel."
End Function

' Create an empty .accdb file. Tries DAO first (the engine Access itself uses),
' then ADOX through each ACE provider. Reports every attempt if all fail.
Public Sub CreateDatabaseFile(ByVal path As String)
    Const dbVersion120 As Long = 128                 ' .accdb format
    Const dbLangGeneral As String = ";LANGID=0x0409;CP=1252;COUNTRY=0"
    Dim provs As Variant, i As Long, eng As Object, db As Object, cat As Object, errs As String

    ' 1) DAO
    On Error Resume Next
    Set eng = CreateObject("DAO.DBEngine.120")
    If Err.Number = 0 Then
        Set db = eng.CreateDatabase(path, dbLangGeneral, dbVersion120)
        If Err.Number = 0 Then
            db.Close
            On Error GoTo 0
            Exit Sub
        End If
    End If
    errs = "  DAO.DBEngine.120: " & Err.Description & vbCrLf
    Err.Clear
    On Error GoTo 0
    DeletePartial path

    ' 2) ADOX, plain connection string (extra properties cause "Multiple-step" errors)
    provs = Providers()
    For i = LBound(provs) To UBound(provs)
        Set cat = CreateObject("ADOX.Catalog")
        On Error Resume Next
        cat.Create "Provider=" & provs(i) & ";Data Source=" & path & ";"
        If Err.Number = 0 Then
            On Error GoTo 0
            mProvider = CStr(provs(i))
            On Error Resume Next
            cat.ActiveConnection.Close
            On Error GoTo 0
            Set cat = Nothing
            Exit Sub
        End If
        errs = errs & "  ADOX " & provs(i) & ": " & Err.Description & vbCrLf
        Err.Clear
        On Error GoTo 0
        Set cat = Nothing
        DeletePartial path
    Next i
    Err.Raise vbObjectError + 1103, APP_NAME, "Could not create " & path & vbCrLf & vbCrLf & errs & vbCrLf & _
        "Workaround: in Access create a Blank database with that name in that folder, close Access, " & _
        "then run Setup again."
End Sub

Private Sub DeletePartial(ByVal path As String)
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    On Error GoTo 0
End Sub

Public Sub CloseQuietly(ByRef cn As Object)
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set cn = Nothing
End Sub

Public Sub RollbackQuietly(ByVal cn As Object)
    On Error Resume Next
    If Not cn Is Nothing Then cn.RollbackTrans
End Sub

' Execute DDL / plain SQL with no parameters.
Public Sub ExecSql(ByVal cn As Object, ByVal sql As String)
    cn.Execute sql, , adCmdText + adExecuteNoRecords
End Sub

' Execute a parameterised statement; returns records affected.
Public Function ExecParams(ByVal cn As Object, ByVal sql As String, ParamArray args() As Variant) As Long
    Dim cmd As Object, n As Variant
    Set cmd = BuildCmd(cn, sql, args)
    cmd.Execute n, , adCmdText + adExecuteNoRecords
    If IsEmpty(n) Or IsNull(n) Then n = 0
    ExecParams = CLng(n)
End Function

' Run a parameterised SELECT; returns a 1-based (row, col) Variant array, or Empty.
' Nulls come back as Empty so they write to cells as blanks.
Public Function QueryRows(ByVal cn As Object, ByVal sql As String, ParamArray args() As Variant) As Variant
    Dim cmd As Object, rs As Object, raw As Variant, out() As Variant
    Dim r As Long, c As Long, nr As Long, nc As Long
    Set cmd = BuildCmd(cn, sql, args)
    Set rs = cmd.Execute
    If rs.EOF Then
        rs.Close
        QueryRows = Empty
        Exit Function
    End If
    raw = rs.GetRows()             ' (col, row), 0-based
    rs.Close
    nc = UBound(raw, 1) + 1
    nr = UBound(raw, 2) + 1
    ReDim out(1 To nr, 1 To nc)
    For r = 1 To nr
        For c = 1 To nc
            If IsNull(raw(c - 1, r - 1)) Then
                out(r, c) = Empty
            Else
                out(r, c) = raw(c - 1, r - 1)
            End If
        Next c
    Next r
    QueryRows = out
End Function

' First column of first row, or Null.
Public Function Scalar(ByVal cn As Object, ByVal sql As String, ParamArray args() As Variant) As Variant
    Dim cmd As Object, rs As Object
    Set cmd = BuildCmd(cn, sql, args)
    Set rs = cmd.Execute
    If rs.EOF Then
        Scalar = Null
    Else
        Scalar = rs.Fields(0).Value
    End If
    rs.Close
End Function

' Autonumber of the last INSERT on this connection.
Public Function LastId(ByVal cn As Object) As Long
    LastId = CLng(Scalar(cn, "SELECT @@IDENTITY"))
End Function

Private Function BuildCmd(ByVal cn As Object, ByVal sql As String, ByVal args As Variant) As Object
    Dim cmd As Object, i As Long
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    If IsArray(args) Then
        For i = LBound(args) To UBound(args)
            cmd.Parameters.Append MakeParam(cmd, args(i))
        Next i
    End If
    Set BuildCmd = cmd
End Function

Private Function MakeParam(ByVal cmd As Object, ByVal v As Variant) As Object
    Dim s As String
    Select Case VarType(v)
        Case vbNull, vbEmpty
            Set MakeParam = cmd.CreateParameter("p", adVarWChar, adParamInput, 255, Null)
        Case vbString
            s = CStr(v)
            If Len(s) > 255 Then
                Set MakeParam = cmd.CreateParameter("p", adLongVarWChar, adParamInput, Len(s), s)
            Else
                Set MakeParam = cmd.CreateParameter("p", adVarWChar, adParamInput, 255, s)
            End If
        Case vbInteger, vbLong, vbByte
            Set MakeParam = cmd.CreateParameter("p", adInteger, adParamInput, 0, CLng(v))
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            Set MakeParam = cmd.CreateParameter("p", adDouble, adParamInput, 0, CDbl(v))
        Case vbDate
            Set MakeParam = cmd.CreateParameter("p", adDBTimeStamp, adParamInput, 0, CDate(v))
        Case vbBoolean
            Set MakeParam = cmd.CreateParameter("p", adBoolean, adParamInput, 0, CBool(v))
        Case Else
            Set MakeParam = cmd.CreateParameter("p", adVarWChar, adParamInput, 255, CStr(v))
    End Select
End Function

' Does a table or saved query exist? (used by the migration runner)
Public Function ObjectExists(ByVal cn As Object, ByVal objName As String) As Boolean
    Dim rs As Object
    On Error Resume Next
    Set rs = cn.Execute("SELECT TOP 1 * FROM [" & objName & "]")
    ObjectExists = (Err.Number = 0)
    If Not rs Is Nothing Then rs.Close
    Err.Clear
End Function
