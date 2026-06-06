'==============================================================================
' Module : modSQL
' Rôle   : Parser {{Param:type}}, substitution SQL, exécution ADO (MySQL/PG/MSSQL).
' Référence requise : Microsoft ActiveX Data Objects 6.1 Library
'==============================================================================
Option Explicit

Public Const PARAM_TYPE_STRING  As String = "string"
Public Const PARAM_TYPE_DATE    As String = "date"
Public Const PARAM_TYPE_LIST    As String = "list"
Public Const PARAM_TYPE_INTEGER As String = "int"
Public Const PARAM_TYPE_DECIMAL As String = "decimal"

'------------------------------------------------------------------------------
' ParseParameters
' Retourne tableau 2D : (i,0)=Nom  (i,1)=Type  (i,2)=Valeur
'------------------------------------------------------------------------------
Public Function ParseParameters(ByVal sqlTemplate As String) As Variant
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "\{\{(\w+)(?::(\w+))?\}\}"
    rx.Global = True: rx.IgnoreCase = True

    Dim matches As Object: Set matches = rx.Execute(sqlTemplate)
    Dim seen()  As String: ReDim seen(0)
    Dim count   As Integer: count = 0
    Dim m As Object

    For Each m In matches
        Dim pn As String: pn = m.SubMatches(0)
        Dim dup As Boolean: dup = False
        Dim j As Integer
        For j = 0 To count - 1
            If LCase(seen(j)) = LCase(pn) Then dup = True: Exit For
        Next j
        If Not dup Then
            ReDim Preserve seen(count)
            seen(count) = pn: count = count + 1
        End If
    Next m

    If count = 0 Then ParseParameters = Empty: Exit Function

    Dim res() As String: ReDim res(count - 1, 2)
    Dim idx As Integer: idx = 0
    Dim seenCount As Integer: seenCount = 0
    ReDim seen(0)

    For Each m In matches
        pn = m.SubMatches(0)
        dup = False
        For j = 0 To seenCount - 1
            If LCase(seen(j)) = LCase(pn) Then dup = True: Exit For
        Next j
        If Not dup Then
            res(idx, 0) = pn
            res(idx, 1) = IIf(m.SubMatches(1) <> "", LCase(m.SubMatches(1)), PARAM_TYPE_STRING)
            res(idx, 2) = ""
            ReDim Preserve seen(seenCount)
            seen(seenCount) = pn: seenCount = seenCount + 1: idx = idx + 1
        End If
    Next m

    ParseParameters = res
End Function

'------------------------------------------------------------------------------
' SubstituteParameters — remplace tous les {{Param}} par les valeurs SQL
'------------------------------------------------------------------------------
Public Function SubstituteParameters(ByVal sql As String, ByRef pa As Variant) As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True: rx.IgnoreCase = True
    Dim i As Integer
    For i = 0 To UBound(pa, 1)
        rx.Pattern = "\{\{" & pa(i, 0) & "(?::\w+)?\}\}"
        sql = rx.Replace(sql, FormatSqlValue(CStr(pa(i, 2)), pa(i, 1)))
    Next i
    SubstituteParameters = sql
End Function

Private Function FormatSqlValue(ByVal v As String, ByVal t As String) As String
    If v = "" Then FormatSqlValue = "NULL": Exit Function
    Select Case LCase(t)
        Case PARAM_TYPE_DATE
            If IsDate(v) Then FormatSqlValue = "'" & Format(CDate(v), "YYYY-MM-DD") & "'" _
            Else Err.Raise vbObjectError + 1001, , "Date invalide : " & v
        Case PARAM_TYPE_LIST
            Dim items() As String: items = Split(v, ",")
            Dim sb As String: sb = "("
            Dim k As Integer
            For k = 0 To UBound(items)
                If k > 0 Then sb = sb & ","
                sb = sb & "'" & Replace(Trim(items(k)), "'", "''") & "'"
            Next k
            FormatSqlValue = sb & ")"
        Case PARAM_TYPE_INTEGER
            If IsNumeric(v) Then FormatSqlValue = CStr(CLng(v)) _
            Else Err.Raise vbObjectError + 1002, , "Entier invalide : " & v
        Case PARAM_TYPE_DECIMAL
            If IsNumeric(v) Then FormatSqlValue = Replace(CStr(CDbl(v)), ",", ".") _
            Else Err.Raise vbObjectError + 1003, , "Décimal invalide : " & v
        Case Else
            FormatSqlValue = "'" & Replace(v, "'", "''") & "'"
    End Select
End Function

'------------------------------------------------------------------------------
' ExecuteQuery — retourne un ADODB.Recordset
'------------------------------------------------------------------------------
Public Function ExecuteQuery(ByVal finalSql As String, ByVal dsKey As String) As Object
    Dim connStr As String: connStr = GetConnectionString(dsKey)
    If connStr = "" Then
        Err.Raise vbObjectError + 2001, "ExecuteQuery", "Source inconnue : " & dsKey
    End If
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    Dim rs   As Object: Set rs   = CreateObject("ADODB.Recordset")
    conn.ConnectionTimeout = 15
    conn.CommandTimeout    = 60
    conn.Open connStr
    rs.Open finalSql, conn, 0, 1, 1   ' ForwardOnly, ReadOnly, CmdText
    Set ExecuteQuery = rs
End Function

'------------------------------------------------------------------------------
' RecordsetToArray — tableau 2D Variant, ligne 0 = en-têtes
'------------------------------------------------------------------------------
Public Function RecordsetToArray(ByVal rs As Object) As Variant
    If rs Is Nothing Or rs.EOF Then
        Dim e(0, 0) As Variant: e(0, 0) = "(aucun résultat)": RecordsetToArray = e: Exit Function
    End If
    Dim cols As Integer: cols = rs.Fields.Count
    Dim data As Variant: data = rs.GetRows()
    Dim rows As Long:    rows = UBound(data, 2) + 1
    ReDim res(rows, cols - 1) As Variant
    Dim c As Integer
    For c = 0 To cols - 1: res(0, c) = rs.Fields(c).Name: Next c
    Dim r As Long
    For r = 0 To rows - 1
        For c = 0 To cols - 1
            res(r + 1, c) = IIf(IsNull(data(c, r)), "", data(c, r))
        Next c
    Next r
    RecordsetToArray = res
End Function

'------------------------------------------------------------------------------
' ParametersToJson — sérialise le tableau de paramètres en JSON
'------------------------------------------------------------------------------
Public Function ParametersToJson(ByVal pa As Variant) As String
    If IsEmpty(pa) Then ParametersToJson = "{}": Exit Function
    Dim json As String: json = "{"
    Dim i As Integer
    For i = 0 To UBound(pa, 1)
        If i > 0 Then json = json & ","
        json = json & """" & pa(i, 0) & """:""" & _
               Replace(CStr(pa(i, 2)), """", "\""") & """"
    Next i
    ParametersToJson = json & "}"
End Function
