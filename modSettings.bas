'==============================================================================
' Module : modSettings
' Rôle   : Persistance des paramètres ORBIT (registre Windows),
'          gestion des mots de passe, bindings Custom XML, logging local.
'==============================================================================
Option Explicit

Private gLogFile As String

' ── Settings registre ─────────────────────────────────────────────────────────
Public Sub SaveOrbitSetting(ByVal section As String, ByVal key As String, ByVal value As String)
    SaveSetting "ORBIT", section, key, value
End Sub

Public Function GetOrbitSetting(ByVal section As String, ByVal key As String, _
                                 Optional ByVal defaultValue As String = "") As String
    GetOrbitSetting = GetSetting("ORBIT", section, key, defaultValue)
End Function

' ── Derniers paramètres utilisés ─────────────────────────────────────────────
Public Sub SaveLastParameters(ByVal sheetName As String, ByVal tableName As String, _
                               ByRef paramArray As Variant)
    If IsEmpty(paramArray) Then Exit Sub
    Dim i As Integer
    For i = 0 To UBound(paramArray, 1)
        SaveSetting "ORBIT", "LastParams", sheetName & "|" & tableName & "|" & paramArray(i, 0), _
                    CStr(paramArray(i, 2))
    Next i
End Sub

Public Function LoadLastParameters(ByVal sheetName As String, ByVal tableName As String, _
                                    ByRef paramArray As Variant) As Boolean
    If IsEmpty(paramArray) Then LoadLastParameters = True: Exit Function
    Dim i As Integer
    Dim allFilled As Boolean: allFilled = True
    For i = 0 To UBound(paramArray, 1)
        Dim v As String
        v = GetSetting("ORBIT", "LastParams", sheetName & "|" & tableName & "|" & paramArray(i, 0), "")
        paramArray(i, 2) = v
        If v = "" Then allFilled = False
    Next i
    LoadLastParameters = allFilled
End Function

' ── Logging fichier local ─────────────────────────────────────────────────────
Public Sub OrbitLog(ByVal level As String, ByVal queryId As String, ByVal message As String)
    On Error Resume Next
    If gLogFile = "" Then
        gLogFile = Environ("TEMP") & "\ORBIT_" & Format(Now, "YYYYMMDD") & ".log"
    End If
    Dim f As Integer: f = FreeFile
    Open gLogFile For Append As #f
    Print #f, Format(Now, "YYYY-MM-DD HH:MM:SS") & vbTab & level & vbTab & queryId & vbTab & message
    Close #f
    On Error GoTo 0
End Sub

Public Function GetLogFilePath() As String
    If gLogFile = "" Then
        gLogFile = Environ("TEMP") & "\ORBIT_" & Format(Now, "YYYYMMDD") & ".log"
    End If
    GetLogFilePath = gLogFile
End Function

Public Sub OpenLogFile()
    Dim p As String: p = GetLogFilePath()
    If Dir(p) <> "" Then Shell "notepad.exe " & p, vbNormalFocus _
    Else MsgBox "Aucun log aujourd'hui.", vbInformation, "ORBIT"
End Sub

' ── Custom XML bindings dans le classeur ──────────────────────────────────────
Private Const ORBIT_NS As String = "http://orbit.internal/bindings/v1"

Public Sub SaveBinding(ByVal sheetName As String, ByVal tableName As String, _
                        ByVal queryId As String, ByVal dataSource As String)
    Dim all As String: all = LoadAllBindingsXml()
    Dim newEl As String
    newEl = "<Binding queryId=""" & XmlEsc(queryId) & """ sheet=""" & XmlEsc(sheetName) & _
            """ table=""" & XmlEsc(tableName) & """ dataSource=""" & XmlEsc(dataSource) & """ lastRun=""/>"
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "<Binding[^/]* sheet=""" & XmlEsc(sheetName) & """[^/]* table=""" & XmlEsc(tableName) & """[^/]*/>"
    rx.Global = True
    If rx.Test(all) Then all = rx.Replace(all, newEl) Else all = all & newEl & vbNewLine
    PersistBindingsXml all
End Sub

Public Function FindBinding(ByVal sheetName As String, ByVal tableName As String) As Object
    Dim all As String: all = LoadAllBindingsXml()
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "<Binding[^/]* sheet=""" & XmlEsc(sheetName) & """[^/]* table=""" & XmlEsc(tableName) & """[^>]*/>"
    Dim m As Object: Set m = rx.Execute(all)
    If m.Count = 0 Then Set FindBinding = Nothing: Exit Function
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("queryId")    = XmlAttr(m(0).Value, "queryId")
    d("sheet")      = XmlAttr(m(0).Value, "sheet")
    d("table")      = XmlAttr(m(0).Value, "table")
    d("dataSource") = XmlAttr(m(0).Value, "dataSource")
    Set FindBinding = d
End Function

Private Function LoadAllBindingsXml() As String
    Dim p As Object
    For Each p In ThisWorkbook.CustomXMLParts
        If InStr(p.NamespaceURI, "orbit.internal") > 0 Then
            Dim x As String: x = p.XML
            LoadAllBindingsXml = Mid(x, InStr(x, ">") + 1, InStrRev(x, "<") - InStr(x, ">") - 1)
            Exit Function
        End If
    Next p
    LoadAllBindingsXml = ""
End Function

Private Sub PersistBindingsXml(ByVal content As String)
    Dim p As Object
    For Each p In ThisWorkbook.CustomXMLParts
        If InStr(p.NamespaceURI, "orbit.internal") > 0 Then p.Delete: Exit For
    Next p
    ThisWorkbook.CustomXMLParts.Add "<OrbitBindings xmlns=""" & ORBIT_NS & """>" & vbNewLine & content & vbNewLine & "</OrbitBindings>"
End Sub

Private Function XmlAttr(ByVal xml As String, ByVal attr As String) As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = " " & attr & "=""([^""]*)"""
    Dim m As Object: Set m = rx.Execute(xml)
    If m.Count > 0 Then XmlAttr = m(0).SubMatches(0) Else XmlAttr = ""
End Function

Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;"): s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;"): s = Replace(s, """", "&quot;")
    XmlEsc = s
End Function
