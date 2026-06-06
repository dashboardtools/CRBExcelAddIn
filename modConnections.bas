'==============================================================================
' Module : modConnections
' Rôle   : Registre des connexions via DSN ODBC.
'          La clé correspond au champ DataSource dans SharePoint.
'          Avec Case Else, le nom DSN peut être mis directement dans SharePoint.
'==============================================================================
Option Explicit

Public Function GetConnectionString(ByVal key As String) As String
    Select Case UCase(Trim(key))
        Case "MYSQL1"
            GetConnectionString = "DSN=ORBIT_MySQL;"
        Case "MYSQL_RH"
            GetConnectionString = "DSN=ORBIT_RH;"
        Case "MYSQL_COMPTA"
            GetConnectionString = "DSN=ORBIT_Compta;"
        Case Else
            If Trim(key) <> "" Then
                GetConnectionString = "DSN=" & Trim(key) & ";"
            Else
                GetConnectionString = ""
            End If
    End Select
End Function

Public Sub TestConnection(Optional ByVal key As String = "MYSQL1")
    Dim connStr As String
    connStr = GetConnectionString(key)
    If connStr = "" Then MsgBox "Clé inconnue : " & key, vbExclamation, "ORBIT": Exit Sub

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    On Error GoTo Echec
    conn.Open connStr
    conn.Close
    MsgBox "✓ Connexion réussie [" & key & "]", vbInformation, "ORBIT"
    Exit Sub
Echec:
    MsgBox "✗ Échec [" & key & "] :" & vbNewLine & Err.Description, vbCritical, "ORBIT"
End Sub
