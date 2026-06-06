'==============================================================================
' Module : modContext
' Rôle   : Capture automatique du contexte d'exécution.
'          Centralise toutes les informations environnementales nécessaires
'          au logging, audit, et pré-remplissage des formulaires.
'
' Utilisé par : modAudit, modDysqualite, modMain
'==============================================================================
Option Explicit

' Structure de contexte d'exécution
' Simulée en VBA via un Scripting.Dictionary public
Private mContext As Object   ' Scripting.Dictionary

'------------------------------------------------------------------------------
' CaptureContext
' Prend un snapshot complet du contexte au moment de l'appel.
' À appeler au début de chaque opération ORBIT.
'------------------------------------------------------------------------------
Public Function CaptureContext() As Object
    Dim ctx As Object
    Set ctx = CreateObject("Scripting.Dictionary")

    ' ── Utilisateur ──────────────────────────────────────────────────────────
    ctx("UserLogin")    = Environ("USERNAME")
    ctx("UserDomain")   = Environ("USERDOMAIN")
    ctx("UserFullName") = GetUserFullName()
    ctx("UserEmail")    = GetUserEmail()

    ' ── Machine ───────────────────────────────────────────────────────────────
    ctx("MachineName")  = Environ("COMPUTERNAME")

    ' ── Fichier Excel ─────────────────────────────────────────────────────────
    Dim wb As Workbook
    Set wb = GetActiveOrbitWorkbook()

    If Not wb Is Nothing Then
        ctx("WorkbookName")   = wb.Name
        ctx("WorkbookPath")   = wb.FullName
        ctx("WorkbookDir")    = wb.Path
        ctx("WorkbookSaved")  = Not wb.Saved
    Else
        ctx("WorkbookName")   = ""
        ctx("WorkbookPath")   = ""
        ctx("WorkbookDir")    = ""
        ctx("WorkbookSaved")  = False
    End If

    ' ── Feuille et tableau actifs ─────────────────────────────────────────────
    On Error Resume Next
    ctx("ActiveSheet")  = Application.ActiveSheet.Name
    ctx("ActiveTable")  = GetActiveTableName()
    On Error GoTo 0

    ' ── Horodatage ────────────────────────────────────────────────────────────
    ctx("Timestamp")    = Now
    ctx("TimestampISO") = Format(Now, "YYYY-MM-DDTHH:MM:SS")
    ctx("DateOnly")     = Format(Date, "YYYY-MM-DD")

    ' ── Version ORBIT ─────────────────────────────────────────────────────────
    ctx("OrbitVersion") = "2.0"
    ctx("ExcelVersion") = Application.Version

    Set mContext = ctx
    Set CaptureContext = ctx
End Function

'------------------------------------------------------------------------------
' GetCurrentContext
' Retourne le contexte capturé le plus récent (ou en capture un nouveau).
'------------------------------------------------------------------------------
Public Function GetCurrentContext() As Object
    If mContext Is Nothing Then
        Set GetCurrentContext = CaptureContext()
    Else
        Set GetCurrentContext = mContext
    End If
End Function

'------------------------------------------------------------------------------
' ContextToJson
' Sérialise le contexte en JSON pour le logging SharePoint.
'------------------------------------------------------------------------------
Public Function ContextToJson(ByVal ctx As Object) As String
    Dim json As String
    Dim key As Variant

    json = "{"
    Dim first As Boolean: first = True

    For Each key In ctx.Keys
        If Not first Then json = json & ","
        json = json & """" & key & """:""" & _
               Replace(CStr(ctx(key)), """", "\""") & """"
        first = False
    Next key

    ContextToJson = json & "}"
End Function

'------------------------------------------------------------------------------
' GetActiveTableName
' Retourne le nom du ListObject sous la cellule active (si applicable).
'------------------------------------------------------------------------------
Public Function GetActiveTableName() As String
    Dim cell As Range
    Dim tbl  As ListObject

    On Error Resume Next
    Set cell = Application.ActiveCell
    If cell Is Nothing Then GetActiveTableName = "": Exit Function

    For Each tbl In Application.ActiveSheet.ListObjects
        If Not Intersect(cell, tbl.Range) Is Nothing Then
            GetActiveTableName = tbl.Name
            Exit Function
        End If
    Next tbl
    On Error GoTo 0

    GetActiveTableName = ""
End Function

'------------------------------------------------------------------------------
' GetActiveOrbitWorkbook
' Retourne le classeur actif (peut être étendu pour cibler spécifiquement
' le classeur qui a déclenché le refresh, pas le .xlam lui-même).
'------------------------------------------------------------------------------
Public Function GetActiveOrbitWorkbook() As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        ' Ignorer le complément ORBIT lui-même
        If LCase(Right(wb.Name, 5)) <> ".xlam" And _
           LCase(Right(wb.Name, 4)) <> ".xla" Then
            Set GetActiveOrbitWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetActiveOrbitWorkbook = Nothing
End Function

'------------------------------------------------------------------------------
' GetUserFullName
' Tente de récupérer le nom complet Windows de l'utilisateur.
'------------------------------------------------------------------------------
Private Function GetUserFullName() As String
    On Error Resume Next
    Dim wmi    As Object
    Dim users  As Object
    Dim user   As Object
    Dim login  As String

    login = Environ("USERNAME")
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set users = wmi.ExecQuery( _
        "SELECT FullName FROM Win32_UserAccount WHERE Name='" & login & "'")

    For Each user In users
        If user.FullName <> "" Then
            GetUserFullName = user.FullName
            Exit Function
        End If
    Next user
    On Error GoTo 0

    GetUserFullName = login   ' Fallback sur le login
End Function

'------------------------------------------------------------------------------
' GetUserEmail
' Récupère l'email depuis Outlook si disponible, sinon construit une valeur
' par convention domain\user → user@domain.com (adaptable).
'------------------------------------------------------------------------------
Public Function GetUserEmail() As String
    On Error Resume Next
    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    If Not olApp Is Nothing Then
        GetUserEmail = olApp.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        If GetUserEmail <> "" Then Exit Function
    End If
    On Error GoTo 0

    ' Fallback : convention domain
    GetUserEmail = Environ("USERNAME") & "@" & Environ("USERDNSDOMAIN")
End Function
