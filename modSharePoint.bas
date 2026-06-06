'==============================================================================
' Module : modSharePoint
' Rôle   : Toutes les interactions avec SharePoint :
'          - Lecture du catalogue de requêtes (OrbitQueries)
'          - Logging d'audit centralisé  (OrbitAuditLog)
'          - Enregistrement des dysqualités (DataQualityIssues)
'          - Lecture des contacts DQ  (DataQualityContacts)
'          - Lecture des listes de valeurs autorisées (paramètres validés)
'
' Authentification : Windows SSO via MSXML2.ServerXMLHTTP (on-premise)
'                   ou Bearer token Azure AD (cloud — voir modAzureAD)
'==============================================================================
Option Explicit

Private gQueryCache()   As Object
Private gCacheLoaded    As Boolean
Private gCacheTime      As Date
Private Const CACHE_MIN As Integer = 5

'==============================================================================
' A. CATALOGUE DE REQUÊTES
'==============================================================================

Public Function GetAllQueries(Optional ByVal forceRefresh As Boolean = False) As Collection
    If Not forceRefresh And gCacheLoaded Then
        If DateDiff("n", gCacheTime, Now) < CACHE_MIN Then
            Dim c As New Collection
            Dim i As Integer
            For i = 0 To UBound(gQueryCache)
                c.Add gQueryCache(i)
            Next i
            Set GetAllQueries = c
            Exit Function
        End If
    End If

    Dim siteUrl  As String: siteUrl  = GetOrbitSetting("SharePoint", "SiteUrl", "")
    Dim listName As String: listName = GetOrbitSetting("SharePoint", "ListName", "OrbitQueries")

    If siteUrl = "" Then
        MsgBox "URL SharePoint non configurée. Ouvrez ORBIT > Paramètres.", vbExclamation, "ORBIT"
        Set GetAllQueries = New Collection
        Exit Function
    End If

    Dim url As String
    url = siteUrl & "/_api/web/lists/getbytitle('" & listName & "')/items" & _
          "?$select=QueryId,Title,SQLQuery,ApiEndpoint,SourceType,TargetSheet,TargetTable," & _
          "DataSource,Description,AllowedValues,IsActive&$filter=IsActive eq 1&$top=500"

    Dim json As String: json = SharePointGet(url)
    If json = "" Then Set GetAllQueries = New Collection: Exit Function

    Dim result As Collection: Set result = ParseItemCollection(json, "QueryId")

    ReDim gQueryCache(result.Count - 1)
    Dim j As Integer: j = 0
    Dim q As Object
    For Each q In result
        Set gQueryCache(j) = q: j = j + 1
    Next q
    gCacheLoaded = True
    gCacheTime   = Now

    Set GetAllQueries = result
End Function

Public Function FindQuery(ByVal queryIdOrName As String) As Object
    Dim all As Collection: Set all = GetAllQueries()
    Dim q As Object
    For Each q In all
        If LCase(q("QueryId")) = LCase(queryIdOrName) Or _
           LCase(q("Name"))    = LCase(queryIdOrName) Then
            Set FindQuery = q: Exit Function
        End If
    Next q
    Set FindQuery = Nothing
End Function

Public Sub InvalidateCache()
    gCacheLoaded = False
End Sub

'==============================================================================
' B. LOGGING D'AUDIT CENTRALISÉ → liste SharePoint "OrbitAuditLog"
'
' Champs de la liste SharePoint à créer :
'   Title (=QueryId), UserLogin, UserFullName, WorkbookName, WorkbookPath,
'   ActiveSheet, ActiveTable, QueryId, ApiEndpoint, SourceType,
'   ParametersJson, RowCount, DurationMs, ExecutionStatus, ErrorMessage,
'   OrbitVersion, MachineName, ExecutedAt
'==============================================================================
Public Sub LogAuditEntry(ByVal ctx As Object, _
                          ByVal queryId As String, _
                          ByVal sourceType As String, _
                          ByVal endpoint As String, _
                          ByVal paramsJson As String, _
                          ByVal rowCount As Long, _
                          ByVal durationMs As Long, _
                          ByVal status As String, _
                          Optional ByVal errorMsg As String = "")
    On Error GoTo LogFailed   ' Le logging ne doit jamais bloquer l'utilisateur

    Dim siteUrl  As String: siteUrl  = GetOrbitSetting("SharePoint", "SiteUrl", "")
    Dim listName As String: listName = GetOrbitSetting("SharePoint", "AuditListName", "OrbitAuditLog")
    If siteUrl = "" Then Exit Sub

    Dim url As String
    url = siteUrl & "/_api/web/lists/getbytitle('" & listName & "')/items"

    Dim body As String
    body = "{""__metadata"":{""type"":""SP.Data." & listName & "ListItem""}," & _
           """Title"":""" & JsonEsc(queryId) & """," & _
           """UserLogin"":""" & JsonEsc(ctx("UserLogin")) & """," & _
           """UserFullName"":""" & JsonEsc(ctx("UserFullName")) & """," & _
           """WorkbookName"":""" & JsonEsc(ctx("WorkbookName")) & """," & _
           """WorkbookPath"":""" & JsonEsc(ctx("WorkbookPath")) & """," & _
           """ActiveSheet"":""" & JsonEsc(ctx("ActiveSheet")) & """," & _
           """ActiveTable"":""" & JsonEsc(ctx("ActiveTable")) & """," & _
           """QueryId"":""" & JsonEsc(queryId) & """," & _
           """SourceType"":""" & JsonEsc(sourceType) & """," & _
           """ApiEndpoint"":""" & JsonEsc(endpoint) & """," & _
           """ParametersJson"":""" & JsonEsc(paramsJson) & """," & _
           """RowCount"":" & rowCount & "," & _
           """DurationMs"":" & durationMs & "," & _
           """ExecutionStatus"":""" & JsonEsc(status) & """," & _
           """ErrorMessage"":""" & JsonEsc(errorMsg) & """," & _
           """OrbitVersion"":""" & JsonEsc(ctx("OrbitVersion")) & """," & _
           """MachineName"":""" & JsonEsc(ctx("MachineName")) & """," & _
           """ExecutedAt"":""" & ctx("TimestampISO") & """}"

    SharePointPost url, body
    OrbitLog "AUDIT", queryId, "Logged to SharePoint — " & status
    Exit Sub

LogFailed:
    ' Dégradation gracieuse : on logue localement si SharePoint est inaccessible
    OrbitLog "AUDIT-LOCAL", queryId, "SharePoint unreachable — " & status & " — " & Err.Description
End Sub

'==============================================================================
' C. INCIDENTS DE DYSQUALITÉ → liste SharePoint "DataQualityIssues"
'
' Champs de la liste :
'   Title (=IncidentID), DeclarationDate, UserLogin, WorkbookName, WorkbookPath,
'   ActiveSheet, ActiveTable, QueryId, ParametersJson, Description,
'   Criticality, Status, ResponsibleTeam
'==============================================================================
Public Function CreateDQIssue(ByVal ctx As Object, _
                               ByVal queryId As String, _
                               ByVal paramsJson As String, _
                               ByVal description As String, _
                               ByVal criticality As String) As String
    On Error GoTo DQFailed

    Dim siteUrl  As String: siteUrl  = GetOrbitSetting("SharePoint", "SiteUrl", "")
    Dim listName As String: listName = GetOrbitSetting("SharePoint", "DQListName", "DataQualityIssues")
    If siteUrl = "" Then CreateDQIssue = "": Exit Function

    ' Générer un ID unique : DQ-YYYYMMDD-HHMMSS-USER
    Dim incidentId As String
    incidentId = "DQ-" & Format(Now, "YYYYMMDD-HHMMSS") & "-" & UCase(Left(ctx("UserLogin"), 4))

    Dim url As String
    url = siteUrl & "/_api/web/lists/getbytitle('" & listName & "')/items"

    Dim body As String
    body = "{""__metadata"":{""type"":""SP.Data." & listName & "ListItem""}," & _
           """Title"":""" & JsonEsc(incidentId) & """," & _
           """DeclarationDate"":""" & ctx("TimestampISO") & """," & _
           """UserLogin"":""" & JsonEsc(ctx("UserLogin")) & """," & _
           """WorkbookName"":""" & JsonEsc(ctx("WorkbookName")) & """," & _
           """WorkbookPath"":""" & JsonEsc(ctx("WorkbookPath")) & """," & _
           """ActiveSheet"":""" & JsonEsc(ctx("ActiveSheet")) & """," & _
           """ActiveTable"":""" & JsonEsc(ctx("ActiveTable")) & """," & _
           """QueryId"":""" & JsonEsc(queryId) & """," & _
           """ParametersJson"":""" & JsonEsc(paramsJson) & """," & _
           """Description"":""" & JsonEsc(description) & """," & _
           """Criticality"":""" & JsonEsc(criticality) & """," & _
           """Status"":""Nouveau""}"

    SharePointPost url, body
    CreateDQIssue = incidentId
    OrbitLog "DQ", incidentId, "Incident enregistré dans SharePoint"
    Exit Function

DQFailed:
    OrbitLog "DQ-ERROR", "", "Échec création incident : " & Err.Description
    CreateDQIssue = ""
End Function

'==============================================================================
' D. CONTACTS DYSQUALITÉ → liste SharePoint "DataQualityContacts"
'
' Champs : Title, Team, DataDomain, Email, Criticality, Scope
'==============================================================================
Public Function GetDQContacts(Optional ByVal domain As String = "") As Collection
    Dim siteUrl  As String: siteUrl  = GetOrbitSetting("SharePoint", "SiteUrl", "")
    Dim listName As String: listName = GetOrbitSetting("SharePoint", "DQContactsListName", "DataQualityContacts")
    If siteUrl = "" Then Set GetDQContacts = New Collection: Exit Function

    Dim url As String
    url = siteUrl & "/_api/web/lists/getbytitle('" & listName & "')/items" & _
          "?$select=Title,Team,DataDomain,Email,Criticality,Scope&$top=100"
    If domain <> "" Then
        url = url & "&$filter=DataDomain eq '" & domain & "'"
    End If

    Set GetDQContacts = ParseItemCollection(SharePointGet(url), "Email")
End Function

'==============================================================================
' E. LISTES DE VALEURS AUTORISÉES (validation des paramètres)
' → Liste SharePoint "OrbitParameterValues"
'
' Champs : Title (=ParameterName), AllowedValues (multiline, une valeur/ligne)
'          QueryId (optionnel — restreindre à une requête)
'==============================================================================
Public Function GetAllowedValues(ByVal paramName As String, _
                                  Optional ByVal queryId As String = "") As Variant
    Dim siteUrl  As String: siteUrl  = GetOrbitSetting("SharePoint", "SiteUrl", "")
    Dim listName As String: listName = GetOrbitSetting("SharePoint", "ParamValuesListName", "OrbitParameterValues")
    If siteUrl = "" Then GetAllowedValues = Empty: Exit Function

    Dim filter As String
    filter = "$filter=Title eq '" & paramName & "'"
    If queryId <> "" Then filter = filter & " and QueryId eq '" & queryId & "'"

    Dim url As String
    url = siteUrl & "/_api/web/lists/getbytitle('" & listName & "')/items?" & filter & _
          "&$select=AllowedValues&$top=1"

    Dim json As String: json = SharePointGet(url)
    Dim raw  As String: raw  = ExtractJsonString(json, "AllowedValues")

    If raw = "" Then GetAllowedValues = Empty: Exit Function

    ' Valeurs séparées par des sauts de ligne ou des virgules
    Dim lines() As String
    If InStr(raw, vbLf) > 0 Then
        lines = Split(raw, vbLf)
    Else
        lines = Split(raw, ",")
    End If

    ' Nettoyer les espaces et retours chariot
    Dim i As Integer
    For i = 0 To UBound(lines)
        lines(i) = Trim(Replace(lines(i), vbCr, ""))
    Next i

    GetAllowedValues = lines
End Function

'==============================================================================
' COUCHE HTTP — fonctions communes
'==============================================================================

'------------------------------------------------------------------------------
' SharePointGet
' Effectue un GET REST SharePoint avec authentification Windows (SSO).
'------------------------------------------------------------------------------
Public Function SharePointGet(ByVal url As String) As String
    On Error GoTo HttpErr

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json;odata=verbose"
    http.setOption 2, 13056   ' SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS (intranet)
    http.send

    If http.Status = 200 Then
        SharePointGet = http.responseText
    Else
        OrbitLog "HTTP-ERR", "GET", http.Status & " " & http.statusText & " — " & url
        SharePointGet = ""
    End If
    Exit Function

HttpErr:
    OrbitLog "HTTP-ERR", "GET", Err.Description & " — " & url
    SharePointGet = ""
End Function

'------------------------------------------------------------------------------
' SharePointPost
' Effectue un POST REST SharePoint (création d'item).
' Récupère automatiquement le FormDigest (token CSRF SharePoint on-premise).
'------------------------------------------------------------------------------
Public Sub SharePointPost(ByVal url As String, ByVal body As String)
    On Error GoTo HttpErr

    Dim digest As String: digest = GetFormDigest()

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Accept",       "application/json;odata=verbose"
    http.setRequestHeader "Content-Type", "application/json;odata=verbose"
    http.setRequestHeader "X-RequestDigest", digest
    http.setOption 2, 13056
    http.send body

    If http.Status < 200 Or http.Status >= 300 Then
        OrbitLog "HTTP-ERR", "POST", http.Status & " " & http.statusText
    End If
    Exit Sub

HttpErr:
    OrbitLog "HTTP-ERR", "POST", Err.Description
End Sub

'------------------------------------------------------------------------------
' GetFormDigest
' Récupère le token CSRF SharePoint requis pour les opérations POST.
' Mis en cache 25 minutes (validité du digest = 30 min).
'------------------------------------------------------------------------------
Private gDigest     As String
Private gDigestTime As Date

Private Function GetFormDigest() As String
    If gDigest <> "" And DateDiff("n", gDigestTime, Now) < 25 Then
        GetFormDigest = gDigest: Exit Function
    End If

    Dim siteUrl As String: siteUrl = GetOrbitSetting("SharePoint", "SiteUrl", "")
    If siteUrl = "" Then GetFormDigest = "": Exit Function

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", siteUrl & "/_api/contextinfo", False
    http.setRequestHeader "Accept", "application/json;odata=verbose"
    http.setOption 2, 13056
    http.send ""

    gDigest     = ExtractJsonString(http.responseText, "FormDigestValue")
    gDigestTime = Now
    GetFormDigest = gDigest
End Function

'==============================================================================
' PARSING JSON SHAREPOINT
'==============================================================================

Private Function ParseItemCollection(ByVal json As String, _
                                      ByVal requiredField As String) As Collection
    Dim result As New Collection
    If json = "" Then Set ParseItemCollection = result: Exit Function

    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Global  = True
    rx.Pattern = "\{[^{}]*""" & requiredField & """[^{}]*\}"

    Dim matches As Object: Set matches = rx.Execute(json)
    Dim m As Object
    For Each m In matches
        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
        d("QueryId")      = ExtractJsonString(m.Value, "QueryId")
        d("Name")         = ExtractJsonString(m.Value, "Title")
        d("SQLQuery")     = ExtractJsonString(m.Value, "SQLQuery")
        d("ApiEndpoint")  = ExtractJsonString(m.Value, "ApiEndpoint")
        d("SourceType")   = ExtractJsonString(m.Value, "SourceType")   ' SQL | API | REST
        d("TargetSheet")  = ExtractJsonString(m.Value, "TargetSheet")
        d("TargetTable")  = ExtractJsonString(m.Value, "TargetTable")
        d("DataSource")   = ExtractJsonString(m.Value, "DataSource")
        d("Description")  = ExtractJsonString(m.Value, "Description")
        d("AllowedValues")= ExtractJsonString(m.Value, "AllowedValues")
        ' Contact DQ
        d("Email")        = ExtractJsonString(m.Value, "Email")
        d("Team")         = ExtractJsonString(m.Value, "Team")
        d("DataDomain")   = ExtractJsonString(m.Value, "DataDomain")
        If d(requiredField) <> "" Then result.Add d
    Next m

    Set ParseItemCollection = result
End Function

Public Function ExtractJsonString(ByVal json As String, ByVal key As String) As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = """" & key & """\s*:\s*(?:null|""((?:[^""\\]|\\.)*)"")"
    Dim m As Object: Set m = rx.Execute(json)
    If m.Count > 0 Then
        Dim v As String: v = m(0).SubMatches(0)
        v = Replace(v, "\""", """"): v = Replace(v, "\\", "\")
        v = Replace(v, "\n", vbNewLine): v = Replace(v, "\r", "")
        v = Replace(v, "\t", vbTab)
        ExtractJsonString = v
    Else
        ExtractJsonString = ""
    End If
End Function

Private Function JsonEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\"): s = Replace(s, """", "\""")
    s = Replace(s, vbCr, ""): s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEsc = s
End Function
