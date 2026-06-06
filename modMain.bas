'==============================================================================
' Module : modMain
' Rôle   : Point d'entrée ORBIT v2.
'          Callbacks ruban, fonction =ORBIT(), orchestration principale.
'==============================================================================
Option Explicit

'==============================================================================
' CALLBACKS RUBAN
'==============================================================================

Public Sub RibbonRefreshAll(control As IRibbonControl)
    RefreshAllOrbitTables
End Sub

Public Sub RibbonRunQuery(control As IRibbonControl)
    RunSingleQuery
End Sub

Public Sub RibbonSignalerDysqualite(control As IRibbonControl)
    SignalerDysqualite
End Sub

Public Sub RibbonSettings(control As IRibbonControl)
    OpenSettingsDialog
End Sub

Public Sub RibbonDiagnostics(control As IRibbonControl)
    OpenLogFile
End Sub

Public Sub RibbonClearTrace(control As IRibbonControl)
    ClearTraceSheet
End Sub

Public Sub RibbonReloadCache(control As IRibbonControl)
    InvalidateCache
    Dim q As Collection: Set q = GetAllQueries(True)
    MsgBox q.Count & " requêtes rechargées depuis SharePoint.", vbInformation, "ORBIT"
End Sub

Public Sub RibbonAbout(control As IRibbonControl)
    MsgBox "ORBIT Data Access Plugin  v2.0" & vbNewLine & vbNewLine & _
           "Plateforme gouvernée d'accès aux données d'entreprise." & vbNewLine & _
           "SQL · REST API · SharePoint Catalogue · Audit centralisé" & vbNewLine & vbNewLine & _
           "Log : " & GetLogFilePath(), vbInformation, "ORBIT"
End Sub

'==============================================================================
' EXÉCUTION REQUÊTE UNIQUE (bouton "Requête…")
'==============================================================================
Public Sub RunSingleQuery()
    Dim ctx     As Object:     Set ctx = CaptureContext()
    Dim queries As Collection: Set queries = GetAllQueries()

    If queries.Count = 0 Then
        MsgBox "Catalogue vide ou inaccessible.", vbExclamation, "ORBIT": Exit Sub
    End If

    ' Sélectionner la requête
    Dim selected As Object: Set selected = ShowQuerySelector(queries)
    If selected Is Nothing Then Exit Sub

    ' Consolider les paramètres (même si une seule requête — pour réutiliser la validation)
    Dim oneQuery As New Collection: oneQuery.Add selected
    Dim consolidated As Variant: consolidated = ConsolidateParameters(oneQuery)
    Dim paramsJson   As String:  paramsJson   = "{}"

    If Not IsEmpty(consolidated) Then
        LoadLastConsolidatedParams consolidated

        frmConsolidatedParams.InitForm consolidated, selected("Name")
        frmConsolidatedParams.Show vbModal

        If frmConsolidatedParams.Cancelled Then Exit Sub
        consolidated = frmConsolidatedParams.ParamArray_

        Dim errs As Collection: Set errs = ValidateConsolidatedParams(consolidated)
        If errs.Count > 0 Then
            Dim errTxt As String
            Dim e As Variant
            For Each e In errs: errTxt = errTxt & "  • " & e & vbNewLine: Next e
            MsgBox "Erreurs de validation :" & vbNewLine & errTxt, vbExclamation, "ORBIT"
            Exit Sub
        End If

        SaveConsolidatedParams consolidated
        paramsJson = ParametersToJson(consolidated)
    End If

    ' Exécuter
    Dim sw As Single: sw = Timer
    Dim rowCount As Long: rowCount = 0
    Dim status As String: status = "OK"
    Dim errMsg As String: errMsg = ""

    On Error GoTo QueryErr

    Dim data As Variant

    If LCase(selected("SourceType")) = "api" Then
        data = CallApi(selected("ApiEndpoint"), consolidated)
    Else
        Dim finalSql As String
        If Not IsEmpty(consolidated) Then
            finalSql = SubstituteParameters(selected("SQLQuery"), consolidated)
        Else
            finalSql = selected("SQLQuery")
        End If
        Dim rs As Object: Set rs = ExecuteQuery(finalSql, selected("DataSource"))
        data = RecordsetToArray(rs)
        rs.Close: Set rs = Nothing
    End If

    InjectDataTable data, selected("TargetSheet"), selected("TargetTable")
    rowCount = UBound(data, 1)
    Dim durationMs As Long: durationMs = CLng((Timer - sw) * 1000)

    LogAuditEntry ctx, selected("QueryId"), selected("SourceType"), _
                  selected("ApiEndpoint"), paramsJson, rowCount, durationMs, "OK", ""

    UpdateTraceSheet selected("QueryId"), selected("TargetSheet"), selected("TargetTable"), _
                     paramsJson, rowCount, durationMs, "OK"

    MsgBox rowCount & " lignes injectées dans [" & selected("TargetSheet") & "]." & _
           selected("TargetTable") & " (" & durationMs & " ms)", _
           vbInformation, "ORBIT"
    Exit Sub

QueryErr:
    errMsg = Err.Description: Err.Clear
    durationMs = CLng((Timer - sw) * 1000)
    LogAuditEntry ctx, selected("QueryId"), selected("SourceType"), _
                  selected("ApiEndpoint"), paramsJson, 0, durationMs, "ERREUR", errMsg
    UpdateTraceSheet selected("QueryId"), selected("TargetSheet"), selected("TargetTable"), _
                     paramsJson, 0, durationMs, "ERREUR"
    OrbitLog "ERR", selected("QueryId"), errMsg
    MsgBox "Erreur ORBIT :" & vbNewLine & errMsg, vbCritical, "ORBIT"
End Sub

'==============================================================================
' FONCTION EXCEL PERSONNALISÉE =ORBIT()
' =ORBIT("QueryId" [; "Param1=Valeur1" ...])
' Saisir en Ctrl+Maj+Entrée sur Excel < 365, spill automatique sur 365.
'==============================================================================
Public Function ORBIT(ByVal queryId As String, ParamArray rawParams() As Variant) As Variant
    On Error GoTo UdfError

    Dim q As Object: Set q = FindQuery(queryId)
    If q Is Nothing Then ORBIT = "#ORBIT : Requête """ & queryId & """ introuvable": Exit Function

    ' Parser et appliquer les paramètres "Nom=Valeur"
    Dim pa As Variant
    If LCase(q("SourceType")) = "api" Then
        pa = ParseParameters(q("ApiEndpoint"))
    Else
        pa = ParseParameters(q("SQLQuery"))
    End If

    Dim i As Integer
    For i = 0 To UBound(rawParams)
        If IsMissing(rawParams(i)) Then GoTo NextRaw
        Dim raw As String: raw = CStr(rawParams(i))
        Dim eq  As Integer: eq = InStr(raw, "=")
        If eq > 1 Then
            Dim pn As String: pn = Trim(Left(raw, eq - 1))
            Dim pv As String: pv = Trim(Mid(raw, eq + 1))
            If Not IsEmpty(pa) Then
                Dim j As Integer
                For j = 0 To UBound(pa, 1)
                    If LCase(pa(j, 0)) = LCase(pn) Then pa(j, 2) = pv: Exit For
                Next j
            End If
        End If
NextRaw:
    Next i

    Dim data As Variant
    If LCase(q("SourceType")) = "api" Then
        data = CallApi(q("ApiEndpoint"), pa)
    Else
        Dim finalSql As String
        If Not IsEmpty(pa) Then finalSql = SubstituteParameters(q("SQLQuery"), pa) _
        Else finalSql = q("SQLQuery")
        Dim rs As Object: Set rs = ExecuteQuery(finalSql, q("DataSource"))
        data = RecordsetToArray(rs): rs.Close: Set rs = Nothing
    End If

    ORBIT = data
    Exit Function

UdfError:
    ORBIT = "#ORBIT ERR : " & Err.Description
End Function

'==============================================================================
' SÉLECTEUR DE REQUÊTE (simplifié — remplaçable par frmQueryBrowser)
'==============================================================================
Public Function ShowQuerySelector(ByVal queries As Collection) As Object
    Dim list As String
    Dim i    As Integer: i = 1
    Dim q    As Object
    For Each q In queries
        list = list & i & ". [" & q("SourceType") & "] " & q("QueryId") & _
               "  —  " & q("Name") & vbNewLine
        i = i + 1
    Next q

    Dim choice As String
    choice = InputBox("Sélectionner une requête :" & vbNewLine & vbNewLine & list, _
                      "ORBIT — Catalogue", "1")
    If choice = "" Or Not IsNumeric(choice) Then Set ShowQuerySelector = Nothing: Exit Function

    Dim idx As Integer: idx = CInt(choice)
    If idx < 1 Or idx > queries.Count Then
        MsgBox "Numéro invalide.", vbExclamation, "ORBIT"
        Set ShowQuerySelector = Nothing: Exit Function
    End If

    Set ShowQuerySelector = queries(idx)
End Function

'==============================================================================
' SETTINGS — dialog simplifié
'==============================================================================
Public Sub OpenSettingsDialog()
    Dim siteUrl  As String: siteUrl  = GetOrbitSetting("SharePoint", "SiteUrl", "")
    Dim listName As String: listName = GetOrbitSetting("SharePoint", "ListName", "OrbitQueries")

    siteUrl = InputBox("URL du site SharePoint :", "ORBIT — Paramètres", siteUrl)
    If siteUrl = "" Then Exit Sub
    SaveOrbitSetting "SharePoint", "SiteUrl", Trim(siteUrl)

    listName = InputBox("Nom de la liste de requêtes :", "ORBIT — Paramètres", listName)
    If listName = "" Then listName = "OrbitQueries"
    SaveOrbitSetting "SharePoint", "ListName", listName

    InvalidateCache
    MsgBox "Paramètres sauvegardés. Cache réinitialisé.", vbInformation, "ORBIT"
End Sub
