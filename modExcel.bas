'==============================================================================
' Module : modExcel
' Rôle   : Injection données dans les ListObject Excel,
'          Refresh All multi-sources (SQL + API),
'          Onglet de traçabilité _ORBIT_TRACE (EDB §8).
'==============================================================================
Option Explicit

Private Const TRACE_SHEET As String = "_ORBIT_TRACE"

'==============================================================================
' INJECTION DONNÉES
'==============================================================================

'------------------------------------------------------------------------------
' InjectDataTable
' Injecte un tableau 2D Variant dans un ListObject Excel.
' Preserve formules, formats, colonnes calculées.
' Compatible avec les résultats SQL (RecordsetToArray) et API (ParseJsonResponse).
'------------------------------------------------------------------------------
Public Sub InjectDataTable(ByVal data As Variant, _
                            ByVal sheetName As String, _
                            ByVal tableName As String)
    Dim ws  As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = GetActiveOrbitWorkbook().Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 4001, "InjectDataTable", "Feuille introuvable : " & sheetName
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    If tbl Is Nothing Then
        Err.Raise vbObjectError + 4002, "InjectDataTable", "Tableau introuvable : " & tableName
    End If

    If IsEmpty(data) Or Not IsArray(data) Then
        If tbl.ListRows.Count > 0 Then tbl.DataBodyRange.Delete xlShiftUp
        Exit Sub
    End If

    Dim rowCount As Long: rowCount = UBound(data, 1)   ' Ligne 0 = en-têtes → données = 1..rowCount
    Dim colCount As Integer: colCount = UBound(data, 2) + 1

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual

    On Error GoTo Cleanup

    ' Snapshot des colonnes calculées
    Dim isCalc()    As Boolean:  ReDim isCalc(tbl.ListColumns.Count - 1)
    Dim calcFmla()  As String:   ReDim calcFmla(tbl.ListColumns.Count - 1)
    Dim c As Integer
    If tbl.ListRows.Count > 0 Then
        For c = 1 To tbl.ListColumns.Count
            Dim fc As Range: Set fc = tbl.ListColumns(c).DataBodyRange.Cells(1, 1)
            If fc.HasFormula Then
                isCalc(c - 1)  = True
                calcFmla(c - 1) = fc.FormulaR1C1
            End If
        Next c
    End If

    ' Construire la map colonnes data → colonnes tableau (matching par nom)
    Dim colMap() As Integer: ReDim colMap(tbl.ListColumns.Count - 1)
    Dim d As Integer
    For c = 0 To tbl.ListColumns.Count - 1
        colMap(c) = -1
        Dim hdr As String: hdr = LCase(tbl.ListColumns(c + 1).Name)
        For d = 0 To colCount - 1
            If LCase(CStr(data(0, d))) = hdr Then colMap(c) = d: Exit For
        Next d
    Next c

    ' Effacer les données existantes
    If tbl.ListRows.Count > 0 Then tbl.DataBodyRange.Delete xlShiftUp

    If rowCount = 0 Then GoTo Cleanup   ' Requête sans résultat

    ' Ajouter les lignes
    Dim r As Long
    For r = 1 To rowCount: tbl.ListRows.Add: Next r

    ' Écrire les valeurs en batch
    ReDim values(1 To rowCount, 1 To tbl.ListColumns.Count) As Variant
    For r = 1 To rowCount
        For c = 0 To tbl.ListColumns.Count - 1
            If Not isCalc(c) And colMap(c) >= 0 Then
                values(r, c + 1) = data(r, colMap(c))
            End If
        Next c
    Next r
    tbl.DataBodyRange.Value = values

    ' Restaurer les formules calculées
    For c = 0 To tbl.ListColumns.Count - 1
        If isCalc(c) And calcFmla(c) <> "" Then
            tbl.ListColumns(c + 1).DataBodyRange.FormulaR1C1 = calcFmla(c)
        End If
    Next c

Cleanup:
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Dim em As String: em = Err.Description: Err.Clear
        Err.Raise vbObjectError + 4003, "InjectDataTable", em
    End If
End Sub

'==============================================================================
' REFRESH ALL — multi-sources SQL + API
'==============================================================================

Public Sub RefreshAllOrbitTables()
    Dim ctx     As Object:     Set ctx = CaptureContext()
    Dim queries As Collection: Set queries = GetAllQueries()
    If queries.Count = 0 Then
        MsgBox "Catalogue vide ou inaccessible.", vbExclamation, "ORBIT": Exit Sub
    End If

    ' Récupérer tous les tableaux du classeur
    Dim allTables As Variant: allTables = GetAllTablesInWorkbook()
    If IsEmpty(allTables) Then
        MsgBox "Aucun tableau Excel dans ce classeur.", vbInformation, "ORBIT": Exit Sub
    End If

    ' Identifier les requêtes concernées (matching TargetSheet + TargetTable)
    Dim matchedQueries As New Collection
    Dim i As Integer
    For i = 0 To UBound(allTables, 1)
        Dim q As Object
        For Each q In queries
            If LCase(q("TargetSheet")) = LCase(allTables(i, 0)) And _
               LCase(q("TargetTable")) = LCase(allTables(i, 1)) Then
                matchedQueries.Add q: Exit For
            End If
        Next q
    Next i

    If matchedQueries.Count = 0 Then
        MsgBox "Aucun tableau ORBIT dans ce classeur.", vbInformation, "ORBIT": Exit Sub
    End If

    ' Consolider les paramètres de toutes les requêtes en un seul dialog
    Dim consolidated As Variant: consolidated = ConsolidateParameters(matchedQueries)
    Dim paramsJson   As String:  paramsJson   = "{}"

    If Not IsEmpty(consolidated) Then
        ' Charger les dernières valeurs
        LoadLastConsolidatedParams consolidated

        ' Afficher le formulaire unique
        frmConsolidatedParams.InitForm consolidated, "Refresh All — " & matchedQueries.Count & " tableau(x)"
        frmConsolidatedParams.Show vbModal

        If frmConsolidatedParams.Cancelled Then Exit Sub
        consolidated = frmConsolidatedParams.ParamArray_

        ' Valider
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

    ' Injecter les paramètres consolidés dans chaque requête
    Set matchedQueries = InjectConsolidatedParams(matchedQueries, consolidated)

    ' Exécuter chaque requête
    Dim success As Integer: success = 0
    Dim failed  As Integer: failed  = 0
    Dim skipped As Integer: skipped = 0

    Dim mq As Object
    For Each mq In matchedQueries
        Dim sw As Single: sw = Timer
        Dim rowCount As Long: rowCount = 0
        Dim status As String: status = "OK"
        Dim errMsg As String: errMsg = ""

        On Error Resume Next
        Err.Clear

        Dim data As Variant

        If LCase(mq("SourceType")) = "api" Then
            ' Appel API REST
            Dim pa As Variant: pa = mq("ResolvedParams")
            If IsEmpty(pa) Then pa = consolidated
            data = CallApi(mq("ApiEndpoint"), pa)
        Else
            ' Requête SQL
            Dim finalSql As String
            pa = mq("ResolvedParams")
            If IsEmpty(pa) Then
                finalSql = mq("SQLQuery")
            Else
                finalSql = SubstituteParameters(mq("SQLQuery"), pa)
            End If
            Dim rs As Object: Set rs = ExecuteQuery(finalSql, mq("DataSource"))
            data = RecordsetToArray(rs)
            If Not rs Is Nothing Then rs.Close: Set rs = Nothing
        End If

        If Err.Number <> 0 Then
            status = "ERREUR": errMsg = Err.Description: failed = failed + 1
            OrbitLog "ERR", mq("QueryId"), errMsg
        Else
            InjectDataTable data, mq("TargetSheet"), mq("TargetTable")
            If Err.Number <> 0 Then
                status = "ERREUR": errMsg = Err.Description: failed = failed + 1
            Else
                rowCount = UBound(data, 1)
                success = success + 1
            End If
        End If

        Err.Clear
        On Error GoTo 0

        Dim durationMs As Long: durationMs = CLng((Timer - sw) * 1000)

        ' Audit SharePoint
        LogAuditEntry ctx, mq("QueryId"), mq("SourceType"), _
                      mq("ApiEndpoint"), paramsJson, rowCount, durationMs, status, errMsg

        ' Mettre à jour l'onglet de traçabilité
        UpdateTraceSheet mq("QueryId"), mq("TargetSheet"), mq("TargetTable"), _
                         paramsJson, rowCount, durationMs, status

        OrbitLog IIf(status = "OK", "OK", "ERR"), mq("QueryId"), _
                 rowCount & " lignes — " & durationMs & "ms"
    Next mq

    MsgBox "Refresh terminé." & vbNewLine & vbNewLine & _
           "✓ Réussis  : " & success & vbNewLine & _
           "✗ Échoués  : " & failed, _
           IIf(failed > 0, vbExclamation, vbInformation), "ORBIT — Refresh All"
End Sub

'==============================================================================
' ONGLET DE TRAÇABILITÉ _ORBIT_TRACE  (EDB §8)
' Génère / met à jour un onglet dédié dans le classeur actif.
'==============================================================================

'------------------------------------------------------------------------------
' UpdateTraceSheet
' Ajoute une ligne dans l'onglet _ORBIT_TRACE pour chaque exécution.
' Crée l'onglet s'il n'existe pas.
'------------------------------------------------------------------------------
Public Sub UpdateTraceSheet(ByVal queryId As String, _
                             ByVal sheetName As String, _
                             ByVal tableName As String, _
                             ByVal paramsJson As String, _
                             ByVal rowCount As Long, _
                             ByVal durationMs As Long, _
                             ByVal status As String)
    On Error GoTo TraceFailed   ' Ne jamais bloquer le refresh pour la trace

    Dim wb As Workbook: Set wb = GetActiveOrbitWorkbook()
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet: Set ws = GetOrCreateTraceSheet(wb)

    ' Trouver la prochaine ligne vide (après l'en-tête)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2 Else lastRow = lastRow + 1

    ' Écrire la ligne de traçabilité
    ws.Cells(lastRow, 1).Value  = Format(Now, "YYYY-MM-DD HH:MM:SS")   ' Horodatage
    ws.Cells(lastRow, 2).Value  = queryId                               ' Requête
    ws.Cells(lastRow, 3).Value  = sheetName & "." & tableName          ' Cible
    ws.Cells(lastRow, 4).Value  = paramsJson                           ' Paramètres JSON
    ws.Cells(lastRow, 5).Value  = rowCount                             ' Lignes
    ws.Cells(lastRow, 6).Value  = durationMs & " ms"                   ' Durée
    ws.Cells(lastRow, 7).Value  = status                               ' Statut
    ws.Cells(lastRow, 8).Value  = Environ("USERNAME")                  ' Utilisateur

    ' Colorer la ligne selon le statut
    Dim rowRange As Range: Set rowRange = ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 8))
    If status = "OK" Then
        rowRange.Interior.Color = RGB(235, 255, 235)   ' Vert clair
    Else
        rowRange.Interior.Color = RGB(255, 235, 235)   ' Rouge clair
    End If

    Exit Sub
TraceFailed:
    OrbitLog "TRACE-ERR", queryId, Err.Description
End Sub

'------------------------------------------------------------------------------
' GetOrCreateTraceSheet
' Retourne l'onglet _ORBIT_TRACE, le crée et le formate si absent.
'------------------------------------------------------------------------------
Private Function GetOrCreateTraceSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(TRACE_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Créer l'onglet en dernière position
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = TRACE_SHEET

        ' Formater l'en-tête
        Dim headers As Variant
        headers = Array("Horodatage", "Requête / API", "Tableau cible", _
                        "Paramètres (JSON)", "Lignes", "Durée", "Statut", "Utilisateur")
        Dim c As Integer
        For c = 0 To UBound(headers)
            ws.Cells(1, c + 1).Value = headers(c)
        Next c

        ' Style en-tête
        With ws.Range("A1:H1")
            .Font.Bold          = True
            .Font.Color         = RGB(255, 255, 255)
            .Interior.Color     = RGB(31, 73, 125)    ' Bleu foncé ORBIT
            .HorizontalAlignment = xlCenter
        End With

        ' Largeurs de colonnes
        ws.Columns("A").ColumnWidth = 20
        ws.Columns("B").ColumnWidth = 25
        ws.Columns("C").ColumnWidth = 25
        ws.Columns("D").ColumnWidth = 40
        ws.Columns("E").ColumnWidth = 8
        ws.Columns("F").ColumnWidth = 12
        ws.Columns("G").ColumnWidth = 10
        ws.Columns("H").ColumnWidth = 18

        ' Figer la première ligne
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = True

        ' Protéger l'onglet en lecture seule (l'utilisateur ne doit pas modifier)
        ws.Protect Password:="ORBIT_TRACE", DrawingObjects:=True, _
                   Contents:=True, Scenarios:=True, AllowFiltering:=True
    End If

    Set GetOrCreateTraceSheet = ws
End Function

'------------------------------------------------------------------------------
' ClearTraceSheet
' Efface toutes les données de traçabilité (garde l'en-tête).
' Appelé depuis le ruban ORBIT > Outils > Effacer la trace.
'------------------------------------------------------------------------------
Public Sub ClearTraceSheet()
    Dim wb As Workbook: Set wb = GetActiveOrbitWorkbook()
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(TRACE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Onglet de trace introuvable.", vbInformation, "ORBIT": Exit Sub

    If MsgBox("Effacer tout l'historique de traçabilité ?", _
              vbQuestion + vbYesNo, "ORBIT — Confirmer") <> vbYes Then Exit Sub

    ws.Unprotect Password:="ORBIT_TRACE"
    If ws.UsedRange.Rows.Count > 1 Then
        ws.Range("A2:H" & ws.UsedRange.Rows.Count).ClearContents
    End If
    ws.Protect Password:="ORBIT_TRACE", AllowFiltering:=True

    MsgBox "Trace effacée.", vbInformation, "ORBIT"
End Sub

'==============================================================================
' UTILITAIRES
'==============================================================================

Public Function GetAllTablesInWorkbook() As Variant
    Dim wb As Workbook: Set wb = GetActiveOrbitWorkbook()
    If wb Is Nothing Then GetAllTablesInWorkbook = Empty: Exit Function

    Dim ws  As Worksheet
    Dim tbl As ListObject
    Dim result() As String
    Dim count As Integer: count = 0

    For Each ws In wb.Worksheets
        If ws.Name = TRACE_SHEET Then GoTo NextSheet   ' Ignorer l'onglet de trace
        For Each tbl In ws.ListObjects
            ReDim Preserve result(count, 1)
            result(count, 0) = ws.Name
            result(count, 1) = tbl.Name
            count = count + 1
        Next tbl
NextSheet:
    Next ws

    If count = 0 Then GetAllTablesInWorkbook = Empty Else GetAllTablesInWorkbook = result
End Function
