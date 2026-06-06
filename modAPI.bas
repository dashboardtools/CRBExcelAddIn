'==============================================================================
' Module : modAPI
' Rôle   : Appels APIs REST avec paramètres dynamiques {{Param}}.
'          Supporte JSON et CSV en réponse.
'          Authentication : Bearer token (Azure AD) ou Windows SSO.
'
' Utilisé pour : services de marché, APIs performance, ESG, benchmarks, etc.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' CallApi
' Appelle un endpoint REST, substitute les paramètres dans l'URL et le body,
' retourne un tableau 2D Variant (compatible avec InjectDataTable).
'
' La définition de la query SharePoint doit avoir :
'   SourceType  = "API"
'   ApiEndpoint = "https://api.example.com/v1/perf?date={{date}}&fund={{fund}}"
'------------------------------------------------------------------------------
Public Function CallApi(ByVal endpointTemplate As String, _
                         ByVal pa As Variant, _
                         Optional ByVal method As String = "GET", _
                         Optional ByVal bodyTemplate As String = "") As Variant
    On Error GoTo ApiError

    ' 1. Substituer les paramètres dans l'URL
    Dim finalUrl  As String: finalUrl  = SubstituteUrlParams(endpointTemplate, pa)
    Dim finalBody As String: finalBody = ""
    If bodyTemplate <> "" Then finalBody = SubstituteUrlParams(bodyTemplate, pa)

    OrbitLog "API", finalUrl, "Appel " & method

    ' 2. Appel HTTP
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open method, finalUrl, False

    ' Headers standards
    http.setRequestHeader "Accept",       "application/json"
    http.setRequestHeader "Content-Type", "application/json"
    http.setOption 2, 13056   ' Ignorer erreurs cert SSL intranet

    ' Authentification Bearer si token disponible
    Dim token As String: token = GetOrbitSetting("Auth", "BearerToken", "")
    If token <> "" Then
        http.setRequestHeader "Authorization", "Bearer " & token
    End If

    http.send IIf(method = "GET", "", finalBody)

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 5001, "CallApi", _
            "HTTP " & http.Status & " : " & http.statusText & vbNewLine & finalUrl
    End If

    ' 3. Parser la réponse
    Dim responseText As String: responseText = http.responseText

    ' Détecter le format de réponse
    If Left(Trim(responseText), 1) = "[" Or Left(Trim(responseText), 1) = "{" Then
        CallApi = ParseJsonResponse(responseText)
    Else
        ' Supposer CSV
        CallApi = ParseCsvResponse(responseText)
    End If

    OrbitLog "API-OK", finalUrl, "Réponse reçue"
    Exit Function

ApiError:
    OrbitLog "API-ERR", endpointTemplate, Err.Description
    Dim errResult(0, 0) As Variant
    errResult(0, 0) = "ERREUR API : " & Err.Description
    CallApi = errResult
End Function

'------------------------------------------------------------------------------
' SubstituteUrlParams
' Remplace {{Param}} dans une URL ou un body JSON.
' Encode les valeurs pour une URL (les espaces → %20, etc.)
'------------------------------------------------------------------------------
Private Function SubstituteUrlParams(ByVal template As String, _
                                      ByVal pa As Variant) As String
    Dim result As String: result = template
    If IsEmpty(pa) Then SubstituteUrlParams = result: Exit Function

    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True: rx.IgnoreCase = True

    Dim i As Integer
    For i = 0 To UBound(pa, 1)
        rx.Pattern = "\{\{" & pa(i, 0) & "(?::\w+)?\}\}"
        ' Encoder la valeur pour URL uniquement si dans un contexte URL
        Dim encoded As String: encoded = UrlEncode(CStr(pa(i, 2)))
        result = rx.Replace(result, encoded)
    Next i

    SubstituteUrlParams = result
End Function

'------------------------------------------------------------------------------
' ParseJsonResponse
' Parse une réponse JSON de type array d'objets → tableau 2D.
' Supporte les formats :
'   [{"col1":"v1","col2":"v2"}, ...]
'   {"data": [{"col1":"v1"}, ...]}
'   {"results": [...]}  (SharePoint-like)
'   {"value": [...]}    (OData)
'------------------------------------------------------------------------------
Private Function ParseJsonResponse(ByVal json As String) As Variant
    ' Trouver le premier tableau JSON
    Dim arrayStart As Long
    arrayStart = InStr(json, "[")
    If arrayStart = 0 Then
        Dim noData(0, 0) As Variant: noData(0, 0) = "(réponse vide)": ParseJsonResponse = noData: Exit Function
    End If

    ' Extraire le contenu entre [ et ] correspondant
    Dim arrayContent As String
    arrayContent = ExtractJsonArray(json, arrayStart)
    If arrayContent = "" Then
        Dim empty2(0, 0) As Variant: empty2(0, 0) = "(tableau vide)": ParseJsonResponse = empty2: Exit Function
    End If

    ' Parser chaque objet {} du tableau
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "\{[^{}]+\}"
    rx.Global  = True

    Dim rowMatches As Object: Set rowMatches = rx.Execute(arrayContent)
    If rowMatches.Count = 0 Then
        Dim empty3(0, 0) As Variant: empty3(0, 0) = "(aucun résultat)": ParseJsonResponse = empty3: Exit Function
    End If

    ' Extraire les colonnes depuis le premier objet
    Dim colNames() As String
    colNames = ExtractJsonKeys(rowMatches(0).Value)
    Dim colCount As Integer: colCount = UBound(colNames) + 1
    Dim rowCount As Integer: rowCount = rowMatches.Count

    ReDim result(rowCount, colCount - 1) As Variant

    ' En-têtes
    Dim c As Integer
    For c = 0 To colCount - 1: result(0, c) = colNames(c): Next c

    ' Données
    Dim rowRx As Object: Set rowRx = CreateObject("VBScript.RegExp")
    rowRx.Global = True

    Dim r As Integer
    For r = 0 To rowCount - 1
        Dim rowJson As String: rowJson = rowMatches(r).Value
        For c = 0 To colCount - 1
            result(r + 1, c) = ExtractJsonString(rowJson, colNames(c))
        Next c
    Next r

    ParseJsonResponse = result
End Function

'------------------------------------------------------------------------------
' ParseCsvResponse
' Parse une réponse CSV (première ligne = en-têtes) → tableau 2D.
'------------------------------------------------------------------------------
Private Function ParseCsvResponse(ByVal csv As String) As Variant
    ' Normaliser les sauts de ligne
    csv = Replace(csv, vbCrLf, vbLf): csv = Replace(csv, vbCr, vbLf)
    csv = Trim(csv)

    Dim lines() As String: lines = Split(csv, vbLf)
    If UBound(lines) < 0 Then
        Dim e(0, 0) As Variant: e(0, 0) = "(CSV vide)": ParseCsvResponse = e: Exit Function
    End If

    Dim headers() As String: headers = Split(lines(0), ",")
    Dim colCount  As Integer: colCount = UBound(headers) + 1
    Dim rowCount  As Integer: rowCount = UBound(lines)   ' lignes données (sans en-tête)

    ReDim result(rowCount, colCount - 1) As Variant

    Dim c As Integer
    For c = 0 To colCount - 1
        result(0, c) = Trim(Replace(Replace(headers(c), """", ""), vbLf, ""))
    Next c

    Dim r As Integer
    For r = 1 To UBound(lines)
        If Trim(lines(r)) = "" Then GoTo NextLine
        Dim cells() As String: cells = Split(lines(r), ",")
        For c = 0 To colCount - 1
            If c <= UBound(cells) Then
                result(r, c) = Trim(Replace(cells(c), """", ""))
            End If
        Next c
NextLine:
    Next r

    ParseCsvResponse = result
End Function

'==============================================================================
' HELPERS JSON
'==============================================================================

Private Function ExtractJsonArray(ByVal json As String, ByVal startPos As Long) As String
    Dim depth As Integer: depth = 0
    Dim i As Long
    For i = startPos To Len(json)
        Dim ch As String: ch = Mid(json, i, 1)
        If ch = "[" Then depth = depth + 1
        If ch = "]" Then
            depth = depth - 1
            If depth = 0 Then
                ExtractJsonArray = Mid(json, startPos + 1, i - startPos - 1)
                Exit Function
            End If
        End If
    Next i
    ExtractJsonArray = ""
End Function

Private Function ExtractJsonKeys(ByVal obj As String) As String()
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = """(\w+)""\s*:"
    rx.Global  = True
    Dim m As Object: Set m = rx.Execute(obj)

    If m.Count = 0 Then
        Dim empty(0) As String: empty(0) = "value": ExtractJsonKeys = empty: Exit Function
    End If

    ReDim keys(m.Count - 1) As String
    Dim i As Integer
    For i = 0 To m.Count - 1: keys(i) = m(i).SubMatches(0): Next i
    ExtractJsonKeys = keys
End Function

'------------------------------------------------------------------------------
' UrlEncode — encode une chaîne pour inclusion dans une URL
'------------------------------------------------------------------------------
Public Function UrlEncode(ByVal s As String) As String
    Dim i As Integer
    Dim result As String
    Dim ch As String
    Dim code As Integer

    For i = 1 To Len(s)
        ch   = Mid(s, i, 1)
        code = Asc(ch)
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                result = result & ch          ' Caractères alphanumériques et -._~
            Case 32
                result = result & "%20"       ' Espace
            Case Else
                result = result & "%" & Hex(code)
        End Select
    Next i

    UrlEncode = result
End Function
