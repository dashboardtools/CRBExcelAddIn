'==============================================================================
' Module : modDysqualite
' Rôle   : Gestion complète du signalement des incidents de qualité de données.
'
' EDB §15 — Fonctionnalités :
'   1. Capture automatique du contexte (fichier, onglet, requête, paramètres)
'   2. Formulaire de déclaration (criticité, description libre)
'   3. Génération email Outlook prérempli vers contacts DQ SharePoint
'   4. Historisation dans liste SharePoint DataQualityIssues
'   5. Log local de l'incident
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' SignalerDysqualite
' Point d'entrée appelé par le bouton ruban "Signaler une dysqualité".
'------------------------------------------------------------------------------
Public Sub SignalerDysqualite()
    ' 1. Capturer le contexte automatiquement
    Dim ctx As Object: Set ctx = CaptureContext()

    ' 2. Récupérer la requête liée au tableau actif
    Dim queryId    As String: queryId    = ""
    Dim paramsJson As String: paramsJson = "{}"
    Dim sourceType As String: sourceType = "SQL"
    Dim endpoint   As String: endpoint   = ""

    Dim binding As Object: Set binding = FindBinding(ctx("ActiveSheet"), ctx("ActiveTable"))
    If Not binding Is Nothing Then
        queryId = binding("queryId")
        ' Essayer de retrouver les derniers paramètres utilisés
        Dim q As Object: Set q = FindQuery(queryId)
        If Not q Is Nothing Then
            sourceType = q("SourceType")
            endpoint   = q("ApiEndpoint")
            Dim pa As Variant: pa = ParseParameters(q("SQLQuery"))
            If Not IsEmpty(pa) Then
                LoadLastConsolidatedParams pa
                paramsJson = ParametersToJson(pa)
            End If
        End If
    End If

    ' 3. Afficher le formulaire de déclaration
    frmDysqualite.InitForm ctx, queryId, paramsJson
    frmDysqualite.Show vbModal

    If frmDysqualite.Cancelled Then Exit Sub

    Dim description As String: description = frmDysqualite.Description
    Dim criticality As String: criticality = frmDysqualite.Criticality

    ' 4. Enregistrer dans SharePoint DataQualityIssues
    Dim incidentId As String
    incidentId = CreateDQIssue(ctx, queryId, paramsJson, description, criticality)

    ' 5. Préparer et ouvrir l'email Outlook
    PrepareDQEmail ctx, queryId, paramsJson, description, criticality, incidentId, endpoint

    ' 6. Log local
    OrbitLog "DQ", queryId, "Incident " & incidentId & " — " & criticality

    ' 7. Confirmation utilisateur
    Dim spStatus As String
    If incidentId <> "" Then
        spStatus = "✓ Enregistré dans SharePoint (ID : " & incidentId & ")"
    Else
        spStatus = "⚠ SharePoint inaccessible — incident non enregistré centralement"
    End If

    MsgBox "Signalement envoyé." & vbNewLine & vbNewLine & _
           spStatus & vbNewLine & _
           "✓ Email Outlook prérempli ouvert.", _
           vbInformation, "ORBIT — Dysqualité signalée"
End Sub

'------------------------------------------------------------------------------
' PrepareDQEmail
' Crée un email Outlook prérempli, non envoyé (l'utilisateur peut le compléter).
' Les destinataires sont lus depuis la liste SharePoint DataQualityContacts.
'------------------------------------------------------------------------------
Private Sub PrepareDQEmail(ByVal ctx As Object, _
                            ByVal queryId As String, _
                            ByVal paramsJson As String, _
                            ByVal description As String, _
                            ByVal criticality As String, _
                            ByVal incidentId As String, _
                            ByVal endpoint As String)
    On Error GoTo OutlookError

    ' Récupérer les contacts DQ depuis SharePoint
    Dim contacts As Collection: Set contacts = GetDQContacts()
    Dim toList   As String

    If contacts.Count > 0 Then
        Dim contact As Object
        For Each contact In contacts
            If toList <> "" Then toList = toList & ";"
            toList = toList & contact("Email")
        Next contact
    Else
        ' Fallback : adresse configurée dans les settings
        toList = GetOrbitSetting("DQ", "DefaultContactEmail", "data-governance@votreentreprise.com")
    End If

    ' Construire le sujet
    Dim subject As String
    Dim critTag As String
    Select Case LCase(criticality)
        Case "critique": critTag = "🔴 CRITIQUE"
        Case "moyen":    critTag = "🟡 MOYEN"
        Case Else:       critTag = "🟢 FAIBLE"
    End Select

    subject = "[DQ ALERT] " & critTag & " — Suspicion dysqualité"
    If queryId <> "" Then subject = subject & " — " & queryId
    subject = subject & " — " & Format(Date, "DD/MM/YYYY")

    ' Construire le corps du mail
    Dim body As String
    body = "Bonjour," & vbCrLf & vbCrLf & _
           "Une suspicion de dysqualité a été identifiée dans un fichier Excel " & _
           "alimenté via le plugin ORBIT Data Access." & vbCrLf & vbCrLf & _
           "══════════════════════════════════════════" & vbCrLf & _
           "INFORMATIONS CONTEXTUELLES" & vbCrLf & _
           "══════════════════════════════════════════" & vbCrLf & _
           "Incident ID    : " & IIf(incidentId <> "", incidentId, "Non enregistré") & vbCrLf & _
           "Criticité      : " & criticality & vbCrLf & _
           "Utilisateur    : " & ctx("UserFullName") & " (" & ctx("UserLogin") & ")" & vbCrLf & _
           "Email          : " & ctx("UserEmail") & vbCrLf & _
           "Date/Heure     : " & ctx("TimestampISO") & vbCrLf & _
           "Machine        : " & ctx("MachineName") & vbCrLf & vbCrLf & _
           "══════════════════════════════════════════" & vbCrLf & _
           "FICHIER EXCEL" & vbCrLf & _
           "══════════════════════════════════════════" & vbCrLf & _
           "Fichier        : " & ctx("WorkbookName") & vbCrLf & _
           "Répertoire     : " & ctx("WorkbookDir") & vbCrLf & _
           "Onglet         : " & ctx("ActiveSheet") & vbCrLf & _
           "Tableau        : " & ctx("ActiveTable") & vbCrLf & vbCrLf & _
           "══════════════════════════════════════════" & vbCrLf & _
           "SOURCE DE DONNÉES" & vbCrLf & _
           "══════════════════════════════════════════" & vbCrLf & _
           "Requête / API  : " & IIf(queryId <> "", queryId, "Non identifiée") & vbCrLf & _
           "Endpoint       : " & IIf(endpoint <> "", endpoint, "N/A") & vbCrLf & _
           "Paramètres     : " & paramsJson & vbCrLf & vbCrLf & _
           "══════════════════════════════════════════" & vbCrLf & _
           "DESCRIPTION DU PROBLÈME" & vbCrLf & _
           "══════════════════════════════════════════" & vbCrLf & _
           description & vbCrLf & vbCrLf & _
           "──────────────────────────────────────────" & vbCrLf & _
           "Merci de bien vouloir analyser ce signalement et revenir vers l'utilisateur." & vbCrLf & _
           "Cet email a été généré automatiquement par le plugin ORBIT v" & _
           GetOrbitSetting("Config", "Version", "2.0") & "." & vbCrLf & vbCrLf & _
           "Cordialement,"

    ' Ouvrir Outlook et préparer le mail (non envoyé)
    Dim olApp  As Object
    Dim olMail As Object

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo OutlookError

    Set olMail = olApp.CreateItem(0)   ' olMailItem

    With olMail
        .To      = toList
        .Subject = subject
        .Body    = body
        .Importance = IIf(LCase(criticality) = "critique", 2, 1)   ' olImportanceHigh / Normal

        ' Pièce jointe optionnelle : configurable
        Dim attachWorkbook As Boolean
        attachWorkbook = (GetOrbitSetting("DQ", "AttachWorkbook", "0") = "1")
        If attachWorkbook And ctx("WorkbookPath") <> "" Then
            If Dir(ctx("WorkbookPath")) <> "" Then
                .Attachments.Add ctx("WorkbookPath")
            End If
        End If

        .Display   ' Ouvrir le mail sans l'envoyer (l'utilisateur peut compléter)
    End With

    Exit Sub

OutlookError:
    OrbitLog "DQ-MAIL-ERR", queryId, Err.Description
    MsgBox "Impossible d'ouvrir Outlook :" & vbNewLine & Err.Description & vbNewLine & vbNewLine & _
           "L'incident a néanmoins été enregistré dans SharePoint (ID : " & incidentId & ").", _
           vbExclamation, "ORBIT — Outlook indisponible"
End Sub
