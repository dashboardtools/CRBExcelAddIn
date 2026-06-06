'==============================================================================
' UserForm : frmDysqualite
' Rôle     : Formulaire de déclaration d'un incident de qualité de données.
'            Contexte pré-rempli automatiquement (fichier, onglet, requête).
'            L'utilisateur saisit uniquement : description + criticité.
'
' INSTALLATION :
'   1. Insertion > UserForm → nommer "frmDysqualite"
'   2. Coller ce code dans le module du formulaire
'   3. Les contrôles sont créés dynamiquement dans InitForm()
'==============================================================================
Option Explicit

Public Cancelled    As Boolean
Public Description  As String
Public Criticality  As String

Private mCtxLabels()  As String
Private mCtxValues()  As String

'------------------------------------------------------------------------------
' InitForm
' Pré-remplit le formulaire avec le contexte capturé automatiquement.
'------------------------------------------------------------------------------
Public Sub InitForm(ByVal ctx As Object, ByVal queryId As String, ByVal paramsJson As String)
    Cancelled   = True
    Description = ""
    Criticality = "Moyen"

    Me.Caption = "ORBIT — Signaler une dysqualité"
    Me.Width   = 520
    Me.Height  = 520

    ' Supprimer anciens contrôles dynamiques
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If Left(c.Name, 3) = "dyn" Then Me.Controls.Remove c.Name
    Next c

    Dim y As Integer: y = 10

    ' ── Titre ────────────────────────────────────────────────────────────────
    With Me.Controls.Add("Forms.Label.1", "dynTitle")
        .Caption   = "🔍  Signalement d'une suspicion de dysqualité"
        .Left      = 12: .Top = y: .Width = 480: .Height = 22
        .Font.Bold = True: .Font.Size = 10
        .ForeColor = RGB(192, 0, 0)
    End With
    y = y + 26

    ' ── Contexte automatique (lecture seule) ─────────────────────────────────
    With Me.Controls.Add("Forms.Label.1", "dynCtxHdr")
        .Caption   = "Contexte détecté automatiquement"
        .Left      = 12: .Top = y: .Width = 480: .Height = 16
        .Font.Bold = True: .Font.Size = 8: .ForeColor = RGB(31, 73, 125)
    End With
    y = y + 18

    ' Fond gris pour la zone contexte
    Dim ctxItems As Variant
    ctxItems = Array( _
        Array("Utilisateur",   ctx("UserLogin") & "  (" & ctx("UserFullName") & ")"), _
        Array("Fichier",       ctx("WorkbookName")), _
        Array("Onglet",        ctx("ActiveSheet")), _
        Array("Tableau",       IIf(ctx("ActiveTable") <> "", ctx("ActiveTable"), "Non détecté")), _
        Array("Requête/API",   IIf(queryId <> "", queryId, "Non identifiée")), _
        Array("Paramètres",    IIf(Len(paramsJson) > 60, Left(paramsJson, 57) & "...", paramsJson)), _
        Array("Horodatage",    ctx("TimestampISO")) _
    )

    Dim item As Variant
    For Each item In ctxItems
        With Me.Controls.Add("Forms.Label.1", "dynCtxL" & y)
            .Caption = item(0) & " :"
            .Left = 16: .Top = y: .Width = 110: .Height = 16
            .Font.Size = 8: .ForeColor = RGB(80, 80, 80): .Font.Bold = True
        End With
        With Me.Controls.Add("Forms.Label.1", "dynCtxV" & y)
            .Caption = CStr(item(1))
            .Left = 130: .Top = y: .Width = 360: .Height = 16
            .Font.Size = 8: .ForeColor = RGB(60, 60, 60)
        End With
        y = y + 18
    Next item

    y = y + 8

    ' ── Séparateur ───────────────────────────────────────────────────────────
    With Me.Controls.Add("Forms.Label.1", "dynSep1")
        .Caption = "": .Left = 12: .Top = y: .Width = 480: .Height = 1
        .BackColor = RGB(192, 0, 0)
    End With
    y = y + 10

    ' ── Criticité ────────────────────────────────────────────────────────────
    With Me.Controls.Add("Forms.Label.1", "dynCritLbl")
        .Caption = "Niveau de criticité *"
        .Left = 12: .Top = y + 3: .Width = 140: .Height = 18
        .Font.Bold = True: .Font.Size = 9
    End With

    With Me.Controls.Add("Forms.ComboBox.1", "dynCriticality")
        .Left = 156: .Top = y: .Width = 160: .Height = 22
        .Style = 2   ' fmStyleDropDownList — liste fermée uniquement
        .AddItem "Faible"
        .AddItem "Moyen"
        .AddItem "Critique"
        .ListIndex = 1   ' Moyen par défaut
    End With
    y = y + 30

    ' ── Description ─────────────────────────────────────────────────────────
    With Me.Controls.Add("Forms.Label.1", "dynDescLbl")
        .Caption = "Description du problème *"
        .Left = 12: .Top = y: .Width = 480: .Height = 16
        .Font.Bold = True: .Font.Size = 9
    End With
    y = y + 18

    With Me.Controls.Add("Forms.Label.1", "dynDescHint")
        .Caption = "Ex : données incohérentes, montant incorrect, données manquantes, doublons, problème de date..."
        .Left = 12: .Top = y: .Width = 480: .Height = 14
        .Font.Size = 7: .Font.Italic = True: .ForeColor = RGB(128, 128, 128)
    End With
    y = y + 16

    With Me.Controls.Add("Forms.TextBox.1", "dynDescription")
        .Left        = 12: .Top = y: .Width = 480: .Height = 70
        .MultiLine   = True
        .WordWrap    = True
        .ScrollBars  = 2   ' Vertical
        .EnterKeyBehavior = True
    End With
    y = y + 78

    ' ── Boutons ──────────────────────────────────────────────────────────────
    With Me.Controls.Add("Forms.CommandButton.1", "dynBtnSend")
        .Caption   = "📧  Signaler"
        .Left      = 200: .Top = y: .Width = 130: .Height = 28
        .BackColor = RGB(192, 0, 0): .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    With Me.Controls.Add("Forms.CommandButton.1", "dynBtnCancel")
        .Caption = "Annuler"
        .Left    = 340: .Top = y: .Width = 90: .Height = 28
    End With

    Me.Height = y + 70
End Sub

' ── Handlers à lier manuellement dans l'IDE ──────────────────────────────────

Public Sub BtnSend_Click()
    Dim descCtl As MSForms.TextBox
    Dim critCtl As MSForms.ComboBox

    Set descCtl = Me.Controls("dynDescription")
    Set critCtl = Me.Controls("dynCriticality")

    If Trim(descCtl.Text) = "" Then
        MsgBox "Merci de décrire le problème constaté.", vbExclamation, "ORBIT"
        Exit Sub
    End If

    Description = Trim(descCtl.Text)
    Criticality = critCtl.Text
    Cancelled   = False
    Me.Hide
End Sub

Public Sub BtnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancelled = True
End Sub
