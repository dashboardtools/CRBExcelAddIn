'==============================================================================
' UserForm : frmConsolidatedParams
' Rôle     : Formulaire UNIQUE pour la saisie de TOUS les paramètres
'            de toutes les requêtes du classeur (EDB §6 — dialog unifié).
'
' Fonctionnalités :
'   - Contrôles adaptés au type (date, liste déroulante, texte libre)
'   - Validation en temps réel (dates, numériques)
'   - Listes de valeurs autorisées (ComboBox) chargées depuis SharePoint
'   - Rechargement des dernières valeurs utilisées
'   - Affichage du nombre de tableaux concernés
'
' INSTALLATION :
'   1. Insertion > UserForm → nommer "frmConsolidatedParams"
'   2. Coller ce code dans le module du formulaire
'   3. Les contrôles sont créés dynamiquement dans InitForm()
'==============================================================================
Option Explicit

Public  ParamArray_  As Variant
Public  Cancelled    As Boolean
Private mControls()  As Object     ' Controls dynamiques
Private mAllowedVals() As Variant  ' Listes de valeurs autorisées par paramètre
Private mCount       As Integer

' Constantes de mise en page
Private Const ROW_H  As Integer = 32
Private Const LBL_W  As Integer = 160
Private Const CTL_W  As Integer = 240
Private Const LEFT_M As Integer = 14
Private Const TOP_M  As Integer = 44

'------------------------------------------------------------------------------
' InitForm
' Construit dynamiquement le formulaire selon le tableau consolidé.
' params  : tableau 2D (nom, type, valeur)
' title   : titre affiché dans la caption du formulaire
'------------------------------------------------------------------------------
Public Sub InitForm(ByRef params As Variant, ByVal title As String)
    ParamArray_ = params
    Cancelled   = True
    mCount      = UBound(params, 1) + 1

    Me.Caption = "ORBIT — Saisie des paramètres"
    Me.Width   = LEFT_M * 2 + LBL_W + CTL_W + 30
    Me.Height  = TOP_M + mCount * ROW_H + 110

    ' Supprimer les anciens contrôles dynamiques
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If Left(ctrl.Name, 3) = "dyn" Then Me.Controls.Remove ctrl.Name
    Next ctrl

    ' ── Sous-titre ─────────────────────────────────────────────────────────
    With Me.Controls.Add("Forms.Label.1", "dynSubTitle")
        .Caption   = title
        .Left      = LEFT_M
        .Top       = 8
        .Width     = LBL_W + CTL_W
        .Height    = 20
        .Font.Bold = True
        .Font.Size = 9
        .ForeColor = RGB(31, 73, 125)
    End With

    ReDim mControls(mCount - 1)
    ReDim mAllowedVals(mCount - 1)

    Dim i   As Integer
    Dim y   As Integer: y = TOP_M

    For i = 0 To mCount - 1
        Dim pName  As String: pName  = params(i, 0)
        Dim pType  As String: pType  = params(i, 1)
        Dim pValue As String: pValue = CStr(params(i, 2))

        ' ── Label ────────────────────────────────────────────────────────────
        With Me.Controls.Add("Forms.Label.1", "dynLbl" & i)
            .Caption   = pName
            .Left      = LEFT_M
            .Top       = y + 7
            .Width     = LBL_W
            .Height    = 18
            .Font.Size = 9
        End With

        ' ── Indicateur de type ───────────────────────────────────────────────
        Dim typeColor As Long
        Select Case LCase(pType)
            Case "date":    typeColor = RGB(0, 112, 192)
            Case "list":    typeColor = RGB(112, 48, 160)
            Case "int", "decimal": typeColor = RGB(0, 176, 80)
            Case Else:      typeColor = RGB(128, 128, 128)
        End Select
        With Me.Controls.Add("Forms.Label.1", "dynTypeLbl" & i)
            .Caption   = "[" & pType & "]"
            .Left      = LEFT_M
            .Top       = y + 18
            .Width     = LBL_W
            .Height    = 12
            .Font.Size = 7
            .ForeColor = typeColor
        End With

        ' ── Contrôle de saisie adapté au type ────────────────────────────────
        Dim allowedVals As Variant: allowedVals = GetAllowedValues(pName)
        mAllowedVals(i) = allowedVals

        Dim ctl As Object

        If Not IsEmpty(allowedVals) And UBound(allowedVals) >= 0 Then
            ' Liste de valeurs autorisées → ComboBox
            Set ctl = Me.Controls.Add("Forms.ComboBox.1", "dynCtl" & i)
            ctl.Left  = LEFT_M + LBL_W + 6
            ctl.Top   = y
            ctl.Width = CTL_W
            ctl.Height = 22
            ctl.Style = 0   ' fmStyleDropDownCombo — permet saisie libre aussi

            Dim v As Variant
            For Each v In allowedVals
                If Trim(CStr(v)) <> "" Then ctl.AddItem Trim(CStr(v))
            Next v

            If pValue <> "" Then ctl.Text = pValue _
            Else If ctl.ListCount > 0 Then ctl.ListIndex = 0

        ElseIf LCase(pType) = "date" Then
            ' Date → TextBox avec format ISO
            Set ctl = Me.Controls.Add("Forms.TextBox.1", "dynCtl" & i)
            ctl.Left            = LEFT_M + LBL_W + 6
            ctl.Top             = y
            ctl.Width           = CTL_W
            ctl.Height          = 22
            ctl.ControlTipText  = "Format : JJ/MM/AAAA ou AAAA-MM-JJ"
            If pValue <> "" Then ctl.Text = pValue _
            Else ctl.Text = Format(Date, "DD/MM/YYYY")

        Else
            ' Texte libre
            Set ctl = Me.Controls.Add("Forms.TextBox.1", "dynCtl" & i)
            ctl.Left   = LEFT_M + LBL_W + 6
            ctl.Top    = y
            ctl.Width  = CTL_W
            ctl.Height = 22
            If pType = "list" Then ctl.ControlTipText = "Valeurs séparées par des virgules : ABC, DEF"
            ctl.Text = pValue
        End If

        Set mControls(i) = ctl
        y = y + ROW_H
    Next i

    ' ── Séparateur ────────────────────────────────────────────────────────────
    y = y + 4
    With Me.Controls.Add("Forms.Label.1", "dynSep")
        .Caption        = ""
        .Left           = LEFT_M
        .Top            = y
        .Width          = LBL_W + CTL_W
        .Height         = 1
        .BackColor      = RGB(200, 200, 200)
        .BorderStyle    = 1
    End With
    y = y + 8

    ' ── Boutons ────────────────────────────────────────────────────────────────
    With Me.Controls.Add("Forms.CommandButton.1", "dynBtnOK")
        .Caption   = "✓  Valider et exécuter"
        .Left      = LEFT_M + LBL_W + 6
        .Top       = y
        .Width     = 140
        .Height    = 28
        .BackColor = RGB(31, 73, 125)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    With Me.Controls.Add("Forms.CommandButton.1", "dynBtnCancel")
        .Caption = "Annuler"
        .Left    = LEFT_M + LBL_W + 152
        .Top     = y
        .Width   = 90
        .Height  = 28
    End With

    Me.Height = y + 70
End Sub

'------------------------------------------------------------------------------
' BtnOK_Click — à lier au bouton dynBtnOK dans l'IDE VBA
'------------------------------------------------------------------------------
Public Sub BtnOK_Click()
    If Not HarvestAndValidate() Then Exit Sub
    Cancelled = False
    Me.Hide
End Sub

'------------------------------------------------------------------------------
' BtnCancel_Click — à lier au bouton dynBtnCancel dans l'IDE VBA
'------------------------------------------------------------------------------
Public Sub BtnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancelled = True
End Sub

'------------------------------------------------------------------------------
' HarvestAndValidate
' Lit les valeurs de tous les contrôles.
' Valide : champs obligatoires, format dates, cohérence.
' Retourne False si une erreur est détectée.
'------------------------------------------------------------------------------
Private Function HarvestAndValidate() As Boolean
    Dim i       As Integer
    Dim missing As String
    Dim invalid As String

    For i = 0 To mCount - 1
        Dim ctl    As Object:  Set ctl = mControls(i)
        Dim pName  As String:  pName   = ParamArray_(i, 0)
        Dim pType  As String:  pType   = ParamArray_(i, 1)
        Dim val    As String:  val     = Trim(ctl.Text)

        ' Champ vide
        If val = "" Then
            missing = missing & "  • " & pName & vbNewLine
        Else
            ' Validation type
            Select Case LCase(pType)
                Case "date"
                    If Not IsDate(val) Then
                        invalid = invalid & "  • " & pName & " : date invalide (""" & val & """)" & vbNewLine
                    End If
                Case "int"
                    If Not IsNumeric(val) Or InStr(val, ".") > 0 Then
                        invalid = invalid & "  • " & pName & " : entier attendu" & vbNewLine
                    End If
                Case "decimal"
                    If Not IsNumeric(val) Then
                        invalid = invalid & "  • " & pName & " : valeur numérique attendue" & vbNewLine
                    End If
            End Select

            ' Validation liste autorisée
            Dim av As Variant: av = mAllowedVals(i)
            If Not IsEmpty(av) Then
                Dim found As Boolean: found = False
                Dim v As Variant
                For Each v In av
                    If LCase(Trim(CStr(v))) = LCase(val) Then found = True: Exit For
                Next v
                If Not found Then
                    invalid = invalid & "  • " & pName & " : valeur """ & val & """ non autorisée" & vbNewLine
                End If
            End If

            ParamArray_(i, 2) = val
        End If
    Next i

    Dim errMsg As String
    If missing <> "" Then errMsg = "Champs obligatoires manquants :" & vbNewLine & missing & vbNewLine
    If invalid <> "" Then errMsg = errMsg & "Valeurs invalides :" & vbNewLine & invalid

    If errMsg <> "" Then
        MsgBox errMsg, vbExclamation, "ORBIT — Erreurs de saisie"
        HarvestAndValidate = False
    Else
        HarvestAndValidate = True
    End If
End Function
