'==============================================================================
' Module : modParameters
' Rôle   : Consolidation des paramètres de TOUTES les requêtes d'un classeur
'          en un seul dialog unifié — l'utilisateur saisit une seule fois.
'
' EDB §6 : "consolider tous les paramètres, supprimer les doublons,
'           construire automatiquement une boîte de dialogue unique"
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' ConsolidateParameters
' Reçoit une Collection de QueryDefinition, retourne un tableau paramArray
' consolidé (dédoublonné) avec les types les plus stricts en cas de conflit.
'
' Retourne un tableau 2D : (i,0)=Nom  (i,1)=Type  (i,2)=Valeur
'------------------------------------------------------------------------------
Public Function ConsolidateParameters(ByVal queries As Collection) As Variant
    Dim allParams As Object: Set allParams = CreateObject("Scripting.Dictionary")

    Dim q As Object
    For Each q In queries
        ' Prendre SQL ou API endpoint selon le type de source
        Dim template As String
        If LCase(q("SourceType")) = "api" Then
            template = q("ApiEndpoint")
        Else
            template = q("SQLQuery")
        End If

        If template = "" Then GoTo NextQuery

        Dim pa As Variant: pa = ParseParameters(template)
        If IsEmpty(pa) Then GoTo NextQuery

        Dim i As Integer
        For i = 0 To UBound(pa, 1)
            Dim pName As String: pName = LCase(pa(i, 0))
            Dim pType As String: pType = pa(i, 1)

            If Not allParams.Exists(pName) Then
                ' Nouveau paramètre
                Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
                entry("Name")         = pa(i, 0)   ' Conserver la casse originale
                entry("Type")         = pType
                entry("Value")        = ""
                entry("AllowedValues")= GetAllowedValues(pa(i, 0), q("QueryId"))
                allParams(pName)      = entry
            Else
                ' Doublon — prendre le type le plus strict
                Dim existing As Object: Set existing = allParams(pName)
                existing("Type") = MergeTypes(existing("Type"), pType)
                ' Fusionner les listes de valeurs autorisées
                If IsEmpty(existing("AllowedValues")) Then
                    existing("AllowedValues") = GetAllowedValues(pa(i, 0), q("QueryId"))
                End If
            End If
        Next i
NextQuery:
    Next q

    If allParams.Count = 0 Then ConsolidateParameters = Empty: Exit Function

    ' Convertir le dictionnaire en tableau 2D
    ReDim result(allParams.Count - 1, 2) As String
    Dim j As Integer: j = 0
    Dim key As Variant
    For Each key In allParams.Keys
        Dim e As Object: Set e = allParams(key)
        result(j, 0) = e("Name")
        result(j, 1) = e("Type")
        result(j, 2) = ""
        j = j + 1
    Next key

    ConsolidateParameters = result
End Function

'------------------------------------------------------------------------------
' MergeTypes
' Résout le type le plus strict entre deux définitions du même paramètre.
' Ordre de priorité : date > decimal > int > list > string
'------------------------------------------------------------------------------
Private Function MergeTypes(ByVal t1 As String, ByVal t2 As String) As String
    Dim priority As Object: Set priority = CreateObject("Scripting.Dictionary")
    priority("date")    = 5
    priority("decimal") = 4
    priority("int")     = 3
    priority("list")    = 2
    priority("string")  = 1

    Dim p1 As Integer: p1 = IIf(priority.Exists(t1), priority(t1), 1)
    Dim p2 As Integer: p2 = IIf(priority.Exists(t2), priority(t2), 1)
    MergeTypes = IIf(p1 >= p2, t1, t2)
End Function

'------------------------------------------------------------------------------
' InjectConsolidatedParams
' Reçoit le tableau consolidé (valeurs remplies) et réinjecte les bonnes
' valeurs dans chaque requête individuelle.
'
' Retourne une collection de QueryDefinition enrichies avec les valeurs.
'------------------------------------------------------------------------------
Public Function InjectConsolidatedParams(ByVal queries As Collection, _
                                          ByVal consolidated As Variant) As Collection
    Dim result As New Collection

    Dim q As Object
    For Each q In queries
        Dim template As String
        If LCase(q("SourceType")) = "api" Then
            template = q("ApiEndpoint")
        Else
            template = q("SQLQuery")
        End If

        If template = "" Then result.Add q: GoTo NextQ

        ' Parser les params de cette requête spécifique
        Dim pa As Variant: pa = ParseParameters(template)
        If IsEmpty(pa) Then result.Add q: GoTo NextQ

        ' Pour chaque paramètre de cette requête, chercher la valeur dans le consolidé
        Dim i As Integer
        For i = 0 To UBound(pa, 1)
            Dim pName As String: pName = LCase(pa(i, 0))
            Dim j As Integer
            For j = 0 To UBound(consolidated, 1)
                If LCase(consolidated(j, 0)) = pName Then
                    pa(i, 2) = consolidated(j, 2)
                    Exit For
                End If
            Next j
        Next i

        ' Stocker le tableau de params dans la query
        q("ResolvedParams") = pa

        result.Add q
NextQ:
    Next q

    Set InjectConsolidatedParams = result
End Function

'------------------------------------------------------------------------------
' ValidateConsolidatedParams
' Valide le tableau consolidé rempli par l'utilisateur.
' Retourne une Collection de messages d'erreur (vide = tout OK).
'------------------------------------------------------------------------------
Public Function ValidateConsolidatedParams(ByVal pa As Variant) As Collection
    Dim errors As New Collection
    If IsEmpty(pa) Then Set ValidateConsolidatedParams = errors: Exit Function

    Dim i As Integer
    For i = 0 To UBound(pa, 1)
        Dim pName  As String: pName  = pa(i, 0)
        Dim pType  As String: pType  = pa(i, 1)
        Dim pValue As String: pValue = CStr(pa(i, 2))

        ' Champ obligatoire
        If Trim(pValue) = "" Then
            errors.Add "Paramètre obligatoire manquant : " & pName
            GoTo NextParam
        End If

        ' Validation par type
        Select Case LCase(pType)
            Case "date"
                If Not IsDate(pValue) Then
                    errors.Add pName & " : date invalide (format attendu : JJ/MM/AAAA ou AAAA-MM-JJ)"
                End If
            Case "int"
                If Not IsNumeric(pValue) Or InStr(pValue, ".") > 0 Then
                    errors.Add pName & " : entier invalide"
                End If
            Case "decimal"
                If Not IsNumeric(pValue) Then
                    errors.Add pName & " : valeur décimale invalide"
                End If
        End Select

        ' Validation liste autorisée (si définie)
        ' Les AllowedValues sont stockés dans le tableau de contexte à part —
        ' on délègue cette validation au formulaire frmConsolidatedParams
        ' qui a accès aux AllowedValues dans son dictionnaire interne.

NextParam:
    Next i

    Set ValidateConsolidatedParams = errors
End Function

'------------------------------------------------------------------------------
' LoadLastConsolidatedParams
' Recharge les dernières valeurs saisies depuis le registre Windows.
' Clé de stockage : "Consolidated|NomParam"
'------------------------------------------------------------------------------
Public Sub LoadLastConsolidatedParams(ByRef pa As Variant)
    If IsEmpty(pa) Then Exit Sub
    Dim i As Integer
    For i = 0 To UBound(pa, 1)
        Dim saved As String
        saved = GetSetting("ORBIT", "ConsolidatedParams", "param|" & LCase(pa(i, 0)), "")
        If saved <> "" Then pa(i, 2) = saved
    Next i
End Sub

Public Sub SaveConsolidatedParams(ByVal pa As Variant)
    If IsEmpty(pa) Then Exit Sub
    Dim i As Integer
    For i = 0 To UBound(pa, 1)
        SaveSetting "ORBIT", "ConsolidatedParams", "param|" & LCase(pa(i, 0)), CStr(pa(i, 2))
    Next i
End Sub
