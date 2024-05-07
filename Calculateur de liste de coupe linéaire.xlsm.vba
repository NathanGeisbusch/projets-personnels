Rem Attribute VBA_ModuleType=VBAModule

Option VBASupport 1

Private Type LongueurDemandee
    quantite As Integer
    longueur As Integer
End Type

Rem QuickSort décroissant pour le type "LongueurDemandee"
Private Sub ReverseQuickSort(arr() As LongueurDemandee, inLow As Long, inHigh As Long)
  Dim pivot   As Variant
  Dim tmpSwap As LongueurDemandee
  Dim tmpLow  As Long
  Dim tmpHigh As Long
  tmpLow = inLow
  tmpHigh = inHigh
  pivot = arr((inLow + inHigh) \ 2).longueur
  While (tmpLow <= tmpHigh)
     While (arr(tmpLow).longueur > pivot And tmpLow < inHigh)
        tmpLow = tmpLow + 1
     Wend
     While (pivot > arr(tmpHigh).longueur And tmpHigh > inLow)
        tmpHigh = tmpHigh - 1
     Wend
     If (tmpLow <= tmpHigh) Then
        tmpSwap.quantite = arr(tmpLow).quantite
        tmpSwap.longueur = arr(tmpLow).longueur
        arr(tmpLow).quantite = arr(tmpHigh).quantite
        arr(tmpLow).longueur = arr(tmpHigh).longueur
        arr(tmpHigh).quantite = tmpSwap.quantite
        arr(tmpHigh).longueur = tmpSwap.longueur
        tmpLow = tmpLow + 1
        tmpHigh = tmpHigh - 1
     End If
  Wend
  If (inLow < tmpHigh) Then ReverseQuickSort arr, inLow, tmpHigh
  If (tmpLow < inHigh) Then ReverseQuickSort arr, tmpLow, inHigh
End Sub

Public Sub calculer_liste_coupe()
    Rem Effaçage d'un éventuel résultat précédent
    [C5].Value2 = ""
    [B7:B86].Value2 = ""
    Application.ScreenUpdating = False

    Rem Longueur de référence
    Dim lRef As Integer
    lRef = [C4].Value2

    Rem Gestion erreur "référence invalide"
    If IsEmpty(lRef) Or Not (IsNumeric(lRef)) Then
        MsgBox "Erreur: Longueur de référence invalide."
        Exit Sub
    End If

    Rem Récupération des cellules contenant les longueurs demandées
    Dim cells As Variant
    Dim cells_size As Long
    cells = [G5:H86].Value2
    cells_size = UBound(cells)
    
    Rem Vérification du nombre de lignes de cellules remplies
    Dim size As Integer
    size = 0
    For i = 1 To cells_size
        If IsEmpty(cells(i, 1)) Or IsEmpty(cells(i, 2)) Then Exit For
        If Not (IsNumeric(cells(i, 1))) Or Not (IsNumeric(cells(i, 2))) Then Exit For
        size = i
    Next

    Rem Gestion erreur "size = 0"
    If size = 0 Then
        MsgBox "Erreur: Aucune longueur encodée."
        Exit Sub
    End If
    
    Rem Création d'un tableau contenant les longueurs
    Dim errorValueTooHigh As Boolean
    Dim errorQuantityZero As Boolean
    Dim longueurs() As LongueurDemandee
    ReDim Preserve longueurs(size - 1)
    errorValueTooHigh = False
    errorQuantityZero = False
    For i = 1 To size
        longueurs(i - 1).quantite = cells(i, 1)
        longueurs(i - 1).longueur = cells(i, 2)
        If longueurs(i - 1).longueur > lRef Then
            errorValueTooHigh = True
            Exit For
        End If
        If longueurs(i - 1).quantite = 0 Then
            errorQuantityZero = True
            Exit For
        End If
    Next

    Rem Gestion erreur "errorValueTooHigh"
    If errorValueTooHigh Then
        MsgBox "Erreur: Une des longueurs est supérieure à la longueur de référence."
        Exit Sub
    End If

    Rem Gestion erreur "errorQuantityZero"
    If errorQuantityZero Then
        MsgBox "Erreur: Une des quantité est égale à zero."
        Exit Sub
    End If

    Rem Libération mémoire du tableau de cellules
    Erase cells

    Rem Tri par ordre décroissant
    Call ReverseQuickSort(longueurs, 0, size - 1)
    
    Rem Calcul
    Rem tableau de résultats
    Dim resultat() As Variant
    Rem sous-tableau de résultats
    Dim resultTmp() As Integer
    Rem nombre de longueurs restantes
    Dim nbLongRestantes As Integer
    nbLongRestantes = size
    Rem dernier index d'un tableau
    Dim resultLastIndex As Integer
    Rem somme des longueurs
    Dim somme As Integer
    While nbLongRestantes > 0
        ReDim Preserve resultTmp(0)
        somme = 0
        Rem sélection de longueurs jusqu'à atteindre la longueur de référence
        Rem For Each demande In longueurs
        For i = 0 To (size - 1)
            While longueurs(i).quantite > 0 And (somme + longueurs(i).longueur) <= lRef
                somme = somme + longueurs(i).longueur
                longueurs(i).quantite = longueurs(i).quantite - 1
                If longueurs(i).quantite = 0 Then nbLongRestantes = nbLongRestantes - 1
                resultLastIndex = UBound(resultTmp)
                resultTmp(resultLastIndex) = longueurs(i).longueur
                ReDim Preserve resultTmp(resultLastIndex + 1)
            Wend
        Next
        Rem Next demande
        Rem Ajout de la perte à la fin du tableau
        resultTmp(UBound(resultTmp)) = lRef - somme
        Rem Ajout au résultat
        If (Not resultat) = -1 Then
            resultLastIndex = 0
        Else
            resultLastIndex = UBound(resultat) + 1
        End If
        ReDim Preserve resultat(resultLastIndex)
        resultat(resultLastIndex) = resultTmp
    Wend

    Rem Affichage du résultat
    Dim lineResult As String
    Dim sizeResult As Long
    Dim sizeResultTmp As Long
    sizeResult = UBound(resultat)
    [C5].Value2 = sizeResult + 1
    For i = 0 To sizeResult
        sizeResultTmp = UBound(resultat(i)) - 1
        lineResult = "("
        For j = 0 To sizeResultTmp
            lineResult = lineResult & " 1x" & resultat(i)(j)
        Next
        lineResult = lineResult & " ) Perte " & resultat(i)(sizeResultTmp + 1)
        Range("B" & (7 + i)).Value2 = lineResult
    Next
    Application.ScreenUpdating = True

End Sub
