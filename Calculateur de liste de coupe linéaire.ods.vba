REM  *****  BASIC  *****

Option VBASupport 1
Option Compatible

Private Type LongueurDemandee
	quantite As Integer
	longueur As Integer
End Type

REM QuickSort décroissant pour le type "LongueurDemandee"
Private Sub ReverseQuickSort(arr As LongueurDemandee, inLow As Long, inHigh As Long)
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

Public Sub calculer_liste_coupe
	REM Effaçage d'un éventuel résultat précédent
	[C5].Value = ""
	For i = 7 To 86
		Range("B"&i).Value = ""
	Next

	REM Longueur de référence
	Dim lRef As Integer
	lRef = [C4].Value

	REM Gestion erreur "référence invalide"
	If IsEmpty(lRef) Or Not(IsNumeric(lRef)) Then
		MsgBox "Erreur: Longueur de référence invalide."
		Exit Sub
	End If

	REM Récupération des cellules contenant les longueurs demandées
	Dim cells as Variant
	Dim cells_size As Long
	cells = [G5:H86].Value
	cells_size = UBound(cells)
	
	REM Vérification du nombre de lignes de cellules remplies
	Dim size as Integer
	size = 0
	For i = 1 To cells_size
		If IsEmpty(cells(i,1)) Or IsEmpty(cells(i,2)) Then Exit For
		If Not(IsNumeric(cells(i,1))) Or Not(IsNumeric(cells(i,2))) Then Exit For
		size = i
	Next

	REM Gestion erreur "size = 0"
	If size = 0 Then
		MsgBox "Erreur: Aucune longueur encodée."
		Exit Sub
	End If
	
	REM Création d'un tableau contenant les longueurs
	Dim errorValueTooHigh As Boolean
	Dim errorQuantityZero As Boolean
	Dim longueurs(size-1) As LongueurDemandee
	errorValueTooHigh = False
	errorQuantityZero = False
	For i = 1 To size
		longueurs(i-1).quantite = cells(i,1)
		longueurs(i-1).longueur = cells(i,2)
		If longueurs(i-1).longueur > lRef Then
			errorValueTooHigh = True
			Exit For
		End If
		If longueurs(i-1).quantite = 0 Then
			errorQuantityZero = True
			Exit For
		End If
	Next

	REM Gestion erreur "errorValueTooHigh"
	If errorValueTooHigh Then
		MsgBox "Erreur: Une des longueurs est supérieure à la longueur de référence."
		Exit Sub
	End If

	REM Gestion erreur "errorQuantityZero"
	If errorQuantityZero Then
		MsgBox "Erreur: Une des quantité est égale à zero."
		Exit Sub
	End If

	REM Libération mémoire du tableau de cellules
	Erase cells

	REM Tri par ordre décroissant
	Call ReverseQuickSort(longueurs, 0, size-1)
	
	REM Calcul
	Dim resultat() As Variant REM tableau de résultats
	Dim resultTmp() As Integer REM sous-tableau de résultats
	Dim nbLongRestantes As Integer REM nombre de longueurs restantes
	nbLongRestantes = size
	Dim resultLastIndex As Integer REM dernier index d'un tableau
	Dim somme As Integer REM somme des longueurs
	While nbLongRestantes > 0
		ReDim Preserve resultTmp(0)
		somme = 0
		REM sélection de longueurs jusqu'à atteindre la longueur de référence
		For Each demande In longueurs
			While demande.quantite > 0 And (somme+demande.longueur) <= lRef
				somme = somme + demande.longueur
				demande.quantite = demande.quantite - 1
				If demande.quantite = 0 Then nbLongRestantes = nbLongRestantes - 1
				resultLastIndex = UBound(resultTmp)
				resultTmp(resultLastIndex) = demande.longueur
				ReDim Preserve resultTmp(resultLastIndex+1)
			Wend
		Next demande
		REM Ajout de la perte à la fin du tableau
		resultTmp(UBound(resultTmp)) = lRef - somme
		REM Ajout au résultat
		resultLastIndex = UBound(resultat)+1
		ReDim Preserve resultat(resultLastIndex)
		resultat(resultLastIndex) = resultTmp
	Wend

	REM Affichage du résultat
	Dim lineResult As String
	Dim sizeResult As Long
	Dim sizeResultTmp As Long
	sizeResult = UBound(resultat)
	[C5].Value = sizeResult+1
	For i = 0 To sizeResult
		sizeResultTmp = UBound(resultat(i))-1
		lineResult = "("
		For j = 0 To sizeResultTmp
			lineResult = lineResult & " 1x" & resultat(i)(j)
		Next
		lineResult = lineResult & " ) Perte " & resultat(i)(sizeResultTmp+1)
		Range("B"&(7+i)).Value = lineResult
	Next

End Sub
