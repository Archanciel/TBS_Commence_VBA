Attribute VB_Name = "TBSSimulationMacro"
Option Explicit

'Constantes utilisée pour le calcul du mix de pack
Const idx300 = 0        '300 $
Const idxBronze = 1     '1000 $
Const idxArgent = 2     '2000 $
Const idxOr = 3         '4000 $
Const idxPlatinum = 4   '10000 $

''''''''''''''''''''''''''''''''''
'Procedure de test de la fonction récursive MaxInvest
''''''''''''''''''''''''''''''''''
Sub Calc()
    Dim MaxInv As Variant
    Dim availableAmount As Long
    Dim maxPackAmount As Variant
    
    availableAmount = InputBox("Enter a value for amount available for invest !", "Version récursive")
    maxPackAmount = InputBox("Enter a value for max amount to invest for invest, 0 if no max !", "Version récursive", 0)
    MaxInv = MaxInvest(availableAmount, maxPackAmount)
    MsgBox (MaxInv)
End Sub

''''''''''''''''''''''''''''''''''
'Calcul du mix de packs pour une somme donnée.
'
'Précondition: la somme passée en entrée est fractionnable en packs sans aucun reste.
'              Typiquement, il s'agit d'une somme retournée par la fct MaxInvest !
'              Par exemple, 1100 viole la précondition puisque les packs achetés feront
'              900 au total !
''''''''''''''''''''''''''''''''''
Sub CalcPackMix()
    Dim PackInvest As Variant
    Dim PackTotal As Variant
    Dim PackMixMaxBronzeArray(idxBronze + 1) As Long
    Dim PackMixArray(3, idxPlatinum + 1) As Variant
    Dim PackMixMaxBronzeStr As String
    Dim PackMixStr As String
    
    PackTotal = InputBox("Enter total pack amount !", "CalcPackMix")
    PackMix PackTotal, PackMixMaxBronzeArray
    PackMixMaxBronzeStr = WritePackMixMaxBronzeStr(PackMixMaxBronzeArray)
    
'    MsgBox (PackMixMaxBronzeStr)
    
    'initialisation PackMixArray
    PackMixArray(0, idx300) = 300
    PackMixArray(1, idx300) = "300"
    PackMixArray(2, idx300) = 0
    PackMixArray(0, idxBronze) = 1000
    PackMixArray(1, idxBronze) = "Bronze/1000"
    PackMixArray(2, idxBronze) = 0
    PackMixArray(0, idxArgent) = 2000
    PackMixArray(1, idxArgent) = "Argent/2000"
    PackMixArray(2, idxArgent) = 0
    PackMixArray(0, idxOr) = 4000
    PackMixArray(1, idxOr) = "Or/4000"
    PackMixArray(2, idxOr) = 0
    PackMixArray(0, idxPlatinum) = 10000
    PackMixArray(1, idxPlatinum) = "Platinum/10000"
    PackMixArray(2, idxPlatinum) = 0
    
    computePackMixArray PackMixMaxBronzeArray, PackMixArray
    
    PackMixStr = WritePackMixStr(PackMixArray)
    
    MsgBox PackMixStr, vbOKOnly, "Total: " & PackTotal
End Sub

Private Sub computePackMixArray(PackMixMaxBronzeArray() As Long, ByRef PackMixArray() As Variant)
    Dim packAmount1000 As Long
    Dim packAmount2000 As Long
    Dim packAmount4000 As Long
    Dim packAmount10000 As Long
    Dim i As Integer
    
    packAmount1000 = PackMixMaxBronzeArray(idxBronze) * 1000
    PackMixArray(2, idx300) = PackMixMaxBronzeArray(idx300)
    
    i = idxPlatinum
    
    While (i > idx300)
        If (packAmount1000 >= PackMixArray(0, i)) Then
            PackMixArray(2, i) = Int(packAmount1000 / PackMixArray(0, i))
            packAmount1000 = packAmount1000 - PackMixArray(2, i) * PackMixArray(0, i)
        End If
        
        i = i - 1
    Wend
    
    'fil
End Sub

Function PackMix(ByVal PackTotal As Long, ByRef PackMixMaxBronzeArray() As Long) As Long
    Dim recRes As Long       'résultat du calcul par récursion
    Dim mult300Res As Long   'résultat du calcul par division par 300, pour les cas comme 2100, 2400, 2700, etc
    
    If (PackTotal < 300) Then
        recRes = 0
    ElseIf (PackTotal < 1000) Then
        PackMixMaxBronzeArray(idx300) = PackMixMaxBronzeArray(idx300) + 1
        recRes = 300 + PackMix(PackTotal - 300, PackMixMaxBronzeArray)
    ElseIf (PackTotal >= 1000) Then
        PackMixMaxBronzeArray(idxBronze) = PackMixMaxBronzeArray(idxBronze) + 1
        recRes = 1000 + PackMix(PackTotal - 1000, PackMixMaxBronzeArray)
    End If
    
    If (recRes > 0 And recRes < PackTotal) Then
        mult300Res = Int(PackTotal / 300)
        PackMixMaxBronzeArray(idx300) = mult300Res
        mult300Res = mult300Res * 300
        If mult300Res > recRes Then
            recRes = mult300Res
            PackMixMaxBronzeArray(idxBronze) = PackMixMaxBronzeArray(idxBronze) - Int(mult300Res / 1000)
        End If
    End If
    
    PackMix = recRes
End Function

Function WritePackMixMaxBronzeStr(ByRef PackMixMaxBronzeArray() As Long) As String
    Dim i As Integer
    Dim resStr As String
    
    For i = 0 To UBound(PackMixMaxBronzeArray) - 1
        resStr = resStr & PackMixMaxBronzeArray(i) & ", "
    Next i
    
    WritePackMixMaxBronzeStr = resStr
End Function

Function WritePackMixStr(ByRef PackMixArray() As Variant) As String
    Dim i As Integer
    Dim resStr As String
    
    For i = idx300 To idxPlatinum
        resStr = resStr & "Pack " & PackMixArray(1, i) & ": " & PackMixArray(2, i) & ". "
    Next i
    
    WritePackMixStr = resStr
End Function

''''''''''''''''''''''''''''''''''
'Fonction récursive qui calcule le montant maximal pouvant être investi en packs publicitaires$
'en fonction du montant n passé en entrée
'
'Paramètres
'
'n      montant disponible pourv être investi en packs
'm      montant maximum à investir en packs. Paramètre optionnel. Aucun effet si pas spécifié.
''''''''''''''''''''''''''''''''''
Function MaxInvest(ByVal n As Long, Optional ByVal m As Variant = 100000000) As Long
    Dim recRes As Long       'résultat du calcul par récursion
    Dim mult300Res As Long   'résultat du calcul par division par 300, pour les cas comme 2100, 2400, 2700, etc
    
    If (m <= 0) Then
        MaxInvest = 0
    End If
    
    If (m = vbNullString) Then  'si la cellule qui fournit le second parm est vide !
        m = 100000000
    End If
    
    If (n < 300) Then
        recRes = 0
    ElseIf (n >= 300 And n < 1000) Then
        recRes = 300 + MaxInvest(n - 300)
    ElseIf (n >= 1000) Then
        recRes = 1000 + MaxInvest(n - 1000)
'Le code commenté posait problème: MaxInvest(10100) renvoyait 10000 et non 10100, alors que 10100 peuvent être
'investis en 2 * 4000 + 7 * 300 !
'    ElseIf (n >= 2000) Then
'        recRes = 2000 + MaxInvest(n - 2000)
'    ElseIf (n >= 4000 And n < 10000) Then
'        recRes = 4000 + MaxInvest(n - 4000)
'    ElseIf (n >= 10000) Then
'        recRes = 10000 + MaxInvest(n - 10000)
    End If
    
    If (recRes > 0 And recRes < n) Then
        mult300Res = Int(n / 300)
        mult300Res = mult300Res * 300
        If mult300Res > recRes Then
            recRes = mult300Res
        End If
    End If
    
'Prise en compte du montant max à investir
    If (recRes > m) Then
        If (m < 300) Then
            MaxInvest = 0
        Else
            MaxInvest = m
        End If
    Else
        MaxInvest = recRes
    End If
End Function
