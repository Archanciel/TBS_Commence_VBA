Attribute VB_Name = "VirementsCommenceExport"
Option Explicit

Private Const PAIEMENT_TYPE_ACHAT_PACK As String = "Achat pack"
Private Const PAIEMENT_TYPE_MEMBERSHIP_SE As String = "Cotisation SE"
Private Const PAIEMENT_TYPE_MEMBERSHIP_PREMIUM As String = "Cotisation Premium"

Private Const TYPE_VIREMENT_TEMPORAIRE_DE_PSEUDO As String = "Transtemp de pseudo"    'Transfert de fonds temporaire sur notre BO
Private Const TYPE_VIREMENT_TEMPORAIRE_A_PSEUDO As String = "Transtemp à pseudo"      'Transfert de fonds temporaire sur le BO du pseudo
Private Const TYPE_VIREMENT_CONTRE_CASH_DE_PSEUDO As String = "Transf contre cash de pseudo"    'Transfert de fonds du pseudo contre mon virement sur le compte bancaire de celui-ci
Private Const TYPE_VIREMENT_CONTRE_CASH_A_PSEUDO As String = "Transf contre cash à pseudo"      'Transfert de fonds au pseudo contre un virement de celui-ci sur notre compte bancaire
Private Const TYPE_VIREMENT_REMBOURSEMENT_A_PSEUDO As String = "Remboursement apport" 'Transfert de fonds sur le BO du pseudo pour rembourser l'apport avancé par celui-ci
Private Const TYPE_VIREMENT_APPORT As String = "Apport"
Private Const TYPE_VIREMENT_PROMO As String = "Promo"
Private Const TYPE_VIREMENT_AUTRE As String = "Autre"
Private Const TYPE_VIREMENT_VIREMENT_SUR_BO As String = "Virement sur BO"

'Formate et traite les données issues des copy/paste des listes de virements en vue de leur
'importation dans Commence
Sub handleVirements()
    Dim rngLibelle As Range
    Dim cell As Range
    Dim curRow As Long
    Dim uidVirementCol As Long
    Dim typeVirementCol As Long
    Dim lastCellRow As Long
    Dim virementSheetCalculatedCellsRange As Range
    Dim operationTag As String
    Dim montantOpCol As Long
    Dim montantOp As Double
    Dim rngPseudoLibelle As Range
    Dim pseudoContrepartie As String
    Dim pseudoContrepartieCol As Long
    Dim compteContrepartieCol As Long
    Dim lookupTablesSheet As Worksheet
    Dim lookupRangePseudoContrat As Range
    Dim opDateCol As Long
    Dim opTimeCol As Long
    
    Application.ScreenUpdating = False
    
    terminateIfNoData Cells(2, Range("LIBELLE_VIREMENT").Column)
    
    formatDateAndTime "DATE_VIREMENT", "TIME_VIREMENT"
    transformMontant "MONTANT_VIREMENT"
    
    lastCellRow = getLastDataRow(ActiveSheet.Range("A:A"))
    Set rngLibelle = getDataRangeFromColRange(ActiveSheet.Range("LIBELLE_VIREMENT"))
    Set rngPseudoLibelle = getDataRangeFromColRange(ActiveSheet.Range("PSEUDO_VIREMENT_LIBELLE"))
    pseudoContrepartieCol = Range("PSEUDO_VIREMENT").Column
    
    uidVirementCol = Range("UID_VIREMENT").Column
    typeVirementCol = Range("TYPE_VIREMENT").Column
    montantOpCol = Range("MONTANT_VIREMENT").Column
    pseudoContrepartieCol = Range("PSEUDO_VIREMENT").Column
    compteContrepartieCol = Range("COMPTE_COUNTERPART_VIREMENT").Column
    opDateCol = Range("DATE_VIREMENT").Column
    opTimeCol = Range("TIME_VIREMENT").Column
    
    'clear col 8 à 10 qui contiennent les valeurs extraites par la suite de la macro
    Set virementSheetCalculatedCellsRange = ActiveSheet.Range(ActiveSheet.Cells(2, typeVirementCol), ActiveSheet.Cells(lastCellRow, uidVirementCol))
    virementSheetCalculatedCellsRange.Clear
    
    'au début, je n'utilisais pas systématiquement le tag #TRANSTEMP !!
    replaceInRange rngLibelle, "Retransfers", "#TRANSTEMP", False
    
    'initialize lookup table references
    Set lookupTablesSheet = Sheets("Lookup tables")
    
    lastCellRow = getLastDataRow(lookupTablesSheet.Range("E:E"))
    Set lookupRangePseudoContrat = lookupTablesSheet.Range(lookupTablesSheet.Cells(2, 6), lookupTablesSheet.Cells(lastCellRow, 7))
    
    'extraction du pseudo de la contrepartie et formatage du compte contrepartie
    For Each cell In rngPseudoLibelle
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        pseudoContrepartie = extractPseudoContrepartieFromPseudoLibelle(cell)
        
        If pseudoContrepartie <> "" Then
            Cells(curRow, pseudoContrepartieCol).Value = pseudoContrepartie
            Cells(curRow, compteContrepartieCol).Value = getCompteForPseudo(pseudoContrepartie, lookupRangePseudoContrat)
        End If
    Next
    
    'Règles de gestion:
    '
    'pour chaque cellule de la colonne LIBELLE_VIREMENT,
    '   si le libellé contient #TRANSTEMP
    '       si le MONTANT_OPERATION est positif
    '           type virement = TYPE_VIREMENT_TEMPORAIRE_DE_PSEUDO
    '       si le MONTANT_OPERATION est négatif
    '           type virement = TYPE_VIREMENT_TEMPORAIRE_A_PSEUDO
    '       end if
    '   sinon si le libellé contient #TRANSCASH
    '       si le MONTANT_OPERATION est positif
    '           type virement = TYPE_VIREMENT_CONTRE_CASH_DE_PSEUDO
    '       si le MONTANT_OPERATION est négatif
    '           type virement = TYPE_VIREMENT_CONTRE_CASH_A_PSEUDO
    '       end if
    '   traiter les autres types de virements ...
    For Each cell In rngLibelle
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        operationTag = extractTranstempFromLibelle(cell)
        
        If (operationTag <> "") Then
            'nous sommes en présence d'un transfert temporaire
            montantOp = Cells(curRow, montantOpCol).Value
            If montantOp >= 0 Then
                Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_TEMPORAIRE_DE_PSEUDO
            Else
                Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_TEMPORAIRE_A_PSEUDO
            End If
            'handle no contrat (inclu contrepartie
        Else
            operationTag = extractTranscashFromLibelle(cell)
            If (operationTag <> "") Then
                'nous sommes en présence d'un transfert de cash
                montantOp = Cells(curRow, montantOpCol).Value
                If montantOp >= 0 Then
                    Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_CONTRE_CASH_DE_PSEUDO
                Else
                    Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_CONTRE_CASH_A_PSEUDO
                End If
                'handle no contrat (inclu contrepartie
            Else
                If (isStringInLibelle(cell, "Apport")) Then
                    'Exemple de libellé: Apport initial
                    '                    Apport initial pour pouvoir activer le compte (lib  ajouté à posteriori !)
                    Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_APPORT
                Else
                    If (isStringInLibelle(cell, "PROMO")) Then
                        'Exemple de libellé: 5 % PROMO
                        Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_PROMO
                    Else
                        If (isStringInLibelle(cell, "Wire Transfer")) Then
                            'Exemple de libellé: Wire Transfer A 2015102900099516
                            Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_VIREMENT_SUR_BO
                        Else
                            If (isStringInLibelle(cell, "#Rembours")) Then
                                'Exemple de libellé: #Rembours partiel apport initial
                                Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_REMBOURSEMENT_A_PSEUDO
                            Else
                                Cells(curRow, typeVirementCol).Value = TYPE_VIREMENT_AUTRE
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'formatting virement UID
        Cells(curRow, uidVirementCol).Value = Cells(curRow, pseudoContrepartieCol).Value & Right(Cells(curRow, typeVirementCol).Value, 15) & Cells(curRow, opDateCol).Value & Cells(curRow, opTimeCol).Value
    Next cell
    
    clearAnySelection
    
    Application.ScreenUpdating = True
End Sub
Private Function getCompteForPseudo(pseudoContrepartie As String, lookupRangePseudoContrat As Range) As String
    Dim compteForPseudo As Variant
    
    compteForPseudo = Application.VLookup(pseudoContrepartie, lookupRangePseudoContrat, 2, False)
    If IsError(compteForPseudo) Then
        compteForPseudo = ""
    End If
    
    getCompteForPseudo = compteForPseudo
End Function
'Extrait du libellé contenu dans la Cell passé en parm le numéro de pack
'qu'il contient.
'
'Précond: le no de pack se trouve au début du libellé !
'
'Exemple de libellé: #12934088431 Paiement de dépot
Private Function extractPackIdFromLibelleDepotCell(cell As Range) As String
    extractPackIdFromLibelleDepotCell = extractItem(cell, "^#([0-9]+) Paiement de dépot")
End Function
'Extrait du libellé de pseudo contenu dans la Cell passé en parm le tag #TRANSTEMP
'
'Exemple de libellé de pseudo: Rose : rosemaman
'                              Admin
Private Function extractPseudoContrepartieFromPseudoLibelle(cell As Range) As String
    Dim arr() As String
    
    arr = Split(cell.Value, " : ")
    
    If (UBound(arr) - LBound(arr)) > 0 Then
        '" : " was found in libellé de pseudo
        extractPseudoContrepartieFromPseudoLibelle = arr(1)
    ElseIf (InStr(1, cell.Value, "admin", vbTextCompare) <= 0) Then
        extractPseudoContrepartieFromPseudoLibelle = ""
    Else
        extractPseudoContrepartieFromPseudoLibelle = "Admin"
    End If
End Function
'Extrait du libellé contenu dans la Cell passé en parm le tag #TRANSTEMP
'
'Exemple de libellé: #TRANSTEMP retour partiel du prêt d 2000 $ du 04.01.2016
'                    #TRANSTEM retour partiel du prêt d 2000 $ du 04.01.2016 >>> libellé erroné #TRANSTEM au lieu de #TRANSTEMP !!!
Private Function extractTranstempFromLibelle(cell As Range) As String
    extractTranstempFromLibelle = extractItem(cell, "^(#TRANSTEM[ P]+)") 'traite le libellé erroné
End Function
'Extrait du libellé contenu dans la Cell passé en parm le tag #TRANSTEMP
'
'Exemple de libellé: #TRANSCASH pour demande Antoine pr mail le 8.2.2016 de 3000 $
Private Function extractTranscashFromLibelle(cell As Range) As String
    extractTranscashFromLibelle = extractItem(cell, "^(#TRANSCASH+)")
End Function
'Renvoie TRUE si le libellé contenu dans la Cell passé en parm contient le mot
'passé en parm (case insensitive)
Private Function isStringInLibelle(cell As Range, str As String) As Boolean
    Dim index As Integer
    
    index = InStr(1, cell.Value, str, vbBinaryCompare)   'case sensitive
    
    If (index <= 0) Then
        isStringInLibelle = False
    Else
        isStringInLibelle = True
    End If
End Function
'Extrait du libellé contenu dans la Cell passé en parm le numéro de pack
'qu'il contient.
'
'Précond: le no de pack se trouve au début du libellé !
'
'Exemple de libellé: #129340130641 SVIP payment
Private Function extractPackIdFromLibelleSEMembershipCell(cell As Range) As String
    extractPackIdFromLibelleSEMembershipCell = extractItem(cell, "^#([0-9]{12}) SVIP payment")
End Function
'Extrait du libellé contenu dans la Cell passé en parm le numéro de pack
'qu'il contient.
'
'Précond: le no de pack se trouve au début du libellé !
'
'Exemple de libellé: #129340104707 Membership payment ou #12934081869 Règlement adhésion
Private Function extractPackIdFromLibellePremiumMembershipCell(cell As Range) As String
    extractPackIdFromLibellePremiumMembershipCell = extractItem(cell, "^#([0-9]*) (Membership payment|Règlement adhésion)")
End Function
'Exporte les données de la feuille Transferts-virements dans un fichier texte tab separated pouvant être importé dans Commence
Sub virementsExportDataForCommence()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    deleteNomComptes "NOM_COMPTES_V"
    deleteTopRow
    saveSheetAsTabDelimTxtFileTimeStamped ActiveSheet.Name
    Application.ScreenUpdating = True
    closeWithoutSave
End Sub


