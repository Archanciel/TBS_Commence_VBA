Attribute VB_Name = "PacksGainsCommenceExport"
Option Explicit
Private Const RENDEMENT_REGULAR_PACK As String = "25"
Private Const RENDEMENT_XMAS_PACK As String = "28"

Private Const GAIN_TYPE_BONUS_ACHAT_PACK_PAR_FILLEUL As String = "Bonus achat pack par filleul"
Private Const GAIN_TYPE_GAIN_PACK_25_PCT As String = "Gain pack 25 %"
Private Const GAIN_TYPE_GAIN_PACK_28_PCT As String = "Gain pack 28 %"
Private Const GAIN_TYPE_GAIN_PACK_UNKNOWN As String = "### Gain pack inconnu ###"
Private Const GAIN_TYPE_BONUS_FILLEUL_MATRICE_PREMIUM As String = "Bonus matrice Premium"
Private Const GAIN_TYPE_BONUS_FILLEUL_MATRICE_SE As String = "Bonus matrice SE"
Private Const GAIN_TYPE_BONUS_FILLEUL_UPGR_PREMIUM As String = "Bonus filleul upgr Premium"
Private Const GAIN_TYPE_BONUS_FILLEUL_UPGR_SE As String = "Bonus filleul upgr SE"
Private Const GAIN_IMPORT_FLAG_TRUE As String = "1"
Private Const GAIN_VERIFIED_FLAG_TRUE As String = "1"

Sub clearSheet()
    Dim delRange As Range
    Dim topLeftTitleCell As Range
    Dim topLeftDataCell As Range
    Dim ans As Long
    
    Set topLeftTitleCell = Range("A1")
    Set topLeftDataCell = Range("A2")
    Set delRange = Range(topLeftDataCell, Cells(topLeftDataCell.End(xlDown).Row, topLeftTitleCell.End(xlToRight).Column))
    
    delRange.Select
    ans = MsgBox("Supprimer la s�lection ?", vbYesNo + vbExclamation)
    
    If ans <> vbYes Then
        clearAnySelection
        Exit Sub
    Else
        delRange.Clear
    End If
End Sub

'Formate les donn�es issues des copy/paste des listes de packs en vue de leur importation
'dans Commence
Sub packsFormatAndSortData()
    Application.ScreenUpdating = False
    formatIdCol ("NOM_PACK")
    formatDateAndTime "DATE_ACHAT", "TIME_ACHAT_PACK"
    'adapte col width for id pack
    Columns("D:D").EntireColumn.AutoFit
    formatRendement 'doit �tre appel� avant transformType !!
    transformType
    transformMontant "MONTANT_PACK"
    transformMontantGain "GAIN_TOTAL"
    replaceEnCoursByZeroEchuByOne
    writeNomComptes
    buildLookupTables
    Sheets("Packs").Select
    clearAnySelection
    Application.ScreenUpdating = True
End Sub
Private Sub formatRendement()
    Dim packTypeRange As Range
    Dim cell As Range
    Dim packTypeStr As String
    Dim rendementCol As Long
    Dim curRow As Long
    
    Set packTypeRange = getDataRangeFromColRange(ActiveSheet.Range("TYPE"))
    
    rendementCol = ActiveSheet.Range("RENDEMENT_PACK").Column
    
    For Each cell In packTypeRange
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        packTypeStr = cell.Value
        
        If (InStr(1, packTypeStr, "xmas", vbTextCompare) > 0) Then
            'XMAS pack avec rendement de 28 %
            Cells(curRow, rendementCol).Value = RENDEMENT_XMAS_PACK
        Else
            If IsEmpty(Cells(curRow, rendementCol).Value) Then
                'il arrive en effet que la feuille pack a �t� process�e auparavant et que l'on rajoute une ligne
                'pour un nouveau pack. Si la feuille contenait des lignes pour, par ex, un xmas pack 1000, le type
                'de pack ayant �t� transform� en Bronze, sans cette garde, le rendement de 28 % serait �cras� en 25 % !!
                '
                'I know, this stinks, but I chosed not to create additional pack types !
                Cells(curRow, rendementCol).Value = RENDEMENT_REGULAR_PACK
            End If
        End If
    Next cell
End Sub
'Exporte les donn�es de la feuille Packs dans un fichier texte tab separated pouvant �tre import� dans Commence
Sub packsExportDataForCommence()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    deleteNomComptes "NOM_COMPTES"
    deleteTopRow
    saveSheetAsTabDelimTxtFileTimeStamped ActiveSheet.Name
    Application.ScreenUpdating = True
    closeWithoutSave
End Sub

'Exporte les donn�es de la feuille Gains dans un fichier texte tab separated pouvant �tre import� dans Commence
Sub gainsExportDataForCommence()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    deleteNomComptes "NOM_COMPTES_G"
    deleteTopRow
    saveSheetAsTabDelimTxtFileTimeStamped ActiveSheet.Name
    Application.ScreenUpdating = True
    closeWithoutSave
End Sub

'Formate et traite les donn�es issues des copy/paste des listes de gains en vue de leur
'importation dans Commence
Sub handleRevenues()
    Dim rngLibelle As Range
    Dim cell As Range
    Dim packId As String
    Dim tauxGain As Integer
    Dim gainPackMonth As String
    Dim pseudoFilleul As String
    Dim matriceLevel As String
    Dim compteReceivingGainCol As Long
    Dim idGainCol As Long
    Dim matriceLevelCol As Long
    Dim curRow As Long
    Dim packIdCol As Long
    Dim typeGainCol As Long
    Dim pseudoFilleulCol As Long
    Dim dateGainCol As Long
    Dim lastCellRow As Long
    Dim lookupTablesSheet As Worksheet
    Dim lookupRangePackContrat As Range
    Dim lookupRangeContratPseudo As Range
    Dim gainSheetCalculatedCellsRange As Range
    Dim noGainCol As Long
    Dim nomCheckBOTBSForGainCol As Long
    Dim montantGainCol As Long
    Dim windowsWideThousandSeparator As String
    Dim importFlagCol As Long
    Dim verifiedFlagCol As Long
    
    Application.ScreenUpdating = False
    windowsWideThousandSeparator = Application.International(xlThousandsSeparator)
    
    formatDateAndTime "DATE_GAIN_COL", "TIME_GAIN"
    transformMontant "MONTANT_GAIN_COL"
    
    Set rngLibelle = Range("LIBELLE")
    
    compteReceivingGainCol = Range("COMPTE_RECEIVING_GAIN").Column
    packIdCol = Range("PACK_ID").Column
    typeGainCol = Range("TYPE_GAIN").Column
    idGainCol = Range("ID_GAIN").Column
    matriceLevelCol = Range("MATRICE_LEVEL").Column
    pseudoFilleulCol = Range("PSEUDO_FILLEUL").Column
    dateGainCol = Range("DATE_GAIN_COL").Column
    noGainCol = Range("NO_GAIN").Column
    nomCheckBOTBSForGainCol = Range("NOM_ID_CHECK_BO_TBS_FOR_GAIN").Column
    montantGainCol = Range("MONTANT_GAIN").Column
    importFlagCol = Range("GAIN_IMPORT").Column
    verifiedFlagCol = Range("GAIN_VERIFIED").Column
    
    Set lookupTablesSheet = Sheets("Lookup tables")
    
    lastCellRow = getLastDataRow(lookupTablesSheet.Range("A:A"))
    Set lookupRangePackContrat = lookupTablesSheet.Range(lookupTablesSheet.Cells(2, 1), lookupTablesSheet.Cells(lastCellRow, 2))
    
    lastCellRow = getLastDataRow(lookupTablesSheet.Range("D:D"))
    Set lookupRangeContratPseudo = lookupTablesSheet.Range(lookupTablesSheet.Cells(2, 5), lookupTablesSheet.Cells(lastCellRow, 6))

    'clear col 6 � 10 qui contiennent les valeurs extraites par la suite de la macro
    lastCellRow = getLastDataRow(ActiveSheet.Range("A:A"))
    Set gainSheetCalculatedCellsRange = ActiveSheet.Range(ActiveSheet.Cells(2, 6), ActiveSheet.Cells(lastCellRow, 10))
    gainSheetCalculatedCellsRange.Clear
    
    For Each cell In rngLibelle
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        Cells(curRow, importFlagCol).Value = GAIN_IMPORT_FLAG_TRUE
        Cells(curRow, verifiedFlagCol).Value = GAIN_VERIFIED_FLAG_TRUE
        packId = extractPackIdFromBonusLibelle(cell)
        
        If (packId <> "") Then
            'gain de type 8 % sur achat de pack par un filleul du d�tenteur du compte
            Cells(curRow, packIdCol).Value = packId
            Cells(curRow, idGainCol).Value = packId & "-b"
            Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_ACHAT_PACK_PAR_FILLEUL
            formatPseudoFilleulForPackId packId, curRow, pseudoFilleulCol, lookupRangePackContrat, lookupRangeContratPseudo
        Else
            packId = extractPackIdFromGainPackLibelle(cell)
            If (packId <> "") Then
                'revenu de pack de 25 ou 28 %
                tauxGain = extractTauxGainFromGainPackLibelle(cell)
                If (tauxGain = 25) Then
                    Cells(curRow, typeGainCol).Value = GAIN_TYPE_GAIN_PACK_25_PCT
                ElseIf (tauxGain = 28) Then
                    Cells(curRow, typeGainCol).Value = GAIN_TYPE_GAIN_PACK_28_PCT
                Else
                    Cells(curRow, typeGainCol).Value = GAIN_TYPE_GAIN_PACK_UNKNOWN
                End If
                gainPackMonth = extractPackMonthFromGainPackLibelle(cell)
                Cells(curRow, packIdCol).Value = packId
                Cells(curRow, idGainCol).Value = packId & "-" & gainPackMonth
                Cells(curRow, noGainCol).Value = gainPackMonth
                Cells(curRow, nomCheckBOTBSForGainCol).Value = buildNomLinkedCheckBOTBS(curRow, Cells(curRow, compteReceivingGainCol).Value, packId, montantGainCol, windowsWideThousandSeparator, gainPackMonth)
            Else
                pseudoFilleul = extractPseudoFilleulMatrixPrem(cell)
                If (pseudoFilleul <> "") Then
                    'bonus mensuel comptabilis� dans la matrice Premium
                    Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                    Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_FILLEUL_MATRICE_PREMIUM
                    Cells(curRow, idGainCol).Value = pseudoFilleul & "-BMP-to-" & Cells(curRow, compteReceivingGainCol).Value & "-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                    matriceLevel = extractMatriceLevelMatrixPrem(cell)
                    Cells(curRow, matriceLevelCol).Value = matriceLevel
                Else
                    pseudoFilleul = extractPseudoFilleulMatrixSE(cell)
                    If (pseudoFilleul <> "") Then
                        'bonus mensuel comptabilis� dans la matrice Super Elite
                        Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                        Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_FILLEUL_MATRICE_SE
                        Cells(curRow, idGainCol).Value = pseudoFilleul & "-BSE-to-" & Cells(curRow, compteReceivingGainCol).Value & "-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                        matriceLevel = extractMatriceLevelMatrixSE(cell)
                        Cells(curRow, matriceLevelCol).Value = matriceLevel
                    Else
                        pseudoFilleul = extractFilleulUpgrToPremium(cell)
                        If (pseudoFilleul <> "") Then
                            'bonus provenant de l'activation ou de l'upgrade en Premium d'un filleul du d�tenteur du compte
                            Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                            Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_FILLEUL_UPGR_PREMIUM
                            Cells(curRow, idGainCol).Value = pseudoFilleul & "-UPGR_PREM-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                        Else
                            pseudoFilleul = extractFilleulUpgrToSE(cell)
                            If (pseudoFilleul <> "") Then
                                'bonus provenant de l'upgrade en Super Elite d'un filleul du d�tenteur du compte
                                Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                                Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_FILLEUL_UPGR_SE
                                Cells(curRow, idGainCol).Value = pseudoFilleul & "-UPGR_SE-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                            Else
                                Cells(curRow, typeGainCol).Value = "### LIBELLE DE GAIN INCONNU ###"
                                MsgBox "Libell� de gain inconnu dans cellule " & cell.Address
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next cell
    
    clearAnySelection
    
    Application.ScreenUpdating = True
End Sub

'Construit le nom (qui a fonction d'identifiant) de la fiche Todo TBS qui doit �tre li�e au gain
'
'Exemple de nom: Gain Compte TBS Antoine 17608086151 75 1/12
'                Gain Compte TBS JPS 12934054302 1 000 3/12
'
'A noter l'utilisation du s�parateur de millier tel que d�finit dans Windows !
Private Function buildNomLinkedCheckBOTBS(curRow As Long, compteReceivingGainStr As String, packId As String, montantGainCol As Long, windowsWideThousandSeparator As String, gainPackMonth As String) As String
    Dim nomLinkedTodoTBS As String
    Dim formatedMontantGain As String
    
    formatedMontantGain = Format(Cells(curRow, montantGainCol).Value, "#" & windowsWideThousandSeparator & "##0")
    nomLinkedTodoTBS = "Gain " & compteReceivingGainStr & " " & packId & " " & formatedMontantGain & " " & gainPackMonth & "/12"
    buildNomLinkedCheckBOTBS = nomLinkedTodoTBS
End Function
Private Sub formatPseudoFilleulForPackId(packId As String, curRow As Long, pseudoFilleulCol As Long, lookupRangePackContrat As Range, lookupRangeContratPseudo As Range)
    Dim nomContratCommence As Variant
    Dim pseudoTBS As Variant
    Dim pseudoFilleul As String
    
    nomContratCommence = Application.VLookup(CDbl(packId), lookupRangePackContrat, 2, False)
    
    If (IsError(nomContratCommence)) Then
        pseudoTBS = "### pack id '" + packId + "' not found in lookup table ###"
    Else
        pseudoTBS = Application.VLookup(nomContratCommence, lookupRangeContratPseudo, 2, False)
        If (IsError(pseudoTBS)) Then
            pseudoTBS = "### contrat Commence '" + nomContratCommence + "' not found in lookup table ###"
        End If
    End If
    
    Cells(curRow, pseudoFilleulCol).Value = pseudoTBS
End Sub
'Extrait du libell� d'annonce de gain le num�ro de pack dont l'achat par un filleul
'a g�n�r� le bonus.
'
'Exemple de libell�: Bonus sponsor pour d�pot(#13441058360)
Private Function extractPackIdFromBonusLibelle(cell As Range) As String
    extractPackIdFromBonusLibelle = extractItem(cell, "d�pot\(#([0-9]+)\)$")
End Function

'Extrait du libell� d'annonce de gain de pack le num�ro de pack  qui a g�n�r� le gain.
'
'Exemple de libell�: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractPackIdFromGainPackLibelle(cell As Range) As String
    extractPackIdFromGainPackLibelle = extractItem(cell, "^#([0-9]+)")
End Function

'Extrait du libell� d'annonce de gain de pack le num�ro du mois du gain.
'
'Exemple de libell�: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractPackMonthFromGainPackLibelle(cell As Range) As String
    extractPackMonthFromGainPackLibelle = extractItem(cell, "([0-9]+)/12\]$")
End Function

'Extrait du libell� d'annonce de gain de pack le num�ro du mois du gain.
'
'Exemple de libell�: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractTauxGainFromGainPackLibelle(cell As Range) As String
    extractTauxGainFromGainPackLibelle = extractItem(cell, "Profit, ([0-9]+)\.")
End Function

'Extrait du libell� d'annonce de bonus matrice Premium le pseudo du filleul.
'
'Exemple de libell�: Niveau r�seau Premium#1 bonus (tamcerise)  ou
'                    VIP Network level#1 bonus (lucky70)
Private Function extractPseudoFilleulMatrixPrem(cell As Range) As String
    Dim strPseudo As String
    
    strPseudo = extractItem(cell, "^Niveau r�seau Premium#\d* bonus \(([a-zA-Z0-9-_]+)\)")
    
    If (strPseudo = "") Then
        'essai avec la version anglaise du libell�
        strPseudo = extractItem(cell, "^VIP Network level#\d* bonus \(([a-zA-Z0-9-_]+)\)")
    End If
    
    extractPseudoFilleulMatrixPrem = strPseudo
End Function

'Extrait du libell� d'annonce de bonus matrice Super Elite le pseudo du filleul.
'
'Exemple de libell�: SVIP level#1 bonus (jpensuisse)
Private Function extractPseudoFilleulMatrixSE(cell As Range) As String
    Dim strPseudo As String
    
    strPseudo = extractItem(cell, "^SVIP level#\d* bonus \(([a-zA-Z0-9-_]+)\)")
    
'    If (strPseudo = "") Then
'        'essai avec la version fran�aise du libell�
'        strPseudo = extractItem(cell, "^VIP Network level#\d* bonus \(([a-zA-Z0-9-_]+)\)")
'    End If

    extractPseudoFilleulMatrixSE = strPseudo
End Function

'Extrait du libell� d'annonce de bonus matrice premium le niveau matriciel du gain.
'
'Exemple de libell�: Niveau r�seau Premium#1 bonus (tamcerise)
'                    VIP Network level#1 bonus (lucky70)
Private Function extractMatriceLevelMatrixPrem(cell As Range) As String
    Dim strLevel As String
    
    strLevel = extractItem(cell, "^Niveau r�seau Premium#(\d*) bonus")
    
    If (strLevel = "") Then
        'essai avec la version anglaise du libell�
        strLevel = extractItem(cell, "^VIP Network level#(\d*) bonus")
    End If
    
    extractMatriceLevelMatrixPrem = strLevel
End Function

'Extrait du libell� d'annonce de bonus matrice premium le niveau matriciel du gain.
'
'WARNING: TU NE CONNAIS PAS ENCORE AVEC CERTITUDE LES LIBELLES EXACTS !
'Exemple de libell�: Niveau r�seau Super Elite#1 bonus (tamcerise)
'                    SVIP level#1 bonus (lucky70)
Private Function extractMatriceLevelMatrixSE(cell As Range) As String
    Dim strLevel As String
    
    strLevel = extractItem(cell, "^Niveau r�seau Super Elite#(\d*) bonus")
    
    If (strLevel = "") Then
        'essai avec la version anglaise du libell�
        strLevel = extractItem(cell, "^SVIP level#(\d*) bonus")
    End If
    
    extractMatriceLevelMatrixSE = strLevel
End Function

'Extrait le pseudo du filleul du libell� d'annonce de bonus en cas d'upgrade de celui-ci � Premium.
'
'Exemple de libell�: Bonus sponsor (rosemaman)
Private Function extractFilleulUpgrToPremium(cell As Range) As String
    extractFilleulUpgrToPremium = extractItem(cell, "^Bonus sponsor \(([a-zA-Z0-9-_]+)\)")
End Function

'Extrait le pseudo du filleul du libell� d'annonce de bonus en cas d'upgrade de celui-ci � Super Elite.
'
'Exemple de libell�: SVIP Sponsor bonus (jpensuisse)
Private Function extractFilleulUpgrToSE(cell As Range) As String
    extractFilleulUpgrToSE = extractItem(cell, "^SVIP Sponsor bonus \(([a-zA-Z0-9-_]+)\)")
End Function

Private Sub formatDate(colName As String)
Attribute formatDate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formatDate Macro
'

'
    ActiveSheet.Range(colName).Select
    Selection.NumberFormat = "d/m/yyyy"
End Sub
Private Sub transformType()
Attribute transformType.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.Range("TYPE").Select
    
    'handling xmas pack denomination
    Selection.replace What:="Xmas pack 1000", Replacement:="Bronze", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.replace What:="Xmas pack 2000", Replacement:="Silver", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.replace What:="Xmas pack 4000", Replacement:="Gold", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.replace What:="Xmas pack 10000", Replacement:="Platinum", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    'handling regular pack denomination
    Selection.replace What:=" USD", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.replace What:="($1000)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.replace What:="($2000)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.replace What:="($4000)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.replace What:="($10000)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "@"
End Sub
Private Sub transformMontantGain(colName As String)
    ActiveSheet.Range(colName).Select
    Selection.replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Private Sub replaceEnCoursByZeroEchuByOne()
Attribute replaceEnCoursByZeroEchuByOne.VB_ProcData.VB_Invoke_Func = " \n14"
'
' replaceEnCoursByZeroEchuByOne Macro
'

'
    ActiveSheet.Range("ECHU").Select
    Selection.replace What:="En cours", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
'Recr�e la  zone NOM_COMPTES qui contient les noms de contrats TBS dans Commence.
'Ces noms sont utilis�s en copy/paste lors de l'entr�e de nouvelles donn�es dans
'la feuille Packs
Private Sub writeNomComptes()
    ActiveSheet.Range("NOM_COMPTES").Select
    ActiveCell.FormulaR1C1 = "Compte TBS Antoine"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "Compte TBS B�atrice"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "Compte TBS JPS"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "Compte TBS Maman"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "Compte TBS Papa"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "Compte TBS Tamara"
End Sub

'Sub appel�e par packsFormatAndSortData()
Private Sub buildLookupTables()
    Dim lookupTablesSheet As Worksheet
    Dim packsSheet As Worksheet
    Dim lookupRangePackContrat As Range
    Dim lookupTableLastCellRowPlusOne As Long
    Dim packsSheetLastCellRow As Long
    Dim lookupTableLastCellRow As Long
    
    Set lookupTablesSheet = Sheets("Lookup tables")
    Set packsSheet = Sheets("Packs")
    
    lookupTableLastCellRowPlusOne = getLastDataRow(lookupTablesSheet.Range("A:A")) + 1
    
    If (lookupTableLastCellRowPlusOne > 1000000) Then
        'le cas si la lookup table ne contient aucune entr�e pack/compte/date pack !
        lookupTableLastCellRowPlusOne = 2
    End If
    
    packsSheetLastCellRow = getLastDataRow(packsSheet.Range("A:A"))
    
    'copie la colonne no de packs
    packsSheet.Select
    packsSheet.Range(packsSheet.Cells(2, 4), packsSheet.Cells(packsSheetLastCellRow, 4)).Select
    Selection.Copy
    lookupTablesSheet.Select
    lookupTablesSheet.Cells(lookupTableLastCellRowPlusOne, 1).Select
    ActiveSheet.Paste
    
    'copie la colonne nom de contrat
    packsSheet.Select
    packsSheet.Range(packsSheet.Cells(2, 1), packsSheet.Cells(packsSheetLastCellRow, 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    lookupTablesSheet.Select
    lookupTablesSheet.Cells(lookupTableLastCellRowPlusOne, 2).Select
    ActiveSheet.Paste
    
    'copie la colonne date achat (utile pour purger les packs plus vieux d'une ann�e de la lookup table !)
    packsSheet.Select
    packsSheet.Range(packsSheet.Cells(2, 3), packsSheet.Cells(packsSheetLastCellRow, 3)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    lookupTablesSheet.Select
    lookupTablesSheet.Cells(lookupTableLastCellRowPlusOne, 3).Select
    ActiveSheet.Paste
    
    'purge any duplicate pack id line
    lookupTableLastCellRow = getLastDataRow(lookupTablesSheet.Range("A:A"))
    Range("A1", Cells(lookupTableLastCellRow, 3)).RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    'adapte col width
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    
    ActiveSheet.Range("A1").Select
End Sub
