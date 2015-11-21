Attribute VB_Name = "PacksGainsCommenceExport"
Option Explicit
Private Const BONUS_ACHAT_PACK_PAR_FILLEUL As String = "Bonus achat pack par filleul"
Private Const GAIN_PACK_25_PCT As String = "Gain pack 25 %"
Private Const BONUS_FILLEUL_MATRICE_PREMIUM = "Bonus matrice Premium"
Private Const BONUS_FILLEUL_MATRICE_SE = "Bonus matrice SE"
Private Const BONUS_FILLEUL_UPGR_PREMIUM = "Bonus filleul ugr Premium"
Private Const BONUS_FILLEUL_UPGR_SE = "Bonus filleul ugr SE"

'Formate les donn�es issues des copy/paste des listes de packs en vue de leur importation
'dans Commence
Sub packsFormatAndSortData()
    Application.ScreenUpdating = False
    formatDate "DATE_ACHAT"
    transformType
    transformMontant "MONTANT_PACK"
    transformMontantGain "GAIN_TOTAL"
    replaceEnCoursByZeroEchuByOne
    setDateUpdateToToday
    triPourDefinitionRang
    writeNomComptes
    buildLookupTables
    Sheets("Packs").Select
    clearAnySelection
    Application.ScreenUpdating = True
End Sub

'Exporte les donn�es de la feuille Packs dans un fichier texte tab separated pouvant �tre import� dans Commence
Sub packsExportDataForCommence()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    deleteNomComptes
    deleteTopRow
    saveSheetAsTabDelimTxtFile "Packs", "Packs JPS et filleuls Commence export.txt"
    Application.ScreenUpdating = True
    closeWithoutSave
End Sub

Private Sub closeWithoutSave()
    MsgBox "La version modifi�e (sans ligne de titres) de la spreadsheet va �tre ferm�e sans �tre sauv�e. Veuillez rouvrir la version .xlsm (sauv�e avant l'exportation) !", vbInformation
    ActiveWorkbook.Close savechanges:=False
End Sub

'Exporte les donn�es de la feuille Gains dans un fichier texte tab separated pouvant �tre import� dans Commence
Sub gainsExportDataForCommence()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    deleteTopRow
    saveSheetAsTabDelimTxtFile "Gains", "Gains JPS et filleuls Commence export.txt"
    Application.ScreenUpdating = True
    closeWithoutSave
End Sub

'Formate et traite les donn�es issues des copy/paste des listes de gains en vue de leur
'importation dans Commence
Sub handleRevenues()
    Dim rngLibelle As Range
    Dim cell As Range
    Dim packId As String
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
    
    Application.ScreenUpdating = False
    
    formatDate "DATE_GAIN_COL"
    transformMontant "MONTANT_GAIN_COL"
    
    Set rngLibelle = Range("LIBELLE")
    
    compteReceivingGainCol = Range("COMPTE_RECEIVING_GAIN").Column
    packIdCol = Range("PACK_ID").Column
    typeGainCol = Range("TYPE_GAIN").Column
    idGainCol = Range("ID_GAIN").Column
    matriceLevelCol = Range("MATRICE_LEVEL").Column
    pseudoFilleulCol = Range("PSEUDO_FILLEUL").Column
    dateGainCol = Range("DATE_GAIN_COL").Column
    
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
        
        packId = extractPackIdFromBonusLibelle(cell)
        
        If (packId <> "") Then
            'gain de type 8 % sur achat de pack par un filleul du d�tenteur du compte
            Cells(curRow, packIdCol).Value = packId
            Cells(curRow, idGainCol).Value = packId & "-b"
            Cells(curRow, typeGainCol).Value = BONUS_ACHAT_PACK_PAR_FILLEUL
            formatPseudoFilleulForPackId packId, curRow, pseudoFilleulCol, lookupRangePackContrat, lookupRangeContratPseudo
        Else
            packId = extractPackIdFromGainPackLibelle(cell)
            If (packId <> "") Then
                'gain de 25 % rapport� par un packs du compte
                gainPackMonth = extractPackMonthFromGainPackLibelle(cell)
                Cells(curRow, packIdCol).Value = packId
                Cells(curRow, idGainCol).Value = packId & "-" & gainPackMonth
                Cells(curRow, typeGainCol).Value = GAIN_PACK_25_PCT
            Else
                pseudoFilleul = extractPseudoFilleulMatrixPrem(cell)
                If (pseudoFilleul <> "") Then
                    'bonus mensuel comptabilis� dans la matrice Premium
                    Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                    Cells(curRow, typeGainCol).Value = BONUS_FILLEUL_MATRICE_PREMIUM
                    Cells(curRow, idGainCol).Value = pseudoFilleul & "-BMP-to-" & Cells(curRow, compteReceivingGainCol).Value & "-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                    matriceLevel = extractMatriceLevelMatrixPrem(cell)
                    Cells(curRow, matriceLevelCol).Value = matriceLevel
                Else
                    pseudoFilleul = extractPseudoFilleulMatrixSE(cell)
                    If (pseudoFilleul <> "") Then
                        'bonus mensuel comptabilis� dans la matrice Super Elite
                        Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                        Cells(curRow, typeGainCol).Value = BONUS_FILLEUL_MATRICE_SE
                        Cells(curRow, idGainCol).Value = pseudoFilleul & "-BSE-to-" & Cells(curRow, compteReceivingGainCol).Value & "-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                        matriceLevel = extractMatriceLevelMatrixSE(cell)
                        Cells(curRow, matriceLevelCol).Value = matriceLevel
                    Else
                        pseudoFilleul = extractFilleulUpgrToPremium(cell)
                        If (pseudoFilleul <> "") Then
                            'bonus provenant de l'activation ou de l'upgrade en Premium d'un filleul du d�tenteur du compte
                            Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                            Cells(curRow, typeGainCol).Value = BONUS_FILLEUL_UPGR_PREMIUM
                            Cells(curRow, idGainCol).Value = pseudoFilleul & "-UPGR_PREM-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                        Else
                            pseudoFilleul = extractFilleulUpgrToSE(cell)
                            If (pseudoFilleul <> "") Then
                                'bonus provenant de l'upgrade en Super Elite d'un filleul du d�tenteur du compte
                                Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                                Cells(curRow, typeGainCol).Value = BONUS_FILLEUL_UPGR_SE
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

Private Sub clearAnySelection()
    Application.CutCopyMode = False
    ActiveSheet.Range("A1").Select
End Sub
Private Sub formatPseudoFilleulForPackId(packId As String, curRow As Long, pseudoFilleulCol As Long, lookupRangePackContrat As Range, lookupRangeContratPseudo As Range)
    Dim nomContratCommence As Variant
    Dim pseudoTBS As Variant
    Dim pseudoFilleul As String
    
    nomContratCommence = Application.VLookup(CDbl(packId), lookupRangePackContrat, 2, False)
    
    If (IsError(nomContratCommence)) Then
        pseudoTBS = "### pack # not found ###"
    End If
    
    pseudoTBS = Application.VLookup(nomContratCommence, lookupRangeContratPseudo, 2, False)
    
    If (IsError(pseudoTBS)) Then
        pseudoTBS = "### contrat Commence not found ###"
    End If
    
    Cells(curRow, pseudoFilleulCol).Value = pseudoTBS
End Sub

'Extrait du libell� d'annonce de bonus le num�ro de pack dont l'achat par un filleul
'a g�n�r� le bonus.
'
'Exemple de libell�: Bonus sponsor pour d�pot(#13441058360)
Private Function extractPackIdFromBonusLibelle(cell As Range) As String
    extractPackIdFromBonusLibelle = extractItem(cell, "d�pot\(#([0-9]{11})\)$")
End Function

'Extrait du libell� d'annonce de gain de pack le num�ro de pack  qui a g�n�r� le gain.
'
'Exemple de libell�: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractPackIdFromGainPackLibelle(cell As Range) As String
    extractPackIdFromGainPackLibelle = extractItem(cell, "^#([0-9]{11})")
End Function

'Extrait du libell� d'annonce de gain de pack le num�ro du mois du gain.
'
'Exemple de libell�: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractPackMonthFromGainPackLibelle(cell As Range) As String
    extractPackMonthFromGainPackLibelle = extractItem(cell, "([0-9]+)/12\]$")
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

    extractPseudoFilleulMatrixPrem = strPseudo
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
    
    extractMatriceLevelMatrixPrem = strLevel
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
Private Function extractItem(cell As Range, regexp As String) As String
    Dim regEx As New VBScript_RegExp_55.regexp
    Dim matches
    regEx.Pattern = regexp
    regEx.IgnoreCase = True 'True to ignore case
    regEx.Global = False 'True matches all occurances, False matches the first occurance
    If regEx.Test(cell.Value) Then
        Set matches = regEx.Execute(cell.Value)
        extractItem = matches(0).SubMatches(0) 'extraction du 1er groupe
    Else
        extractItem = ""
    End If
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
'
' transformType Macro
'

'
    ActiveSheet.Range("TYPE").Select
    Selection.Replace What:=" USD", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="($1000)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="($2000)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="($4000)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="($10000)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "@"
End Sub
Private Sub transformMontant(colName As String)
'
' transformMontant Macro
'

'
    ActiveSheet.Range(colName).Select
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Private Sub transformMontantGain(colName As String)
'
' transformMontant Macro
'

'
    ActiveSheet.Range(colName).Select
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
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
    Selection.Replace What:="En cours", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Private Sub setDateUpdateToToday()
Attribute setDateUpdateToToday.VB_ProcData.VB_Invoke_Func = " \n14"
'
' setDateUpdateToToday Macro
'

'
    Dim lastCellRow As Long
    Dim firstColCell As Range
    
    lastCellRow = getLastDataRow(Range("RANG"))
    ActiveSheet.Range("DATE_UPDATE").Select
    Set firstColCell = ActiveCell.Offset(1, 0)
    firstColCell.FormulaR1C1 = "=NOW()"
    Selection.NumberFormat = "d/m/yyyy"
    firstColCell.Copy
    ActiveSheet.Range(Cells(firstColCell.Row, firstColCell.Column), Cells(lastCellRow, firstColCell.Column)).Select
    ActiveSheet.Paste
End Sub
Function getLastDataRow(colCell As Range) As Long
    Dim lastCell As Range
    Dim lastCellRow As Long
    
    Set lastCell = colCell.End(xlDown)
    getLastDataRow = lastCell.Row
End Function
Private Sub triPourDefinitionRang()
Attribute triPourDefinitionRang.VB_ProcData.VB_Invoke_Func = " \n14"
'
' triPourDefinitionRang Macro
'

'
    Cells.Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("A2:A29"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("C2:C29"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("D2:D29"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:J29")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'Sauve une feuille sp�cifique dans un fichier txt tab delimited
Private Sub saveSheetAsTabDelimTxtFile(sheetName As String, fileName As String)
    Dim ans As Long
    Dim sSaveAsFilePath As String

    On Error GoTo ErrHandler:
    
    sSaveAsFilePath = "D:\Users\Jean-Pierre\OneDrive\Documents\Excel\" & fileName

    If Dir(sSaveAsFilePath) <> "" Then
        ans = MsgBox("Le fichier " & sSaveAsFilePath & " existe d�j�. Remplacer ?", vbYesNo + vbExclamation)
        If ans <> vbYes Then
            Exit Sub
        Else
            Kill sSaveAsFilePath
        End If
    End If
    
    Sheets(sheetName).Copy '//Copy sheet Packs to new workbook
    ActiveWorkbook.SaveAs sSaveAsFilePath, xlTextWindows '//Save as text (tab delimited) file
    
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then '//Double sure we don't close this workbook
        ActiveWorkbook.Close False
    End If

My_Exit:
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume My_Exit
End Sub

'Supprime la ligne contenant les en-t�tes de colonnes afin qu'elles ne soient pas export�es.
'
'Cette suppression n'affecte que la version txt de la speadsheet et non la version xlsm !
Private Sub deleteTopRow()
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
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

'Delete la  zone NOM_COMPTES qui contient les noms de contrats TBS dans Commence.
'En effet, ces donn�es ne doivent pas �tre export�es !
'
'Ces noms sont utilis�s en copy/paste lors de l'entr�e de nouvelles donn�es dans
'la feuille Packs
Private Sub deleteNomComptes()
    ActiveSheet.Range("NOM_COMPTES").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = ""
End Sub

Private Sub buildLookupTables()
    Dim lookupTablesSheet As Worksheet
    Dim packsSheet As Worksheet
    Dim lookupRangePackContrat As Range
    Dim lookupTableLastCellRowPlusOne As Long
    Dim packsSheetLastCellRow As Long
    
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
    
    'adapte col width
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    
    ActiveSheet.Range("A1").Select
End Sub
