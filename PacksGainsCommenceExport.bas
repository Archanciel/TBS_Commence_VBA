Attribute VB_Name = "PacksGainsCommenceExport"
Option Explicit
Private Const BONUS_ACHAT_PACK_PAR_FILLEUL As String = "Bonus achat pack par filleul"
Private Const GAIN_PACK_25_PCT As String = "Gain pack 25 %"
Private Const BONUS_FILLEUL_MATRICE_PREMIUM = "Bonus matrice premium"
Private Const BONUS_FILLEUL_MATRICE_ELITE = "Bonus matrice premium"
Private Const BONUS_FILLEUL_UPGRADE = "Bonus filleul ugrade"

'Formate les données issues des copy/paste des listes de packs en vue de leur importation
'dans Commence
Sub packsFormatAndSortData()
    Application.ScreenUpdating = False
    formatDateAchat "DATE_ACHAT"
    transformType
    transformMontantPack
    transformMontantGain "GAIN_TOTAL"
    replaceEnCoursByZeroEchuByOne
    setDateUpdateToToday
    triPourDefinitionRang
    writeNomComptes
    buildLookupTables
    Sheets("Packs").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("A1").Select
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
End Sub

'Exporte les données de la feuille Packs dans un fichier texte tab separated pouvant être importé dans Commence
Sub packsExportDataForCommence()
    Application.ScreenUpdating = False
    deleteNomComptes
    deleteTopRow
    saveAsTabDelimTxtFile "Packs JPS et filleuls Commence export.txt"
    Application.ScreenUpdating = True
    MsgBox "La version .txt de la spreadsheet va être fermée. Veuillez rouvrir la version .xlsm !", vbInformation
    ActiveWorkbook.Close savechanges:=False
End Sub

'Exporte les données de la feuille Gains dans un fichier texte tab separated pouvant être importé dans Commence
Sub gainsExportDataForCommence()
    Application.ScreenUpdating = False
    deleteTopRow
    saveAsTabDelimTxtFile "Gains JPS et filleuls Commence export.txt"
    ActiveSheet.Range("A1").Select
    Application.ScreenUpdating = True
End Sub

'Formate et traite les données issues des copy/paste des listes de gains en vue de leur
'importation dans Commence
Sub handleRevenues()
    Dim rng As Range
    Dim cell As Range
    Dim packId As String
    Dim gainPackMonth As String
    Dim pseudoFilleul As String
    Dim matriceLevel As String
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
    
    Application.ScreenUpdating = False
    
    Set rng = Range("LIBELLE")
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
    Set lookupRangeContratPseudo = lookupTablesSheet.Range(lookupTablesSheet.Cells(2, 4), lookupTablesSheet.Cells(lastCellRow, 5))

    For Each cell In rng
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        packId = extractPackIdFromBonusLibelle(cell)
        
        If (packId <> "") Then
            'gain de type 8 % sur achat de pack par un filleul du détenteur du compte
            Cells(curRow, packIdCol).Value = packId
            Cells(curRow, idGainCol).Value = packId
            Cells(curRow, typeGainCol).Value = BONUS_ACHAT_PACK_PAR_FILLEUL
            formatPseudoFilleulForPackId packId, curRow, pseudoFilleulCol, lookupRangePackContrat, lookupRangeContratPseudo
        Else
            packId = extractPackIdFromGainPackLibelle(cell)
            If (packId <> "") Then
                'gain de 25 % rapporté par un packs du compte
                gainPackMonth = extractPackMonthFromGainPackLibelle(cell)
                Cells(curRow, packIdCol).Value = packId
                Cells(curRow, idGainCol).Value = packId & "-" & gainPackMonth
                Cells(curRow, typeGainCol).Value = GAIN_PACK_25_PCT
            Else
                pseudoFilleul = extractPseudoFilleulMatrixPrem(cell)
                If (pseudoFilleul <> "") Then
                    'bonus mensuel comptabilisé dans la matrice Premium
                    Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                    Cells(curRow, typeGainCol).Value = BONUS_FILLEUL_MATRICE_PREMIUM
                    Cells(curRow, idGainCol).Value = pseudoFilleul & "-BMP-" & Cells(curRow, dateGainCol).Value
                    matriceLevel = extractMatriceLevelMatrixPrem(cell)
                    Cells(curRow, matriceLevelCol).Value = matriceLevel
                Else
                    pseudoFilleul = extractFilleulUpgr(cell)
                    If (pseudoFilleul <> "") Then
                        'bonus provenant de l'activation ou de l'upgrade d'un filleul du détenteur du compte
                        Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                        Cells(curRow, typeGainCol).Value = BONUS_FILLEUL_UPGRADE
                        Cells(curRow, idGainCol).Value = pseudoFilleul & "-UPGR-" & Cells(curRow, dateGainCol).Value
                    Else
                        Cells(curRow, typeGainCol).Value = "### LIBELLE DE GAIN INCONNU ###"
                        MsgBox "Libellé de gain inconnu dans cellule " & cell.Address
                        Exit For
                    End If
                End If
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
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

'Extrait du libellé d'annonce de bonus le numéro de pack dont l'achat par un filleul
'a généré le bonus.
'
'Exemple de libellé: Bonus sponsor pour dépot(#13441058360)
Private Function extractPackIdFromBonusLibelle(cell As Range) As String
    extractPackIdFromBonusLibelle = extractItem(cell, "dépot\(#([0-9]{11})\)$")
End Function

'Extrait du libellé d'annonce de gain de pack le numéro de pack  qui a généré le gain.
'
'Exemple de libellé: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractPackIdFromGainPackLibelle(cell As Range) As String
    extractPackIdFromGainPackLibelle = extractItem(cell, "^#([0-9]{11})")
End Function

'Extrait du libellé d'annonce de gain de pack le numéro du mois du gain.
'
'Exemple de libellé: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractPackMonthFromGainPackLibelle(cell As Range) As String
    extractPackMonthFromGainPackLibelle = extractItem(cell, "([0-9]+)/12\]$")
End Function

'Extrait du libellé d'annonce de bonus matrice premium le pseudo du filleul.
'
'Exemle de libellé: Niveau réseau Premium#1 bonus (tamcerise)
Private Function extractPseudoFilleulMatrixPrem(cell As Range) As String
    extractPseudoFilleulMatrixPrem = extractItem(cell, "^Niveau réseau Premium#\d* bonus \(([a-zA-Z0-9-_]+)\)")
End Function

'Extrait du libellé d'annonce de bonus matrice premium le niveau matriciel du gain.
'
'Exemle de libellé: Niveau réseau Premium#1 bonus (tamcerise)
Private Function extractMatriceLevelMatrixPrem(cell As Range) As String
    extractMatriceLevelMatrixPrem = extractItem(cell, "^Niveau réseau Premium#(\d*) bonus")
End Function

'Extrait le pseudo du filleul du libellé d'annonce de bonus en cas d'upgrade de celui-ci.
'
'Exemle de libellé: Bonus sponsor (rosemaman)
Private Function extractFilleulUpgr(cell As Range) As String
    extractFilleulUpgr = extractItem(cell, "^Bonus sponsor \(([a-zA-Z0-9-_]+)\)")
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

Private Sub formatDateAchat(colName As String)
Attribute formatDateAchat.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formatDateAchat Macro
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
Private Sub transformMontantPack()
'
' transformMontant Macro
'

'
    ActiveSheet.Range("MONTANT_PACK").Select
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
Private Sub saveAsTabDelimTxtFile(fileName As String)
    ActiveWorkbook.SaveAs fileName:= _
        "D:\Users\Jean-Pierre\OneDrive\Documents\Excel\" & fileName _
        , FileFormat:=xlText, CreateBackup:=False
End Sub

'Supprime la ligne contenant les en-têtes de colonnes afin qu'elles ne soient pas exportées.
'
'Cette suppression n'affecte que la version txt de la speadsheet et non la version xlsm !
Private Sub deleteTopRow()
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
End Sub

'Recrée la  zone NOM_COMPTES qui contient les noms de contrats TBS dans Commence.
'Ces noms sont utilisés en copy/paste lors de l'entrée de nouvelles données dans
'la feuille Packs
Private Sub writeNomComptes()
    ActiveSheet.Range("NOM_COMPTES").Select
    ActiveCell.FormulaR1C1 = "Compte TBS Antoine"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "Compte TBS Béatrice"
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
'En effet, ces données ne doivent pas être exportées !
'
'Ces noms sont utilisés en copy/paste lors de l'entrée de nouvelles données dans
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
    Dim lastCellRow As Long
    
    Set lookupTablesSheet = Sheets("Lookup tables")
    Set packsSheet = Sheets("Packs")
    
    'vide la table
    lastCellRow = getLastDataRow(lookupTablesSheet.Range("A:A"))
    Set lookupRangePackContrat = lookupTablesSheet.Range(lookupTablesSheet.Cells(2, 1), lookupTablesSheet.Cells(lastCellRow, 2))
    lookupRangePackContrat.Cells.ClearContents
    
    'copie la colonne no de packs
    packsSheet.Select
    packsSheet.Columns("D:D").Select
    Selection.Copy
    lookupTablesSheet.Select
    lookupTablesSheet.Columns("A:A").Select
    ActiveSheet.Paste
    
    'copie la colonne nom de contrat
    packsSheet.Select
    packsSheet.Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    lookupTablesSheet.Select
    lookupTablesSheet.Columns("B:B").Select
    ActiveSheet.Paste
    
    'adapte col width
    Columns("B:B").EntireColumn.AutoFit
    Columns("A:A").EntireColumn.AutoFit
    ActiveSheet.Range("A1").Select
End Sub
