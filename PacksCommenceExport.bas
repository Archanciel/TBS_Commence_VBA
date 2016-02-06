Attribute VB_Name = "PacksCommenceExport"
Option Explicit
Private Const RENDEMENT_REGULAR_PACK As String = "25"
Private Const RENDEMENT_XMAS_PACK As String = "28"

'Formate les données issues des copy/paste des listes de packs en vue de leur importation
'dans Commence
Sub packsFormatAndSortData()
    Application.ScreenUpdating = False
    
    terminateIfNoData Cells(2, Range("ECHU").Column)
    
    formatIdCol ("NOM_PACK")
    formatDateAndTime "DATE_ACHAT", "TIME_ACHAT_PACK"
    'adapte col width for id pack
    Columns("D:D").EntireColumn.AutoFit
    formatRendement 'doit être appelé avant transformType !!
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
                'il arrive en effet que la feuille pack a été processée auparavant et que l'on rajoute une ligne
                'pour un nouveau pack. Si la feuille contenait des lignes pour, par ex, un xmas pack 1000, le type
                'de pack ayant été transformé en Bronze, sans cette garde, le rendement de 28 % serait écrasé en 25 % !!
                '
                'I know, this stinks, but I chosed not to create additional pack types !
                Cells(curRow, rendementCol).Value = RENDEMENT_REGULAR_PACK
            End If
        End If
    Next cell
End Sub
'Exporte les données de la feuille Packs dans un fichier texte tab separated pouvant être importé dans Commence
Sub packsExportDataForCommence()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    deleteNomComptes "NOM_COMPTES"
    deleteTopRow
    saveSheetAsTabDelimTxtFileTimeStamped ActiveSheet.Name
    Application.ScreenUpdating = True
    closeWithoutSave
End Sub
Private Sub transformType()
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
    ActiveSheet.Range("ECHU").Select
    Selection.replace What:="En cours", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
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

'Sub appelée par packsFormatAndSortData()
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
        'le cas si la lookup table ne contient aucune entrée pack/compte/date pack !
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
    
    'copie la colonne date achat (utile pour purger les packs plus vieux d'une année de la lookup table !)
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



