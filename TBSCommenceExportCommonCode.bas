Attribute VB_Name = "TBSCommenceExportCommonCode"
Option Explicit
'Ce module contient des macros, procédures et fonctions communes à tous les
'modules spécifiques aux feuilles du workbook
Private Const COMMENCE_DATA_EXPORT_DIR As String = "D:\Users\Jean-Pierre\OneDrive\Documents\Commence\FILES\"

Sub MacroCopySelectedAccountNameInEmptyCells()
Attribute MacroCopySelectedAccountNameInEmptyCells.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' Copie le nom de compte actif dans les cellules vides de la colonne COMPTE
'
' Touche de raccourci du clavier: Ctrl+Shift+N (N for account Name)
'
    Dim firstEmptyCellRow As Long
    Dim lastEmptyCellRow As Long
    Dim emptyAccountNameCells As Range
    Dim lastFilledCellRow As Long
    
    Selection.Copy
    firstEmptyCellRow = getLastDataRow(Range("A:A")) + 1
    lastFilledCellRow = getLastDataRowFromEnd(Range("A:A"))
    
    
    If (firstEmptyCellRow > 1000000) Then
        'le cas si la colonne contenant les noms de compte est vide !
        firstEmptyCellRow = 2
    End If
    
    lastEmptyCellRow = getLastDataRow(Range("B:B"))
    
    If (firstEmptyCellRow > lastEmptyCellRow) Then
        MsgBox "Aucune cellule vide où copier le nom du compte détectée: macro interrompue sans modification de la spreadsheet !", vbInformation
        Exit Sub
    End If
    
    If (firstEmptyCellRow < lastFilledCellRow) Then
        MsgBox "Cellule(s) vide(s) suivie(s) de cellule(s) non vide(s): macro interrompue sans modification de la spreadsheet !", vbInformation
        Exit Sub
    End If
    
    Set emptyAccountNameCells = ActiveSheet.Range(Cells(firstEmptyCellRow, 1), Cells(lastEmptyCellRow, 1))
    
    emptyAccountNameCells.Select
    ActiveSheet.Paste
End Sub

Sub MacroFormatMatchingTagForMultipleTranstempRows()
Attribute MacroFormatMatchingTagForMultipleTranstempRows.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' 1/ build a matching string tag using the smallest date/time stamp of the selected rows (not necessarily contiguous).
' 2/ copy the matching tag into the empty TRANSTEMP_MATCHING_MANUAL_TAG cell of the multi selected rows
'
' Touche de raccourci du clavier: Ctrl+Shift+M (M for Matching)
'
    Dim rngSelection As Range
    Dim rngSelectionArea As Range
    Dim rngCell As Range
    Dim lngRow As Long
    Dim lngActualRow As Long
    Dim strMatchingTag As String
    Dim strDateTimeStamp As String
    Dim lngDateVirementCol As Long
    Dim lngMatchingTagCol As Long
    Dim datTmpDate As Date
    Dim datSmallestDate As Date
    Dim strSmallestDate As String
    
    On Error GoTo errorhandler
    
    ensureWorksheetIs ("Transferts-virements")
    lngDateVirementCol = Range("DATE_VIREMENT").Column
    lngMatchingTagCol = Range("TRANSTEMP_MATCHING_MANUAL_TAG").Column
    datSmallestDate = Now

    ' Get the current selection
    Set rngSelection = Application.Selection

    ' Walk through the areas to get smallest operation date
    For Each rngSelectionArea In rngSelection.Areas
        ' Walk through the rows
        For lngRow = 1 To rngSelectionArea.Rows.Count Step 1
            ' Get the row reference
            Set rngCell = rngSelectionArea.Rows(lngRow)

            ' Get the actual row index (in the worksheet).
            ' The other row index is relative to the collection.
            lngActualRow = rngCell.Row

            ' Get any cell value by using the actual row index
            ' Example:
            strDateTimeStamp = ActiveSheet.Cells(lngActualRow, lngDateVirementCol).Value
            datTmpDate = CDate(strDateTimeStamp)
            
            If (datSmallestDate > datTmpDate) Then
                datSmallestDate = datTmpDate
                strSmallestDate = strDateTimeStamp
            End If
        Next
    Next
    
    'stripping out last ":00" chars
    strSmallestDate = Left(strSmallestDate, Len(strSmallestDate) - 3)
    strMatchingTag = "A-" & strSmallestDate

    ' Walk through the areas to paste the matching tag
    For Each rngSelectionArea In rngSelection.Areas
        ' Walk through the rows
        For lngRow = 1 To rngSelectionArea.Rows.Count Step 1
            ' Get the row reference
            Set rngCell = rngSelectionArea.Rows(lngRow)

            ' Get the actual row index (in the worksheet).
            ' The other row index is relative to the collection.
            lngActualRow = rngCell.Row

            ' Get any cell value by using the actual row index
            ' Example:
            ActiveSheet.Cells(lngActualRow, lngMatchingTagCol).Value = strMatchingTag
        Next
    Next
    
    Exit Sub
errorhandler:
    MsgBox "Erreur: " & Err.Number & vbCrLf & Err.Description & ". Macro interrompue !", vbCritical
    Exit Sub
End Sub
Sub ensureWorksheetIs(strWorksheetName As String)
    If (ActiveSheet.Name <> strWorksheetName) Then
        MsgBox "Cette macro ne s'applique qu'à la feuille " & strWorksheetName & " !", vbExclamation
        End
    End If
End Sub
Sub clearAnySelection()
    Application.CutCopyMode = False
    ActiveSheet.Range("A1").Select
End Sub

'A partir de la colonne contenant une date + heure de ce fornat: 26.11.2015 18:26:00,
'extrait l'heure sans les secondes et la place dans la colonne timeColName. La date
'sans l'heure remplace la date initiale dans la colonne datsTimeColName
Sub formatDateAndTime(dateTimeColName As String, timeColName As String)
    Dim dateTimeRange As Range
    Dim timeRange As Range
    Dim lastDateTimeCellRow As Long
    Dim timeCellCol As Long
    Dim cell As Range
    
    Set dateTimeRange = ActiveSheet.Range(dateTimeColName)
    lastDateTimeCellRow = getLastDataRow(dateTimeRange)
    Set dateTimeRange = ActiveSheet.Range(dateTimeRange.Cells(2, 1), dateTimeRange.Cells(lastDateTimeCellRow, 1))
    
    Set timeRange = ActiveSheet.Range(timeColName)
    timeCellCol = timeRange.Cells(1, 1).Column  'obtaining the col number of the column containg the opervation time
       
    For Each cell In dateTimeRange
        splitDateTime cell.Row, cell.Column, timeCellCol
    Next cell
    
    dateTimeRange.Select
    Selection.NumberFormat = "d/m/yyyy"
    
    timeRange.Select
    Selection.NumberFormat = "hh:mm"
End Sub

Sub splitDateTime(rowNum As Long, initialDateColNum As Long, timeColNum As Long)
    Dim dateTimeString As String
    Dim l, n, m As Integer
    Dim timePartWithSec As String
    Dim timePartNoSec As String
    Dim datePart As String
 
    ' Cache the original value
    '
    dateTimeString = Cells(rowNum, initialDateColNum).Value

    ' Calculate the length and the location where the white space is placed
    '
    l = Len(dateTimeString)
    n = InStr(1, dateTimeString, " ")
    
    ' Separate the date and time strings
    '
    datePart = Left(dateTimeString, n)
    timePartWithSec = Right(dateTimeString, l - n)
    
    If (timePartWithSec = vbNullString) Then
        'the case if the date / time was already split !
        Exit Sub
    End If
    
    m = Len(timePartWithSec)
    timePartNoSec = Left(timePartWithSec, m - 3) 'strip out seconds

    ' Make sure the fields are text
    '
    Cells(rowNum, timeColNum).NumberFormat = "@"
    Cells(rowNum, initialDateColNum).NumberFormat = "@"

    ' Populate the date and time cells
    '
    Cells(rowNum, timeColNum).Value = timePartNoSec
    Cells(rowNum, initialDateColNum).Value = datePart
End Sub
'make amount values nicely right aligned like real numbers
Sub transformMontant(colName As String)
    ActiveSheet.Range(colName).Select
    Selection.replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
'so that the UID's are displayed correctly, and not in scientific notation
Sub formatIdCol(colName As String)
    ActiveSheet.Range(colName).Select
    Selection.NumberFormat = "0"
End Sub
Function extractItem(cell As Range, regexp As String) As String
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

'reçois un whole column range en parm et renvoie le même range, mais
'amputé de la première cellule qui contient le titre de la colonne.

'Ex d'utilisation:
'    Set packTypeRange = getDataRangeFromColRange(ActiveSheet.Range("TYPE"))
Function getDataRangeFromColRange(colRange As Range) As Range
    Dim lastColRangeRow As Long
    lastColRangeRow = getLastDataRow(colRange)
    Set getDataRangeFromColRange = ActiveSheet.Range(colRange.Cells(2, 1), colRange.Cells(lastColRangeRow, 1))
End Function
Function getLastDataRow(colCell As Range) As Long
    Dim lastCell As Range
    Dim lastCellRow As Long
    
    Set lastCell = colCell.End(xlDown)
    getLastDataRow = lastCell.Row
End Function
Function getLastDataRowFromEnd(colCell As Range) As Long
'Find the last used row in a Column: column A in this example
    Dim lastRow As Long
    
    With ActiveSheet
        lastRow = .Cells(.Rows.Count, colCell.Column).End(xlUp).Row
    End With
    
    getLastDataRowFromEnd = lastRow
End Function
'remplacement d'un string par un autre dans le range passé en parm
Sub replaceInRange(replaceRange As Range, strToReplace As String, replacementStr As String, boolMatchCase As Boolean)
    replaceRange.Select
    
    'handling xmas pack denomination
    Selection.replace What:=strToReplace, Replacement:=replacementStr, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=boolMatchCase, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

'Sauve une feuille spécifique dans un fichier txt tab delimited
'en ajoutant la date et l'heure courante au nom de fichier sauvé
Sub saveSheetAsTabDelimTxtFileTimeStamped(sheetName As String)
    Dim currentDateTimeStr As String
    
    currentDateTimeStr = Format(Now(), "yyyy-mm-dd_hh.mm.ss")
    saveSheetAsTabDelimTxtFile sheetName, sheetName & "_Comm_imp_" & currentDateTimeStr & ".txt"
End Sub

'Sauve une feuille spécifique dans un fichier txt tab delimited
Sub saveSheetAsTabDelimTxtFile(sheetName As String, fileName As String)
    Dim ans As Long
    Dim sSaveAsFilePath As String

    On Error GoTo ErrHandler:
    
    sSaveAsFilePath = COMMENCE_DATA_EXPORT_DIR & fileName

    If Dir(sSaveAsFilePath) <> "" Then
        ans = MsgBox("Le fichier " & sSaveAsFilePath & " existe déjà. Remplacer ?", vbYesNo + vbExclamation)
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

'Ferme le workbook sans le sauver
Sub closeWithoutSave()
    MsgBox "La version modifiée (sans ligne de titres) de la spreadsheet va être fermée sans être sauvée. Veuillez rouvrir la version .xlsm (sauvée avant l'exportation) !", vbInformation
    ActiveWorkbook.Close savechanges:=False
End Sub

'Delete la  zone NOM_COMPTES qui contient les noms de contrats TBS dans Commence.
'En effet, ces données ne doivent pas être exportées !
'
'Ces noms sont utilisés en copy/paste lors de l'entrée de nouvelles données dans
'la feuille Packs
Sub deleteNomComptes(rangeNameStr As String)
    ActiveSheet.Range(rangeNameStr).Select
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

'Supprime la ligne contenant les en-têtes de colonnes afin qu'elles ne soient pas exportées.
'
'Cette suppression n'affecte que la version txt de la speadsheet et non la version xlsm !
Sub deleteTopRow()
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
End Sub

'Stop processing if the passed range contains no data
Sub terminateIfNoData(testedRange As Range)
    If (isRangeEmpty(testedRange)) Then
        MsgBox "Feuille vide, pas de données à processer, macro interrompue !", vbExclamation
        clearAnySelection
        Application.ScreenUpdating = True
        End
    End If
End Sub

Private Function isRangeEmpty(testedRange As Range) As Boolean
    Dim neValues As Range
    Dim neFormulas As Range
    
    On Error Resume Next
    Set neValues = Intersect(Cells(2, 1).EntireRow.SpecialCells(xlConstants), testedRange)
    Set neFormulas = Intersect(Cells(2, 1).EntireRow.SpecialCells(xlFormulas), testedRange)
    On Error GoTo 0
    
    If neValues Is Nothing And neFormulas Is Nothing Then
        isRangeEmpty = True
    Else
        isRangeEmpty = False
    End If
End Function

Sub clearSheet()
    Dim delRange As Range
    Dim topLeftTitleCell As Range
    Dim topLeftDataCell As Range
    Dim ans As Long
    
    Set topLeftTitleCell = Range("A1")
    Set topLeftDataCell = Range("A2")
    Set delRange = Range(topLeftDataCell, Cells(topLeftDataCell.End(xlDown).Row, topLeftTitleCell.End(xlToRight).Column))
    
    delRange.Select
    ans = MsgBox("Supprimer la sélection ?", vbYesNo + vbExclamation)
    
    If ans <> vbYes Then
        clearAnySelection
        Exit Sub
    Else
        delRange.Clear
    End If
End Sub





