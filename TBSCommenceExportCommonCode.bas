Attribute VB_Name = "TBSCommenceExportCommonCode"
Option Explicit
'Ce module contient des macros, procédures et fonctions communes à tous les
'modules spécifiques aux feuilles du workbook

Sub MacroCopySelectedAccountNameInEmptyCells()
Attribute MacroCopySelectedAccountNameInEmptyCells.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' Copie le nom de compte actif dans les cellules vides de la colonne COMPTE
'
' Touche de raccourci du clavier: Ctrl+Shift+N
'
    Dim firstEmptyCellRow As Long
    Dim lastEmptyCellRow As Long
    Dim emptyAccountNameCells As Range
    
    Selection.Copy
    firstEmptyCellRow = getLastDataRow(Range("A:A")) + 1
    
    If (firstEmptyCellRow > 1000000) Then
        'le cas si la colonne contenant les noms de compte est vide !
        firstEmptyCellRow = 2
    End If
    
    lastEmptyCellRow = getLastDataRow(Range("B:B"))
    
    If (firstEmptyCellRow > lastEmptyCellRow) Then
        MsgBox "Aucune cellule vide où copier le nom du compte détectée !", vbInformation
        Exit Sub
    End If
    
    Set emptyAccountNameCells = ActiveSheet.Range(Cells(firstEmptyCellRow, 1), Cells(lastEmptyCellRow, 1))
    
    emptyAccountNameCells.Select
    ActiveSheet.Paste
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

'remplacement d'un string par un autre dans le range passé en parm
Sub replaceInRange(replaceRange As Range, strToReplace As String, replacementStr As String, boolMatchCase As Boolean)
'
' transformType Macro
'

'
    replaceRange.Select
    
    'handling xmas pack denomination
    Selection.replace What:=strToReplace, Replacement:=replacementStr, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=boolMatchCase, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

