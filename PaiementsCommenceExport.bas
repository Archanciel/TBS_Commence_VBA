Attribute VB_Name = "PaiementsCommenceExport"
Option Explicit


'Formate et traite les données issues des copy/paste des listes de paiements en vue de leur
'importation dans Commence
Sub handlePaiements()
    Dim rngLibelle As Range
    Dim cell As Range
    Dim packId As String
    Dim compteSubjectToPaiementCol As Long
    Dim idPaiementCol As Long
    Dim matriceLevelCol As Long
    Dim curRow As Long
    Dim packIdCol As Long
    Dim typePaiementCol As Long
    Dim datePaiementCol As Long
    Dim timePaiementCol As Long
    Dim lastCellRow As Long
    
    Application.ScreenUpdating = False
    
    formatDateAndTime "DATE_OP", "TIME_OP"
    
    clearAnySelection
    
    Application.ScreenUpdating = True
End Sub
