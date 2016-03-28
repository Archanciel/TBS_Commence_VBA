Attribute VB_Name = "PaiementsCommenceExport"
Option Explicit

Private Const PAIEMENT_TYPE_ACHAT_PACK As String = "Achat pack"
Private Const PAIEMENT_TYPE_MEMBERSHIP_OMEGA As String = "Cotisation Omega"
Private Const PAIEMENT_TYPE_MEMBERSHIP_SE As String = "Cotisation SE"
Private Const PAIEMENT_TYPE_MEMBERSHIP_PREMIUM As String = "Cotisation Premium"

'Formate et traite les donn�es issues des copy/paste des listes de paiements en vue de leur
'importation dans Commence
Sub handlePaiements()
    Dim rngLibelle As Range
    Dim cell As Range
    Dim paiemenrPackId As String
    Dim curRow As Long
    Dim paiementPackIdCol As Long
    Dim typePaiementCol As Long
    Dim timePaiementCol As Long
    Dim lastCellRow As Long
    Dim paiementSheetCalculatedCellsRange As Range
    
    Application.ScreenUpdating = False
    
    terminateIfNoData Cells(2, Range("PAIEMENT_LIBELLE").Column)
    
    formatDateAndTime "DATE_OP", "TIME_OP"
    transformMontant "MONTANT_PAIEMENT"
    formatIdCol "PAIEMENT_ID"
    
    lastCellRow = getLastDataRow(ActiveSheet.Range("A:A"))
    Set rngLibelle = getDataRangeFromColRange(ActiveSheet.Range("PAIEMENT_LIBELLE"))
    
    paiementPackIdCol = Range("PAIEMENT_PACK_ID").Column
    typePaiementCol = Range("PAIEMENT_TYPE").Column
    
    'clear col 9 � 10 qui contiennent les valeurs extraites par la suite de la macro
    Set paiementSheetCalculatedCellsRange = ActiveSheet.Range(ActiveSheet.Cells(2, paiementPackIdCol), ActiveSheet.Cells(lastCellRow, typePaiementCol))
    paiementSheetCalculatedCellsRange.Clear
    
    For Each cell In rngLibelle
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        paiemenrPackId = extractPackIdFromLibelleDepotCell(cell)
        
        If (paiemenrPackId <> "") Then
            'paiement pour un achat de pack (d�p�t)
            Cells(curRow, paiementPackIdCol).Value = paiemenrPackId
            Cells(curRow, typePaiementCol).Value = PAIEMENT_TYPE_ACHAT_PACK
        Else
            paiemenrPackId = extractPackIdFromLibelleSEMembershipCell(cell)
            If (paiemenrPackId <> "") Then
                'paiement pour cotisation SE
                Cells(curRow, typePaiementCol).Value = PAIEMENT_TYPE_MEMBERSHIP_SE
            Else
                paiemenrPackId = extractPackIdFromLibellePremiumMembershipCell(cell)
                If (paiemenrPackId <> "") Then
                    'paiement pour cotisation Premium
                    Cells(curRow, typePaiementCol).Value = PAIEMENT_TYPE_MEMBERSHIP_PREMIUM
                Else
                    paiemenrPackId = extractPackIdFromLibelleOmegaMembershipCell(cell)
                    If (paiemenrPackId <> "") Then
                        'paiement pour cotisation Omega
                        Cells(curRow, typePaiementCol).Value = PAIEMENT_TYPE_MEMBERSHIP_OMEGA
                    Else
                        MsgBox "Libell� de paiement inconnu dans cellule " & cell.Address & " !", vbInformation
                        Exit For
                    End If
                End If
            End If
        End If
    Next cell
    
    formatIdCol ("PAIEMENT_PACK_ID")
    clearAnySelection
    
    Application.ScreenUpdating = True
End Sub
'Exporte les donn�es de la feuille Paiements dans un fichier texte tab separated pouvant �tre import� dans Commence
Sub paiementsExportDataForCommence()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    deleteNomComptes "NOM_COMPTES_P"
    deleteTopRow
    saveSheetAsTabDelimTxtFileTimeStamped ActiveSheet.Name
    Application.ScreenUpdating = True
    closeWithoutSave
End Sub
'Extrait du libell� contenu dans la Cell pass� en parm le num�ro de pack
'qu'il contient.
'
'Pr�cond: le no de pack se trouve au d�but du libell� !
'
'Exemple de libell�: #12934088431 Paiement de d�pot
Private Function extractPackIdFromLibelleDepotCell(cell As Range) As String
    extractPackIdFromLibelleDepotCell = extractItem(cell, "^#([0-9]+) Paiement de d�pot")
End Function
'Extrait du libell� contenu dans la Cell pass� en parm le num�ro de pack
'qu'il contient.
'
'Pr�cond: le no de pack se trouve au d�but du libell� !
'
'Exemple de libell�: #129340130641 SVIP payment
Private Function extractPackIdFromLibelleSEMembershipCell(cell As Range) As String
    extractPackIdFromLibelleSEMembershipCell = extractItem(cell, "^#([0-9]{12}) SVIP payment")
End Function
'Extrait du libell� contenu dans la Cell pass� en parm le num�ro de pack
'qu'il contient.
'
'Pr�cond: le no de pack se trouve au d�but du libell� !
'
'Exemple de libell�: #129340104707 Membership payment ou #12934081869 R�glement adh�sion
Private Function extractPackIdFromLibellePremiumMembershipCell(cell As Range) As String
    extractPackIdFromLibellePremiumMembershipCell = extractItem(cell, "^#([0-9]*) (Membership payment|R�glement adh�sion)")
End Function
'Extrait du libell� contenu dans la Cell pass� en parm le num�ro de pack
'qu'il contient.
'
'Pr�cond: le no de pack se trouve au d�but du libell� !
'
'Exemple de libell�: #134390321226 OMEGA payment
Private Function extractPackIdFromLibelleOmegaMembershipCell(cell As Range) As String
    extractPackIdFromLibelleOmegaMembershipCell = extractItem(cell, "^#([0-9]*) Omega payment")
End Function

