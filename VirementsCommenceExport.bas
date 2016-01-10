Attribute VB_Name = "VirementsCommenceExport"
Option Explicit

Private Const PAIEMENT_TYPE_ACHAT_PACK As String = "Achat pack"
Private Const PAIEMENT_TYPE_MEMBERSHIP_SE As String = "Cotisation SE"
Private Const PAIEMENT_TYPE_MEMBERSHIP_PREMIUM As String = "Cotisation Premium"

Private Const TYPE_VIREMENT_DE As String = "Transfert de"
Private Const TYPE_VIREMENT_A As String = "Transfert à"

'Formate et traite les données issues des copy/paste des listes de virements en vue de leur
'importation dans Commence
Sub handleVirements()
    Dim rngLibelle As Range
    Dim cell As Range
    Dim paiemenrPackId As String
    Dim curRow As Long
    Dim uidVirementCol As Long
    Dim typeVirementCol As Long
    Dim timePaiementCol As Long
    Dim lastCellRow As Long
    Dim virementSheetCalculatedCellsRange As Range
    
    Application.ScreenUpdating = False
    
    formatDateAndTime "DATE_VIREMENT", "TIME_VIREMENT"
    transformMontant "MONTANT_VIREMENT"
    
    lastCellRow = getLastDataRow(ActiveSheet.Range("A:A"))
    Set rngLibelle = getDataRangeFromColRange(ActiveSheet.Range("LIBELLE_VIREMENT"))
    
    uidVirementCol = Range("UID_VIREMENT").Column
    typeVirementCol = Range("TYPE_VIREMENT").Column
    
    'clear col 8 à 10 qui contiennent les valeurs extraites par la suite de la macro
    Set virementSheetCalculatedCellsRange = ActiveSheet.Range(ActiveSheet.Cells(2, typeVirementCol), ActiveSheet.Cells(lastCellRow, uidVirementCol))
    virementSheetCalculatedCellsRange.Clear
    
    'au début, je n'utilisais pas systématiquement le tag #TRANSTEMP !!
    replaceInRange rngLibelle, "Retransfers", "#TRANSTEMP", False
    
    'Règles de gestion:
    '
    'pour chaque cellule de la colonne LIBELLE_VIREMENT,
    '   si le libellé contient #TRANSTEMP
    '       si le MONTANT_VIREMENT est positif
    '           type virement = TYPE_VIREMENT_DE (transfert de)
    '       si le MONTANT_VIREMENT est négatif
    '           type virement = TYPE_VIREMENT_A (transfert à)
    '       end if
    '       pseudo_virement = extract pseudo
    '       compte contrepartie = getCompteForPseudo() from lookup table
    '   sinon, si
    For Each cell In rngLibelle
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        paiemenrPackId = extractPackIdFromLibelleDepotCell(cell)
        
        If (paiemenrPackId <> "") Then
            'paiement pour un achat de pack (dépôt)
            Cells(curRow, uidVirementCol).Value = paiemenrPackId
            Cells(curRow, typeVirementCol).Value = PAIEMENT_TYPE_ACHAT_PACK
        Else
            paiemenrPackId = extractPackIdFromLibelleSEMembershipCell(cell)
            If (paiemenrPackId <> "") Then
                'paiement pour cotisation SE
                Cells(curRow, typeVirementCol).Value = PAIEMENT_TYPE_MEMBERSHIP_SE
            Else
                paiemenrPackId = extractPackIdFromLibellePremiumMembershipCell(cell)
                If (paiemenrPackId <> "") Then
                    'paiement pour cotisation Premium
                    Cells(curRow, typeVirementCol).Value = PAIEMENT_TYPE_MEMBERSHIP_PREMIUM
                Else
                    MsgBox "Libellé de paiement inconnu dans cellule " & cell.Address
                    Exit For
                End If
            End If
        End If
    Next cell
    
    clearAnySelection
    
    Application.ScreenUpdating = True
End Sub
'Extrait du libellé contenu dans la Cell passé en parm le numéro de pack
'qu'il contient.
'
'Précond: le no de pack se trouve au début du libellé !
'
'Exemple de libellé: #12934088431 Paiement de dépot
Private Function extractPackIdFromLibelleDepotCell(cell As Range) As String
    extractPackIdFromLibelleDepotCell = extractItem(cell, "^#([0-9]+) Paiement de dépot")
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

