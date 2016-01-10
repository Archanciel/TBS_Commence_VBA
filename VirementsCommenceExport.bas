Attribute VB_Name = "VirementsCommenceExport"
Option Explicit

Private Const PAIEMENT_TYPE_ACHAT_PACK As String = "Achat pack"
Private Const PAIEMENT_TYPE_MEMBERSHIP_SE As String = "Cotisation SE"
Private Const PAIEMENT_TYPE_MEMBERSHIP_PREMIUM As String = "Cotisation Premium"

Private Const TYPE_VIREMENT_DE As String = "Transfert de"
Private Const TYPE_VIREMENT_A As String = "Transfert �"

'Formate et traite les donn�es issues des copy/paste des listes de virements en vue de leur
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
    
    'clear col 8 � 10 qui contiennent les valeurs extraites par la suite de la macro
    Set virementSheetCalculatedCellsRange = ActiveSheet.Range(ActiveSheet.Cells(2, typeVirementCol), ActiveSheet.Cells(lastCellRow, uidVirementCol))
    virementSheetCalculatedCellsRange.Clear
    
    'au d�but, je n'utilisais pas syst�matiquement le tag #TRANSTEMP !!
    replaceInRange rngLibelle, "Retransfers", "#TRANSTEMP", False
    
    'R�gles de gestion:
    '
    'pour chaque cellule de la colonne LIBELLE_VIREMENT,
    '   si le libell� contient #TRANSTEMP
    '       si le MONTANT_VIREMENT est positif
    '           type virement = TYPE_VIREMENT_DE (transfert de)
    '       si le MONTANT_VIREMENT est n�gatif
    '           type virement = TYPE_VIREMENT_A (transfert �)
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
            'paiement pour un achat de pack (d�p�t)
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
                    MsgBox "Libell� de paiement inconnu dans cellule " & cell.Address
                    Exit For
                End If
            End If
        End If
    Next cell
    
    clearAnySelection
    
    Application.ScreenUpdating = True
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

