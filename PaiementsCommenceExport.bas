Attribute VB_Name = "PaiementsCommenceExport"
Option Explicit


'Formate et traite les données issues des copy/paste des listes de paiements en vue de leur
'importation dans Commence
Sub handlePaiements()
    Dim rngLibelle As Range
    Dim cell As Range
    Dim paiemenrPackId As String
    Dim idPaiementCol As Long
    Dim curRow As Long
    Dim paiementPackIdCol As Long
    Dim typePaiementCol As Long
    Dim datePaiementCol As Long
    Dim timePaiementCol As Long
    Dim lastCellRow As Long
    Dim paiementSheetCalculatedCellsRange As Range
    
    Application.ScreenUpdating = False
    
    formatDateAndTime "DATE_OP", "TIME_OP"
    transformMontant "MONTANT_PAIEMENT"
    formatIdCol "PAIEMENT_ID"
    
    Set rngLibelle = Range("PAIEMENT_LIBELLE")
    
    paiementPackIdCol = Range("PAIEMENT_PACK_ID").Column
    typePaiementCol = Range("PAIEMENT_TYPE").Column
    idPaiementCol = Range("PAIEMENT_ID").Column
    datePaiementCol = Range("DATE_OP").Column
    
    'clear col 9 à 10 qui contiennent les valeurs extraites par la suite de la macro
    lastCellRow = getLastDataRow(ActiveSheet.Range("A:A"))
    Set paiementSheetCalculatedCellsRange = ActiveSheet.Range(ActiveSheet.Cells(2, paiementPackIdCol), ActiveSheet.Cells(lastCellRow, typePaiementCol))
    paiementSheetCalculatedCellsRange.Clear
    
    For Each cell In rngLibelle
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        paiemenrPackId = extractPackIdFromLibelleDepotCell(cell)
        
        If (paiemenrPackId <> "") Then
            'paiement pour un achat de pack (dépôt)
            Cells(curRow, paiementPackIdCol).Value = paiemenrPackId
            Cells(curRow, idPaiementCol).Value = paiemenrPackId & "-b"
            Cells(curRow, typePaiementCol).Value = BONUS_ACHAT_PACK_PAR_FILLEUL
            formatPseudoFilleulForPackId paiemenrPackId, curRow, pseudoFilleulCol, lookupRangePackContrat, lookupRangeContratPseudo
        Else
            paiemenrPackId = extractPackIdFromPaiementPackLibelle(cell)
            If (paiemenrPackId <> "") Then
                'gain de 25 % rapporté par un packs du compte
                gainPackMonth = extractPackMonthFromPaiementPackLibelle(cell)
                Cells(curRow, paiementPackIdCol).Value = paiemenrPackId
                Cells(curRow, idPaiementCol).Value = paiemenrPackId & "-" & gainPackMonth
                Cells(curRow, typePaiementCol).Value = GAIN_PACK_25_PCT
            Else
                pseudoFilleul = extractPseudoFilleulMatrixPrem(cell)
                If (pseudoFilleul <> "") Then
                    'bonus mensuel comptabilisé dans la matrice Premium
                    Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                    Cells(curRow, typePaiementCol).Value = BONUS_FILLEUL_MATRICE_PREMIUM
                    Cells(curRow, idPaiementCol).Value = pseudoFilleul & "-BMP-to-" & Cells(curRow, compteSubjectToPaiementCol).Value & "-" & Format(Cells(curRow, datePaiementCol).Value2, "dd.mm.yy")
                    matriceLevel = extractMatriceLevelMatrixPrem(cell)
                    Cells(curRow, matriceLevelCol).Value = matriceLevel
                Else
                    pseudoFilleul = extractPseudoFilleulMatrixSE(cell)
                    If (pseudoFilleul <> "") Then
                        'bonus mensuel comptabilisé dans la matrice Super Elite
                        Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                        Cells(curRow, typePaiementCol).Value = BONUS_FILLEUL_MATRICE_SE
                        Cells(curRow, idPaiementCol).Value = pseudoFilleul & "-BSE-to-" & Cells(curRow, compteSubjectToPaiementCol).Value & "-" & Format(Cells(curRow, datePaiementCol).Value2, "dd.mm.yy")
                        matriceLevel = extractMatriceLevelMatrixSE(cell)
                        Cells(curRow, matriceLevelCol).Value = matriceLevel
                    Else
                        pseudoFilleul = extractFilleulUpgrToPremium(cell)
                        If (pseudoFilleul <> "") Then
                            'bonus provenant de l'activation ou de l'upgrade en Premium d'un filleul du détenteur du compte
                            Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                            Cells(curRow, typePaiementCol).Value = BONUS_FILLEUL_UPGR_PREMIUM
                            Cells(curRow, idPaiementCol).Value = pseudoFilleul & "-UPGR_PREM-" & Format(Cells(curRow, datePaiementCol).Value2, "dd.mm.yy")
                        Else
                            pseudoFilleul = extractFilleulUpgrToSE(cell)
                            If (pseudoFilleul <> "") Then
                                'bonus provenant de l'upgrade en Super Elite d'un filleul du détenteur du compte
                                Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                                Cells(curRow, typePaiementCol).Value = BONUS_FILLEUL_UPGR_SE
                                Cells(curRow, idPaiementCol).Value = pseudoFilleul & "-UPGR_SE-" & Format(Cells(curRow, datePaiementCol).Value2, "dd.mm.yy")
                            Else
                                MsgBox "Libellé de gain inconnu dans cellule " & cell.Address
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
'Extrait du libellé contenu dans la Cell passé en parm le numéro de pack
'qu'il contient.
'
'Précond: le no de pack se trouve au début du libellé !
'
'Exemple de libellé: #12934088431 Paiement de dépot
Private Function extractPackIdFromLibelleDepotCell(cell As Range) As String
    extractPackIdFromLibelleDepotCell = extractItem(cell, "^#([0-9]{11}) Paiement de dépot")
End Function

'Formate et traite les données issues des copy/paste des listes de paiements en vue de leur
'importation dans Commence
Sub handlePaiementsInit_TODEL()
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
    
    formatDateAndTime "DATE_OP"
    
    transformMontant "MONTANT_GAIN_COL"
    
    Set rngLibelle = Range("LIBELLE")
    
    compteSubjectToPaiementCol = Range("COMPTE_RECEIVING_GAIN").Column
    packIdCol = Range("PACK_ID").Column
    typePaiementCol = Range("TYPE_GAIN").Column
    idPaiementCol = Range("ID_GAIN").Column
    matriceLevelCol = Range("MATRICE_LEVEL").Column
    pseudoFilleulCol = Range("PSEUDO_FILLEUL").Column
    datePaiementCol = Range("DATE_GAIN_COL").Column
    
    Set lookupTablesSheet = Sheets("Lookup tables")
    
    lastCellRow = getLastDataRow(lookupTablesSheet.Range("A:A"))
    Set lookupRangePackContrat = lookupTablesSheet.Range(lookupTablesSheet.Cells(2, 1), lookupTablesSheet.Cells(lastCellRow, 2))
    
    lastCellRow = getLastDataRow(lookupTablesSheet.Range("D:D"))
    Set lookupRangeContratPseudo = lookupTablesSheet.Range(lookupTablesSheet.Cells(2, 5), lookupTablesSheet.Cells(lastCellRow, 6))

    'clear col 6 à 10 qui contiennent les valeurs extraites par la suite de la macro
    lastCellRow = getLastDataRow(ActiveSheet.Range("A:A"))
    Set gainSheetCalculatedCellsRange = ActiveSheet.Range(ActiveSheet.Cells(2, 6), ActiveSheet.Cells(lastCellRow, 10))
    gainSheetCalculatedCellsRange.Clear
    
    For Each cell In rngLibelle
        If (cell.Value = "") Then
            Exit For
        End If
        
        curRow = cell.Row
        
        packId = extractPackIdFromLibelleCell(cell)
        
        If (packId <> "") Then
            'gain de type 8 % sur achat de pack par un filleul du détenteur du compte
            Cells(curRow, packIdCol).Value = packId
            Cells(curRow, idPaiementCol).Value = packId & "-b"
            Cells(curRow, typePaiementCol).Value = BONUS_ACHAT_PACK_PAR_FILLEUL
            formatPseudoFilleulForPackId packId, curRow, pseudoFilleulCol, lookupRangePackContrat, lookupRangeContratPseudo
        Else
            packId = extractPackIdFromPaiementPackLibelle(cell)
            If (packId <> "") Then
                'gain de 25 % rapporté par un packs du compte
                gainPackMonth = extractPackMonthFromPaiementPackLibelle(cell)
                Cells(curRow, packIdCol).Value = packId
                Cells(curRow, idPaiementCol).Value = packId & "-" & gainPackMonth
                Cells(curRow, typePaiementCol).Value = GAIN_PACK_25_PCT
            Else
                pseudoFilleul = extractPseudoFilleulMatrixPrem(cell)
                If (pseudoFilleul <> "") Then
                    'bonus mensuel comptabilisé dans la matrice Premium
                    Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                    Cells(curRow, typePaiementCol).Value = BONUS_FILLEUL_MATRICE_PREMIUM
                    Cells(curRow, idPaiementCol).Value = pseudoFilleul & "-BMP-to-" & Cells(curRow, compteSubjectToPaiementCol).Value & "-" & Format(Cells(curRow, datePaiementCol).Value2, "dd.mm.yy")
                    matriceLevel = extractMatriceLevelMatrixPrem(cell)
                    Cells(curRow, matriceLevelCol).Value = matriceLevel
                Else
                    pseudoFilleul = extractPseudoFilleulMatrixSE(cell)
                    If (pseudoFilleul <> "") Then
                        'bonus mensuel comptabilisé dans la matrice Super Elite
                        Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                        Cells(curRow, typePaiementCol).Value = BONUS_FILLEUL_MATRICE_SE
                        Cells(curRow, idPaiementCol).Value = pseudoFilleul & "-BSE-to-" & Cells(curRow, compteSubjectToPaiementCol).Value & "-" & Format(Cells(curRow, datePaiementCol).Value2, "dd.mm.yy")
                        matriceLevel = extractMatriceLevelMatrixSE(cell)
                        Cells(curRow, matriceLevelCol).Value = matriceLevel
                    Else
                        pseudoFilleul = extractFilleulUpgrToPremium(cell)
                        If (pseudoFilleul <> "") Then
                            'bonus provenant de l'activation ou de l'upgrade en Premium d'un filleul du détenteur du compte
                            Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                            Cells(curRow, typePaiementCol).Value = BONUS_FILLEUL_UPGR_PREMIUM
                            Cells(curRow, idPaiementCol).Value = pseudoFilleul & "-UPGR_PREM-" & Format(Cells(curRow, datePaiementCol).Value2, "dd.mm.yy")
                        Else
                            pseudoFilleul = extractFilleulUpgrToSE(cell)
                            If (pseudoFilleul <> "") Then
                                'bonus provenant de l'upgrade en Super Elite d'un filleul du détenteur du compte
                                Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                                Cells(curRow, typePaiementCol).Value = BONUS_FILLEUL_UPGR_SE
                                Cells(curRow, idPaiementCol).Value = pseudoFilleul & "-UPGR_SE-" & Format(Cells(curRow, datePaiementCol).Value2, "dd.mm.yy")
                            Else
                                MsgBox "Libellé de gain inconnu dans cellule " & cell.Address
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

