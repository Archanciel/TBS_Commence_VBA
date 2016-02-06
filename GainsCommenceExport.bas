Attribute VB_Name = "GainsCommenceExport"
Option Explicit
Private Const GAIN_TYPE_BONUS_ACHAT_PACK_PAR_FILLEUL As String = "Bonus achat pack par filleul"
Private Const GAIN_TYPE_GAIN_PACK_25_PCT As String = "Gain pack 25 %"
Private Const GAIN_TYPE_GAIN_PACK_28_PCT As String = "Gain pack 28 %"
Private Const GAIN_TYPE_GAIN_PACK_UNKNOWN As String = "### Gain pack inconnu ###"
Private Const GAIN_TYPE_BONUS_FILLEUL_MATRICE_PREMIUM As String = "Bonus matrice Premium"
Private Const GAIN_TYPE_BONUS_FILLEUL_MATRICE_SE As String = "Bonus matrice SE"
Private Const GAIN_TYPE_BONUS_FILLEUL_UPGR_PREMIUM As String = "Bonus filleul upgr Premium"
Private Const GAIN_TYPE_BONUS_FILLEUL_UPGR_SE As String = "Bonus filleul upgr SE"
Private Const GAIN_IMPORT_FLAG_TRUE As String = "1"
Private Const GAIN_VERIFIED_FLAG_TRUE As String = "1"

'Exporte les données de la feuille Gains dans un fichier texte tab separated pouvant être importé dans Commence
Sub gainsExportDataForCommence()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    deleteNomComptes "NOM_COMPTES_G"
    deleteTopRow
    saveSheetAsTabDelimTxtFileTimeStamped ActiveSheet.Name
    Application.ScreenUpdating = True
    closeWithoutSave
End Sub

'Formate et traite les données issues des copy/paste des listes de gains en vue de leur
'importation dans Commence
Sub handleRevenues()
    Dim rngLibelle As Range
    Dim cell As Range
    Dim packId As String
    Dim tauxGain As Integer
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
    Dim noGainCol As Long
    Dim montantGainCol As Long
    Dim windowsWideThousandSeparator As String
    Dim importFlagCol As Long
    Dim verifiedFlagCol As Long
    
    Application.ScreenUpdating = False
    
    terminateIfNoData Cells(2, Range("LIBELLE").Column)
    
    windowsWideThousandSeparator = Application.International(xlThousandsSeparator)
    
    formatDateAndTime "DATE_GAIN_COL", "TIME_GAIN"
    transformMontant "MONTANT_GAIN_COL"
    
    Set rngLibelle = Range("LIBELLE")
    
    compteReceivingGainCol = Range("COMPTE_RECEIVING_GAIN").Column
    packIdCol = Range("PACK_ID").Column
    typeGainCol = Range("TYPE_GAIN").Column
    idGainCol = Range("ID_GAIN").Column
    matriceLevelCol = Range("MATRICE_LEVEL").Column
    pseudoFilleulCol = Range("PSEUDO_FILLEUL").Column
    dateGainCol = Range("DATE_GAIN_COL").Column
    noGainCol = Range("NO_GAIN").Column
    montantGainCol = Range("MONTANT_GAIN").Column
    importFlagCol = Range("GAIN_IMPORT").Column
    verifiedFlagCol = Range("GAIN_VERIFIED").Column
    
    'initialize lookup table references
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
        
        Cells(curRow, importFlagCol).Value = GAIN_IMPORT_FLAG_TRUE
        Cells(curRow, verifiedFlagCol).Value = GAIN_VERIFIED_FLAG_TRUE
        packId = extractPackIdFromBonusLibelle(cell)
        
        If (packId <> "") Then
            'gain de type 8 % sur achat de pack par un filleul du détenteur du compte
            Cells(curRow, packIdCol).Value = packId
            Cells(curRow, idGainCol).Value = packId & "-b"
            Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_ACHAT_PACK_PAR_FILLEUL
            formatPseudoFilleulForPackId packId, curRow, pseudoFilleulCol, lookupRangePackContrat, lookupRangeContratPseudo
        Else
            packId = extractPackIdFromGainPackLibelle(cell)
            If (packId <> "") Then
                'revenu de pack de 25 ou 28 %
                tauxGain = extractTauxGainFromGainPackLibelle(cell)
                If (tauxGain = 25) Then
                    Cells(curRow, typeGainCol).Value = GAIN_TYPE_GAIN_PACK_25_PCT
                ElseIf (tauxGain = 28) Then
                    Cells(curRow, typeGainCol).Value = GAIN_TYPE_GAIN_PACK_28_PCT
                Else
                    Cells(curRow, typeGainCol).Value = GAIN_TYPE_GAIN_PACK_UNKNOWN
                End If
                gainPackMonth = extractPackMonthFromGainPackLibelle(cell)
                Cells(curRow, packIdCol).Value = packId
                Cells(curRow, idGainCol).Value = packId & "-" & gainPackMonth
                Cells(curRow, noGainCol).Value = gainPackMonth
            Else
                pseudoFilleul = extractPseudoFilleulMatrixPrem(cell)
                If (pseudoFilleul <> "") Then
                    'bonus mensuel comptabilisé dans la matrice Premium
                    Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                    Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_FILLEUL_MATRICE_PREMIUM
                    Cells(curRow, idGainCol).Value = pseudoFilleul & "-BMP-to-" & Cells(curRow, compteReceivingGainCol).Value & "-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                    matriceLevel = extractMatriceLevelMatrixPrem(cell)
                    Cells(curRow, matriceLevelCol).Value = matriceLevel
                Else
                    pseudoFilleul = extractPseudoFilleulMatrixSE(cell)
                    If (pseudoFilleul <> "") Then
                        'bonus mensuel comptabilisé dans la matrice Super Elite
                        Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                        Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_FILLEUL_MATRICE_SE
                        Cells(curRow, idGainCol).Value = pseudoFilleul & "-BSE-to-" & Cells(curRow, compteReceivingGainCol).Value & "-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                        matriceLevel = extractMatriceLevelMatrixSE(cell)
                        Cells(curRow, matriceLevelCol).Value = matriceLevel
                    Else
                        pseudoFilleul = extractFilleulUpgrToPremium(cell)
                        If (pseudoFilleul <> "") Then
                            'bonus provenant de l'activation ou de l'upgrade en Premium d'un filleul du détenteur du compte
                            Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                            Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_FILLEUL_UPGR_PREMIUM
                            Cells(curRow, idGainCol).Value = pseudoFilleul & "-UPGR_PREM-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                        Else
                            pseudoFilleul = extractFilleulUpgrToSE(cell)
                            If (pseudoFilleul <> "") Then
                                'bonus provenant de l'upgrade en Super Elite d'un filleul du détenteur du compte
                                Cells(curRow, pseudoFilleulCol).Value = pseudoFilleul
                                Cells(curRow, typeGainCol).Value = GAIN_TYPE_BONUS_FILLEUL_UPGR_SE
                                Cells(curRow, idGainCol).Value = pseudoFilleul & "-UPGR_SE-" & Format(Cells(curRow, dateGainCol).Value2, "dd.mm.yy")
                            Else
                                Cells(curRow, typeGainCol).Value = "### LIBELLE DE GAIN INCONNU ###"
                                MsgBox "Libellé de gain inconnu dans cellule " & cell.Address & " !", vbInformation
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next cell
    
    formatIdCol "PACK_ID"
    clearAnySelection
    
    Application.ScreenUpdating = True
End Sub
Private Sub formatPseudoFilleulForPackId(packId As String, curRow As Long, pseudoFilleulCol As Long, lookupRangePackContrat As Range, lookupRangeContratPseudo As Range)
    Dim nomContratCommence As Variant
    Dim pseudoTBS As Variant
    Dim pseudoFilleul As String
    
    nomContratCommence = Application.VLookup(CDbl(packId), lookupRangePackContrat, 2, False)
    
    If (IsError(nomContratCommence)) Then
        pseudoTBS = "### pack id '" + packId + "' not found in lookup table ###"
    Else
        pseudoTBS = Application.VLookup(nomContratCommence, lookupRangeContratPseudo, 2, False)
        If (IsError(pseudoTBS)) Then
            pseudoTBS = "### contrat Commence '" + nomContratCommence + "' not found in lookup table ###"
        End If
    End If
    
    Cells(curRow, pseudoFilleulCol).Value = pseudoTBS
End Sub
'Extrait du libellé d'annonce de gain le numéro de pack dont l'achat par un filleul
'a généré le bonus.
'
'Exemple de libellé: Bonus sponsor pour dépot(#13441058360)
Private Function extractPackIdFromBonusLibelle(cell As Range) As String
    extractPackIdFromBonusLibelle = extractItem(cell, "dépot\(#([0-9]+)\)$")
End Function

'Extrait du libellé d'annonce de gain de pack le numéro de pack  qui a généré le gain.
'
'Exemple de libellé: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractPackIdFromGainPackLibelle(cell As Range) As String
    extractPackIdFromGainPackLibelle = extractItem(cell, "^#([0-9]+)")
End Function

'Extrait du libellé d'annonce de gain de pack le numéro du mois du gain.
'
'Exemple de libellé: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractPackMonthFromGainPackLibelle(cell As Range) As String
    extractPackMonthFromGainPackLibelle = extractItem(cell, "([0-9]+)/12\]$")
End Function

'Extrait du libellé d'annonce de gain de pack le numéro du mois du gain.
'
'Exemple de libellé: #12934041280-> Profit, 25.00% of 10000.00 deposited [1/12]
Private Function extractTauxGainFromGainPackLibelle(cell As Range) As String
    extractTauxGainFromGainPackLibelle = extractItem(cell, "Profit, ([0-9]+)\.")
End Function

'Extrait du libellé d'annonce de bonus matrice Premium le pseudo du filleul.
'
'Exemple de libellé: Niveau réseau Premium#1 bonus (tamcerise)  ou
'                    VIP Network level#1 bonus (lucky70)
Private Function extractPseudoFilleulMatrixPrem(cell As Range) As String
    Dim strPseudo As String
    
    strPseudo = extractItem(cell, "^Niveau réseau Premium#\d* bonus \(([a-zA-Z0-9-_]+)\)")
    
    If (strPseudo = "") Then
        'essai avec la version anglaise du libellé
        strPseudo = extractItem(cell, "^VIP Network level#\d* bonus \(([a-zA-Z0-9-_]+)\)")
    End If
    
    extractPseudoFilleulMatrixPrem = strPseudo
End Function

'Extrait du libellé d'annonce de bonus matrice Super Elite le pseudo du filleul.
'
'Exemple de libellé: SVIP level#1 bonus (jpensuisse)    >>> old !
'                    SVIP Network level#1 bonus (rosemaman)
Private Function extractPseudoFilleulMatrixSE(cell As Range) As String
    Dim strPseudo As String
    
    strPseudo = extractItem(cell, "^SVIP Network level#\d* bonus \(([a-zA-Z0-9-_]+)\)")
    
    If (strPseudo = "") Then
        'essai avec la version level# du libellé
        strPseudo = extractItem(cell, "^SVIP level#\d* bonus \(([a-zA-Z0-9-_]+)\)")
    End If
    
'    If (strPseudo = "") Then
'        'essai avec la version française du libellé
'        strPseudo = extractItem(cell, "^VIP Network level#\d* bonus \(([a-zA-Z0-9-_]+)\)")
'    End If

    extractPseudoFilleulMatrixSE = strPseudo
End Function

'Extrait du libellé d'annonce de bonus matrice premium le niveau matriciel du gain.
'
'Exemple de libellé: Niveau réseau Premium#1 bonus (tamcerise)
'                    VIP Network level#1 bonus (lucky70)
Private Function extractMatriceLevelMatrixPrem(cell As Range) As String
    Dim strLevel As String
    
    strLevel = extractItem(cell, "^Niveau réseau Premium#(\d*) bonus")
    
    If (strLevel = "") Then
        'essai avec la version anglaise du libellé
        strLevel = extractItem(cell, "^VIP Network level#(\d*) bonus")
    End If
    
    extractMatriceLevelMatrixPrem = strLevel
End Function

'Extrait du libellé d'annonce de bonus matrice premium le niveau matriciel du gain.
'
'WARNING: TU NE CONNAIS PAS ENCORE AVEC CERTITUDE LES LIBELLES EXACTS !
'Exemple de libellé: Niveau réseau Super Elite#1 bonus (tamcerise)
'                    SVIP level#1 bonus (lucky70)   >>> old !
'                    SVIP Network level#1 bonus (rosemaman)
Private Function extractMatriceLevelMatrixSE(cell As Range) As String
    Dim strLevel As String
    
    strLevel = extractItem(cell, "^Niveau réseau Super Elite#(\d*) bonus")
    
    If (strLevel = "") Then
        'essai avec la version anglaise du libellé
        strLevel = extractItem(cell, "^SVIP Network level#(\d*) bonus")
    End If
    
    If (strLevel = "") Then
        'essai avec la version SVIP level# du libellé
        strLevel = extractItem(cell, "^SVIP level#(\d*) bonus")
    End If
    
    extractMatriceLevelMatrixSE = strLevel
End Function

'Extrait le pseudo du filleul du libellé d'annonce de bonus en cas d'upgrade de celui-ci à Premium.
'
'Exemple de libellé: Bonus sponsor (rosemaman)
Private Function extractFilleulUpgrToPremium(cell As Range) As String
    extractFilleulUpgrToPremium = extractItem(cell, "^Bonus sponsor \(([a-zA-Z0-9-_]+)\)")
End Function

'Extrait le pseudo du filleul du libellé d'annonce de bonus en cas d'upgrade de celui-ci à Super Elite.
'
'Exemple de libellé: SVIP Sponsor bonus (jpensuisse)
Private Function extractFilleulUpgrToSE(cell As Range) As String
    extractFilleulUpgrToSE = extractItem(cell, "^SVIP Sponsor bonus \(([a-zA-Z0-9-_]+)\)")
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
