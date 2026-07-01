Sub Bouton1_Cliquer()
    
    ' --- CONSTANTES ---
    Const lineToStartCheckaxe As Integer = 2
    
    Const colTableParam As String = "A"
    Const colTableValue As String = "B"
    Const colPercentAxe As String = "E"
    Const colTotalEntries As String = "D"
    Const colTotalEuros As String = "C"
    
    Const colEurosTarifsFile As String = "F"
    Const colEntriesTarifsFile As String = "D"
    Const colPercentTarifsFile As String = "J"
    
    Const lineTotalAxe As String = "2"
    
    Const colDateComptaFile As String = "A"
    Const colJournalComptaFile As String = "B"
    Const colCompteComptaFile As String = "C"
    Const colPieceComptaFile As String = "D"
    Const colLabelComptaFile As String = "E"
    Const colDebitComptaFile As String = "F"
    Const colCreditComptaFile As String = "G"
    ' --- CONSTANTES ---

    Dim lineToStartComptaFile As Integer
    Dim wsTarifs As Worksheet
    Dim plageLabel As Range
    Dim lstAxe As Variant
    Dim wsCalculs As Worksheet
    Dim plageAxe As Range
    Dim j As Long ' Compteur pour repérer la ligne de l'axe
    Dim sommeTotale As Double
    Dim sommeEntrees As Integer
    Dim dictTable As Object
    
    Dim texteLabel As String
    Dim axeTrouve As Boolean
    
    Dim cle As Variant
    Dim celluleSource As Range
    Dim texteSource As String
    Dim plageComptaFileToClean As Range
    Dim axe As Variant
    
    lineToStartComptaFile = 18
    
' =========================================================================
' SCRIPT
' =========================================================================
    Set wsCalculs = Get_sheet("calculs_axes")
    'je nettoie les résultats du calcul précédent
    wsCalculs.Range(colTotalEuros & lineTotalAxe & ":" & colPercentAxe & CStr(lineToStartComptaFile - 2)).ClearContents
    
    Set plageAxe = Plage_to_check(wsCalculs, lineToStartCheckaxe, colTableValue)
    lstAxe = Application.Transpose(plageAxe.Value)
    
    Set wsTarifs = Get_sheet("fichier_tarifs")
    Set plageLabel = Plage_to_check(wsTarifs, 1, colTableParam)
    
    ' 2. On boucle sur chaque cellule de l'autre plage à vérifier
    For Each celluleSource In plageLabel
        texteSource = Trim(CStr(celluleSource.Value)) ' On nettoie les espaces
        j = 1
        axeTrouve = False
        
        ' 3. On compare avec chaque élément de notre liste de labels
        For Each axe In lstAxe
            texteLabel = Trim(CStr(axe))
            
            ' Vérification : est-ce que le texte se termine par le label ?
            If texteSource Like "*" & texteLabel Then
                axeTrouve = True
            
                Dim totalEuro As Variant
                Dim totalEntrees As Variant
                Dim totalPercent As Variant
                
                totalEuro = wsTarifs.Cells(celluleSource.Row, colEurosTarifsFile).Value
                totalEntrees = wsTarifs.Cells(celluleSource.Row, colEntriesTarifsFile).Value
                totalPercent = wsTarifs.Cells(celluleSource.Row, colPercentTarifsFile).Value

                wsCalculs.Cells(j + 1, colTotalEuros).Value = wsCalculs.Cells(j + 1, colTotalEuros).Value + totalEuro
                wsCalculs.Cells(j + 1, colTotalEntries).Value = wsCalculs.Cells(j + 1, colTotalEntries).Value + totalEntrees
                wsCalculs.Cells(j + 1, colPercentAxe).Value = wsCalculs.Cells(j + 1, colPercentAxe).Value + totalPercent
                
                Exit For ' On a trouvé, pas la peine de tester les autres labels pour cette ligne
            End If
            j = j + 1
        Next axe
        
        If axeTrouve = False Then
            MsgBox "Aucun axe correspondant n'a été trouvé pour la valeur : '" & texteSource & "'", _
                   vbExclamation, "Axe introuvable"
            End
        End If
        
    Next celluleSource
    
    
    ' je modifie les valeurs de la colonne E de l'onglet "calculs_axes" en pourcentage
    wsCalculs.Columns(colPercentAxe).NumberFormat = "0.0%"
    derniereLigneEntries = wsCalculs.Cells(3, colTotalEntries).End(xlDown).Row
    sommeTotale = Application.WorksheetFunction.Sum(wsCalculs.Range(colTotalEuros & "2:" & colTotalEuros & derniereLigneEntries))
    wsCalculs.Range(colTotalEuros & lineTotalAxe).Value = sommeTotale
    
    sommeEntrees = Application.WorksheetFunction.Sum(wsCalculs.Columns(colTotalEntries))
    wsCalculs.Range(colTotalEntries & lineTotalAxe).Value = sommeEntrees
    
    ' valorisation comptaFile
    Set plageComptaFileToClean = Plage_to_check(wsCalculs, lineToStartComptaFile, colCreditComptaFile)
    
    Dim vraieDerniereLigne As Long

    ' On prend la ligne de départ de la plage + son nombre de lignes - 1
    vraieDerniereLigne = plageComptaFileToClean.Rows(plageComptaFileToClean.Rows.Count).Row
    
    wsCalculs.Range(colDateComptaFile & lineToStartComptaFile & ":" & colCreditComptaFile & vraieDerniereLigne).ClearContents
    
    ' récupératoin dictionnaire de valeurs
    Set dictTable = GetDictTable(wsCalculs, lineToStartCheckaxe, colTableParam)
    
    If dictTable("Date")("value") <> "" Then
        dictTable("Date")("value") = Date
    End If
    
    For Each cle In dictTable.Keys
        If cle Like "706*" Then
            wsCalculs.Range(colDateComptaFile & lineToStartComptaFile).Value = dictTable("Date")("value")
            wsCalculs.Range(colJournalComptaFile & lineToStartComptaFile).Value = dictTable("Journal")("value")
            wsCalculs.Range(colCompteComptaFile & lineToStartComptaFile).Value = CStr(cle)
            wsCalculs.Range(colPieceComptaFile & lineToStartComptaFile).Value = dictTable("Piece")("value") & dictTable("Date")("value")
            wsCalculs.Range(colLabelComptaFile & lineToStartComptaFile).Value = dictTable(cle)("value") & "-" & dictTable("Date")("value")
            
            dictTable(cle)("total") = Trim(Replace(dictTable(cle)("total"), "€", ""))
            
            If Mid(cle, 6, 1) = "0" Then
                wsCalculs.Range(colDebitComptaFile & lineToStartComptaFile).NumberFormat = "0.00"
                wsCalculs.Range(colDebitComptaFile & lineToStartComptaFile).Value = dictTable(cle)("total")
                wsCalculs.Range(colCreditComptaFile & lineToStartComptaFile).NumberFormat = "0.00"
                wsCalculs.Range(colCreditComptaFile & lineToStartComptaFile).Value = 0
            Else
                wsCalculs.Range(colCreditComptaFile & lineToStartComptaFile).NumberFormat = "0.00"
                wsCalculs.Range(colCreditComptaFile & lineToStartComptaFile).Value = dictTable(cle)("total")
                wsCalculs.Range(colDebitComptaFile & lineToStartComptaFile).NumberFormat = "0.00"
                wsCalculs.Range(colDebitComptaFile & lineToStartComptaFile).Value = 0
            End If
            lineToStartComptaFile = lineToStartComptaFile + 1
        End If
    Next cle
    
    MsgBox "Génération terminée.", vbCritical
End Sub

' =========================================================================
' LES FONCTIONS
' =========================================================================

Function Get_sheet(table_name As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Sheets(table_name)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "L'onglet '" & table_name & "' est introuvable.", vbCritical
        End
    End If
    Set Get_sheet = ws
End Function


Function Plage_to_check(sheet_to_use As Worksheet, start_row As Integer, LettreColonne As String) As Range
    Dim DerniereLigne As Long

    DerniereLigne = sheet_to_use.Cells(sheet_to_use.Rows.Count, LettreColonne).End(xlUp).Row
        
    ' On attribue l'objet final au nom de la fonction avec un SET
    Set Plage_to_check = sheet_to_use.Range(LettreColonne & start_row & ":" & LettreColonne & DerniereLigne)

End Function


Function GetDictTable(wsCalculs As Worksheet, lineToStartCheckaxe As Integer, colTableParam As String) As Object
    Dim MonDico As Object
    Dim DictVal As Object

    Set MonDico = CreateObject("Scripting.Dictionary")
    
    Do While wsCalculs.Cells(lineToStartCheckaxe, colTableParam).Value <> ""
        Set DictVal = CreateObject("Scripting.Dictionary")
        
        DictVal("value") = Trim(wsCalculs.Cells(lineToStartCheckaxe, 2).Value)
        DictVal("total") = wsCalculs.Cells(lineToStartCheckaxe, 3).Value
        
        cle = wsCalculs.Cells(lineToStartCheckaxe, 1).Value
        Set MonDico(Trim(cle)) = DictVal
        lineToStartCheckaxe = lineToStartCheckaxe + 1
    Loop
    
    Set GetDictTable = MonDico
End Function







