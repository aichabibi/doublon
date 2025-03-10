Sub Verifier_Doublons()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim matricule As String, dateDebut As String, activite As String
    Dim dict As Object
    Dim msg As String
    Dim nom As String, prenom As String
    Dim results As Object
    Dim key As Variant
    Dim colMatricule As Integer, colDateDebut As Integer
    Dim colNom As Integer, colPrenom As Integer, colActivite As Integer, colCumul As Integer
    Dim found As Boolean
    Dim activiteValides As Variant

    ' Liste des valeurs valides pour la colonne ACTIVITE '
    activiteValides = Array("IGD HORS IDF 1 REP.", "IGD HORS IDF 2 REP.", "IGD HORS IDF LOG. + 1 REP.", _
                            "IGD HORS IDF LOG. + 2 REP.", "IGD IDF 1 REP.", "IGD IDF 2 REP.", _
                            "IGD IDF LOG. + 1 REP.", "IGD IDF LOG. + 2 REP.", "IPD Repas hors locaux (TX)", _
                            "Repas pris restaurant", "IPD Ticket restaurant", "Panier Sedentaire (TX)")

    ' Sélectionner la feuille '
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Synthese CRA 32D1")
    On Error GoTo 0
    
    ' Vérifier si la feuille existe '
    If ws Is Nothing Then
        MsgBox "La feuille 'Synthese CRA 32D1' n'existe pas !", vbCritical, "Erreur"
        Exit Sub
    End If
    
    ' Trouver la dernière ligne de données '
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Colonne A pour s'assurer que toutes les lignes sont prises

    ' Trouver les colonnes dynamiquement '
    found = False
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case Trim(ws.Cells(1, i).Value)
            Case "MATRICULE": colMatricule = i: found = True
            Case "DATE DEBUT": colDateDebut = i: found = True
            Case "NOM": colNom = i
            Case "PRENOM": colPrenom = i
            Case "ACTIVITE": colActivite = i ' Trouver la colonne ACTIVITE '
            Case "CUMUL": colCumul = i ' Trouver la colonne CUMUL '
        End Select
    Next i

    ' Vérifier si les colonnes essentielles sont trouvées '
    If colMatricule = 0 Or colDateDebut = 0 Or colActivite = 0 Or colCumul = 0 Then
        MsgBox "Impossible de trouver les colonnes 'MATRICULE', 'DATE DEBUT', 'ACTIVITE' ou 'CUMUL'. Vérifiez les noms des colonnes.", vbCritical, "Erreur"
        Exit Sub
    End If
    
    ' Initialiser le dictionnaire et la liste des résultats '
    Set dict = CreateObject("Scripting.Dictionary")
    Set results = CreateObject("System.Collections.ArrayList")
    
    ' Parcourir les lignes pour stocker les occurrences '
    For i = 2 To lastRow
        matricule = ws.Cells(i, colMatricule).Value
        dateDebut = ws.Cells(i, colDateDebut).Value
        activite = ws.Cells(i, colActivite).Value
        cumul = ws.Cells(i, colCumul).Value
        
        ' Vérifier si l'activité est valide (si elle est dans la liste des activités valides) et si CUMUL n'est pas égal à 0 '
        If IsInArray(activite, activiteValides) And cumul <> 0 Then
            ' Clé unique = combinaison de DATE DEBUT + MATRICULE + ACTIVITE
            key = dateDebut & "_" & matricule & "_" & activite
            
            ' Ajouter dans le dictionnaire
            If dict.exists(key) Then
                dict(key) = dict(key) + 1
            Else
                dict.Add key, 1
            End If
        End If
    Next i
    
    ' Vérifier les doublons et stocker les résultats '
    For Each key In dict.keys
        If dict(key) > 1 Then ' Filtrer les doublons (plus de 1 occurrence) '
            matricule = Split(key, "_")(1)
            dateDebut = Split(key, "_")(0)
            activite = Split(key, "_")(2)
            
            ' Trouver le NOM et PRÉNOM associés '
            For i = 2 To lastRow
                If ws.Cells(i, colMatricule).Value = matricule And ws.Cells(i, colDateDebut).Value = dateDebut And ws.Cells(i, colActivite).Value = activite Then
                    nom = ws.Cells(i, colNom).Value
                    prenom = ws.Cells(i, colPrenom).Value
                    Exit For
                End If
            Next i
            
            results.Add dateDebut & " | " & matricule & " | " & nom & " " & prenom & " | " & activite
        End If
    Next key
    
    ' Vérifier si des doublons existent '
    If results.Count > 0 Then
        results.Sort ' Trier les résultats pour plus de clarté '
        
        ' Exporter les résultats dans une feuille si trop long '
        If results.Count > 20 Then
            Dim wsNew As Worksheet
            Set wsNew = ThisWorkbook.Sheets.Add
            wsNew.Name = "Doublons Détectés"
            
            wsNew.Cells(1, 1).Value = "DATE DEBUT"
            wsNew.Cells(1, 2).Value = "MATRICULE"
            wsNew.Cells(1, 3).Value = "NOM PRÉNOM"
            wsNew.Cells(1, 4).Value = "ACTIVITE"
            
            For i = 0 To results.Count - 1
                wsNew.Cells(i + 2, 1).Value = Split(results(i), " | ")(0)
                wsNew.Cells(i + 2, 2).Value = Split(results(i), " | ")(1)
                wsNew.Cells(i + 2, 3).Value = Split(results(i), " | ")(2)
                wsNew.Cells(i + 2, 4).Value = Split(results(i), " | ")(3)
            Next i
            
            MsgBox "La liste des doublons a été exportée dans la feuille 'Doublons Détectés'.", vbInformation, "Export Terminé"
        Else
            ' Affichage direct dans une boîte de dialogue '
            Dim output As String
            output = "Liste des doublons détectés :" & vbNewLine & String(50, "-") & vbNewLine
            
            For Each Item In results
                output = output & Item & vbNewLine
            Next Item
            
            MsgBox output, vbExclamation, "Alerte : Matricules en double"
        End If
    Else
        MsgBox "Aucun doublon détecté.", vbInformation, "Vérification Terminée"
    End If
    
    ' Nettoyage '
    Set dict = Nothing
    Set results = Nothing
    Set ws = Nothing
End Sub

' Fonction pour vérifier si un élément est dans un tableau '
Function IsInArray(val As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = val Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function
