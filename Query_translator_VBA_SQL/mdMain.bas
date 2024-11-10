Attribute VB_Name = "mdMain"
Sub CheckAllFond()
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("EXTRACTION")
    Set rng = ws.Range("H22:H29")
    
    If ws.Cells(20, 8).Value = True Then 'Si la case est cochée
        
        'Boucle pour cocher toutes les cases dans la plage spécifiée
        
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then
                If shp.FormControlType = xlCheckBox Then
                    If Not Intersect(shp.TopLeftCell, rng) Is Nothing Then
                        shp.ControlFormat.Value = xlOn
                    End If
                End If
            End If
        Next shp
        
    ElseIf ws.Cells(20, 8).Value = False Then 'Si la case est décochée
        
        'Boucle pour décocher toutes les cases dans la plage spécifiée
        
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then
                If shp.FormControlType = xlCheckBox Then
                    If Not Intersect(shp.TopLeftCell, rng) Is Nothing Then
                        shp.ControlFormat.Value = xlOff
                    End If
                End If
            End If
        Next shp
        
    End If

End Sub

Sub CheckAllChamp()
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("EXTRACTION")
    Set rng = ws.Range("D5:D37")
    
    If ws.Cells(4, 4).Value = True Then 'Si la case est cochée
        
        'Boucle pour cocher toutes les cases dans la plage spécifiée
        
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then
                If shp.FormControlType = xlCheckBox Then
                    If Not Intersect(shp.TopLeftCell, rng) Is Nothing Then
                        shp.ControlFormat.Value = xlOn
                    End If
                End If
            End If
        Next shp
        
    ElseIf ws.Cells(4, 4).Value = False Then 'Si la case est décochée
        
        'Boucle pour décocher toutes les cases dans la plage spécifiée
        
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then
                If shp.FormControlType = xlCheckBox Then
                    If Not Intersect(shp.TopLeftCell, rng) Is Nothing Then
                        shp.ControlFormat.Value = xlOff
                    End If
                End If
            End If
        Next shp
        
    End If

End Sub

Function GetFileSourcePath() As String
    
    GetFileSourcePath = ""
    
    Application.FileDialog(msoFileDialogFilePicker).Filters.Clear
    
    With Application.FileDialog(msoFileDialogFilePicker)
    
        .AllowMultiSelect = False
        
        .Filters.Add "", "*.xlsx;*.xlsm"
        
        If .Show <> 0 Then
        
            GetFileSourcePath = .SelectedItems(1)
        
        End If
    
    End With
    
End Function

Function GetSheetNameFromWorkbook(ByVal filePath As String) As String
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    
    ' Ouvrir le classeur Excel
    Set wb = Workbooks.Open(filePath)
    
    ' Afficher les noms des feuilles et laisser l'utilisateur choisir
    sheetName = InputBox("Veuillez choisir une feuille:", "Choix de la feuille", wb.Sheets(1).Name)
    
    ' Fermer le classeur sans enregistrer les modifications
    wb.Close SaveChanges:=False
    
    ' Retourner le nom de la feuille sélectionnée
    GetSheetNameFromWorkbook = sheetName

End Function

Sub LoadFileSourcePath()
    
    'On Error GoTo 1
    
    Dim FileSourcePath As String 'To store the path of selected file
    Dim FileName As String 'To store the name of the selected file
    Dim SelectedSheet As String 'To store the selected sheet name
    
    FileSourcePath = Trim(GetFileSourcePath())
    'FileName = Split(FileSourcePath, "\")(UBound(Split(FileSourcePath, "\"))) 'The file name
    
    If Not FileSourcePath = "" Then
    
        ThisWorkbook.Sheets("EXTRACTION").Range("E2") = FileSourcePath 'The file path
        
        'Ask user to select a sheet from the selected workbook
        SelectedSheet = GetSheetNameFromWorkbook(FileSourcePath)
        
        If SelectedSheet <> "" Then
            ThisWorkbook.Sheets("EXTRACTION").Range("H5") = SelectedSheet 'Display selected sheet name
        Else
            ThisWorkbook.Sheets("EXTRACTION").Range("H5") = "Aucune feuille choisie"
            ThisWorkbook.Sheets("EXTRACTION").Range("E2") = "" 'The file path
        End If
        
    Else
        ThisWorkbook.Sheets("EXTRACTION").Range("H5") = "Aucune feuille choisie"
        ThisWorkbook.Sheets("EXTRACTION").Range("E2") = "" 'The file path
    End If
    
1
End Sub

Function RemoveFirstString(chaine As String) As String
    
    If Len(chaine) > 0 Then
        RemoveFirstString = Right(chaine, Len(chaine) - 1)
    Else
        RemoveFirstString = ""
    End If
    
End Function

Function champSelected() As String
    
    champSelected = ""
    
    Dim ws As Worksheet
    
    'Définir la feuille de selection
    Set ws = ThisWorkbook.Sheets("EXTRACTION")
    
    'Boucle sur toutes les lignes de la colonne D
    For i = 6 To 37
        
        If ws.Cells(i, 4) = True Then
            
            champSelected = champSelected & "," & ws.Cells(i, 3)

        End If
        
    Next i
    
    If champSelected <> "" Then
        'Nettoyer la première virgule
        champSelected = RemoveFirstString(Trim(champSelected))
    Else
        MsgBox "Vous devrez cocher au moins un champ pour la requête.", vbCritical, "Extraction"
    End If
    
End Function

Function FondSelected() As String
    
    FondSelected = ""
    
    Dim ws As Worksheet
    
    'Définir la feuille de selection
    Set ws = ThisWorkbook.Sheets("EXTRACTION")
    
    'Boucle sur toutes les lignes de la colonne H
    For i = 23 To 29
        
        If ws.Cells(i, 8) = True Then
            
            FondSelected = FondSelected & "," & "'" & ws.Cells(i, 7) & "'"

        End If
        
    Next i
    
    If FondSelected <> "" Then
        'Nettoyer la première virgule
        FondSelected = "(" & RemoveFirstString(Trim(FondSelected)) & ")"
    Else
        FondSelected = ""
    End If
    
    
    'MsgBox FondSelected
    
End Function


Function ConditionSelected() As String
    
    Dim ws As Worksheet
    Dim condition As String
    Dim finalCondition As String
    
    'Définir la feuille de sélection
    Set ws = ThisWorkbook.Sheets("EXTRACTION")
    
    'Initialiser la condition finale
    finalCondition = ""
    
    'Boucle sur toutes les lignes de la colonne G
    For i = 14 To 18
        
        If Not ws.Cells(i, 7).Value = "" And ws.Cells(i, 8).Value <> "" And ws.Cells(i, 9).Value <> "" Then
            
            If IsNumeric(ws.Cells(i, 9)) Then
                condition = Trim(ws.Cells(i, 7).Value & " " & ws.Cells(i, 8).Value & " " & ws.Cells(i, 9).Value)
            Else
                condition = Trim(ws.Cells(i, 7).Value & " " & ws.Cells(i, 8).Value & " " & "'" & ws.Cells(i, 9).Value) & "'"
            End If
            
            'Ajouter la condition à la condition finale
            If finalCondition <> "" Then
                finalCondition = finalCondition & " AND " & condition
            Else
                finalCondition = condition
            End If
            
        End If
        
    Next i
    
    'Retourner la condition finale
    ConditionSelected = finalCondition
    
    'MsgBox ConditionSelected
    
End Function

Function orderBy() As String

    Dim ws As Worksheet
    Dim orderSelected As String
    
    Set ws = ThisWorkbook.Sheets("EXTRACTION")
    
    orderSelected = ""
    
    'Boucle sur la colonne H and I
    For i = 6 To 7
        If ws.Cells(i, 8) <> "" And ws.Cells(i, 9) <> "" Then
            If ws.Cells(i, 9) = "Ascending" Then
                orderSelected = orderSelected & ", " & ws.Cells(i, 8) & " " & "ASC"
            Else
                orderSelected = orderSelected & ", " & ws.Cells(i, 8) & " " & "DESC"
            End If
        End If
    Next i
    
    orderBy = RemoveFirstString(Trim(orderSelected))
    'MsgBox orderBy

End Function

Function FileDestination() As String
    
    'Function de création du ficher de sortie
    On Error GoTo 1
    
    Dim ws As Worksheet
    Dim nouveauClasseur As Workbook
    Dim sheetDestination As Worksheet
    Dim Jour As String
    Dim Heure As String
    Dim FileName As String
    Dim pathFolder As String
    
    'Obtenir la date et l'eure actuelle
    Jour = Format(Date, "dd-mm-yyyy")
    Heure = Format(Now, "hh-mm-ss")
    
    'Définir le sheet source
    Set ws = ThisWorkbook.Sheets("EXTRACTION")
    
    If ws.Range("H8") = "CSV file" Then
    
        'Définir le nom du fichier
        FileName = Jour & " " & Heure & ".csv"
        
        'Définir le fichier destination en fonction du format choisi
        pathFolder = ThisWorkbook.Path & Application.PathSeparator & FileName 'Chemin du classeur
    Else
        'Définir le nom du fichier
        FileName = Jour & " " & Heure & ".xlsx"
        
        'Définir le fichier destination en fonction du format choisi
        pathFolder = ThisWorkbook.Path & Application.PathSeparator & FileName 'Chemin du classeur
    End If
    
    'Créer et enregistre un nouveau classeur
    Set nouveauClasseur = Workbooks.Add
    nouveauClasseur.SaveAs pathFolder
    
    'Renommer maintenant la feuille copiée
    Set sheetDestination = nouveauClasseur.Sheets(nouveauClasseur.Sheets.Count)
    sheetDestination.Name = "Output"
    
    FileDestination = pathFolder
    
1
    
End Function


Sub Extraction()
    
    On Error GoTo ErrorHandler
    
    'Declaring variables
    Dim wThis As Worksheet
    Dim wDestination As Workbook
    Dim FileDestinationPath As String
    Dim FileSource As String
    Dim sheetName As String
    Dim sheetIndex As String
    Dim connection As New ADODB.connection
    Dim query As String
    Dim rs As New ADODB.Recordset
    Dim champToFetch As String
    Dim fondToFetch As String
    Dim conditionToFetch As String
    Dim orderField As String
    
    'Definir wsThis
    Set wThis = ThisWorkbook.Sheets("EXTRACTION")
    
    'Définir les information de la feuille source
    FileSource = ThisWorkbook.Sheets("EXTRACTION").Range("E2")
    sheetName = ThisWorkbook.Sheets("EXTRACTION").Range("H5").Value
    
    'Se connecter
    connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileSource & ";Extended Properties='Excel 12.0 XML;HDR=YES';"
    
    'Obtenir les champs et conditions sélectionnés
    champToFetch = Trim(champSelected)
    fondToFetch = Trim(FondSelected)
    conditionToFetch = Trim(ConditionSelected)
    orderField = Trim(orderBy)
    'MsgBox orderField

    'Créer la requête
    If champToFetch <> "" And fondToFetch = "" Then
    
        '1er cas
        If orderField = "" Then
            query = "SELECT " & champToFetch & " FROM [" & sheetName & "$]"
        Else
            query = "SELECT " & champToFetch & _
            " FROM [" & sheetName & "$]" & _
            " ORDER BY " & orderField
        End If
        
    ElseIf champToFetch <> "" And fondToFetch <> "" And conditionToFetch = "" Then
    
        '2eme cas
        If orderField = "" Then
            query = "SELECT " & champToFetch & " FROM [" & sheetName & "$] WHERE bps_libellefonds IN " & FondSelected
        Else
            query = "SELECT " & champToFetch & _
            " FROM [" & sheetName & "$]" & _
            " WHERE bps_libellefonds IN " & FondSelected & _
            " ORDER BY " & orderField
        End If
    
    ElseIf champToFetch <> "" And fondToFetch <> "" And conditionToFetch <> "" Then
        
        '3eme cas
        If orderField = "" Then
            query = "SELECT " & champToFetch & _
            " FROM [" & sheetName & "$]" & _
            " WHERE bps_libellefonds IN " & fondToFetch & _
            " AND " & conditionToFetch
        Else
            query = "SELECT " & champToFetch & _
            " FROM [" & sheetName & "$]" & _
            " WHERE bps_libellefonds IN " & FondSelected & _
            " AND " & ConditionSelected & _
            " ORDER BY " & orderField
        End If
    
    End If
    
    'Lancer la requête
    rs.Open query, connection
    
    If Not rs.EOF Then
        
        'Créer le classeur de destination
        FileDestinationPath = Trim(FileDestination)
        
        'Ouvrir le classeur de destination
        Set wDestination = Workbooks.Open(FileDestinationPath)
        
        'Ecrire l'en-tête du tableau
        Dim i As Long
        For i = 0 To rs.Fields.Count - 1
            wDestination.Sheets("Output").Cells(1, i + 1) = rs.Fields(i).Name
        Next i
        
        'Copier le résultat de la requete
        wDestination.Sheets("Output").Range("A2").CopyFromRecordset rs
        
        'Fermer le nouveau classeur
        wDestination.Close SaveChanges:=True
        
        'Envoyer la requete sur la plage G34
        ThisWorkbook.Sheets("EXTRACTION").Range("G34") = query
        
        MsgBox "La requête a bien été effectuée!", vbInformation, "Succès"
    
    Else
        MsgBox "La requête a échouée ! Veuillez réessayer avec des conditions correctes.", vbInformation, "Requête avec SQL"
    End If
    
    
    'Se deconnecter
    rs.Close
    connection.Close
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Une erreur est survenue : " & Err.Description
    Exit Sub

End Sub
