Sub AutoSortData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Remplacez par le nom de votre feuille

    ' Tri des données selon la colonne A (par ordre croissant)
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row), _
                           Order:=xlAscending
    ws.Sort.SetRange Range("A1:D" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row) ' Modifier la plage
    ws.Sort.Apply

    MsgBox "Données triées avec succès !"
End Sub
