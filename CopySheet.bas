Sub CopySheet()

'copies the worksheet "Vorlage" multiple times and renames it based on a list in the worksheet "Daten"

'declare variables
Dim i As Long, LastRow As Long, ws As Worksheet

'activate sheet and count the rows, store it in the variabe LastRow
Sheets("Daten").Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'copy the sheet and name them after the values in the list
For i = 3 To LastRow    'skips the first 2 rows
    Sheets("Vorlage").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = Sheets("Daten").Cells(i, 1)
Next i

End Sub
