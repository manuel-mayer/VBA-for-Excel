Sub CopySheetList()

'copies the worksheet "Vorlage" and renames it after a list.

'declare variables
Dim I As Variant, ws As Worksheet, alist As ArrayList

'Define List
Set alist = New ArrayList
alist.Add "1.6.5"
alist.Add "1.8.1.2"
alist.Add "1.8.2.1"
alist.Add "3.2.3"

'copy the sheet and name them after the values in the list
For Each I In alist
    Sheets("Vorlage").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = I
Next I

End Sub
