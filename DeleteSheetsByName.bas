Sub DeleteSheetsByName()

'deletes all worksheets if name contains specific text
'user can input this text in messagebox

    Dim shName As String
    Dim xName As String
    Dim xWs As Worksheet
    Dim cnt As Integer
    shName = Application.InputBox("Löscht Datenblätter wenn sie diesen Namen enthalten:", "", _
                                    ThisWorkbook.ActiveSheet.Name, , , , , 2)
    If shName = "" Then Exit Sub
    xName = "*" & shName & "*"
'    MsgBox xName
    Application.DisplayAlerts = False
    cnt = 0
    For Each xWs In ThisWorkbook.Sheets
        If xWs.Name Like xName Then
            xWs.Delete
            cnt = cnt + 1
        End If
    Next xWs
    Application.DisplayAlerts = True
    MsgBox cnt & " Datenblätter wurden gelöscht", vbInformation, ""
End Sub
