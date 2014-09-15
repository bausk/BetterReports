Attribute VB_Name = "TestBetterReports"
Sub load_csv()
    Dim fStr As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Cancel Selected"
            Exit Sub
        End If
        'fStr is the file path and name of the file you selected.
        fStr = .SelectedItems(1)
    End With
End Sub

Sub DropDownList()
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="eh, meh"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Sub TestRangeSelection()
Dim ThisRng As range
Set ThisRng = Application.InputBox("Select a range", "Get Range", Type:=8)

End Sub
