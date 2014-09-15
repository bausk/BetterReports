Attribute VB_Name = "FileUtil"
Function get_csv_file_path() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Выберите источник данных"
        .Filters.Clear
        .Filters.Add "Comma Separated Values", "*.csv"
        .Show
        If .SelectedItems.Count = 0 Then
            'MsgBox "Cancel Selected"
            Exit Function
        End If
        'fStr is the file path and name of the file you selected.
        get_csv_file_path = .SelectedItems(1)
    End With
End Function
