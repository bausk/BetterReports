Attribute VB_Name = "FileUtil"
Function get_csv_file_path() As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Выберите источник данных"
        .InitialFileName = Utility.get_cwd()
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


Function get_strings_from_file(file_path, ByRef result As Boolean, ByRef spec() As String) As String()
result = False
Dim data_array() As String
Dim spec_line As String

On Error GoTo FAIL

Set fso = CreateObject("ADODB.Stream")
Open file_path For Input As #1
Line Input #1, spec_line

spec = Split(spec_line, ",")

i = 0
Do Until EOF(1)
    ReDim Preserve data_array(i)
    Dim aaa
    Line Input #1, aaa
    data_array(i) = Utility.UTF8_16(aaa)

    'data_array(i) = StrConv(data_array(i), vbUnicode)
    i = i + 1
Loop
Close #1
get_strings_from_file = data_array
result = True

FAIL:
On Error GoTo 0
End Function
