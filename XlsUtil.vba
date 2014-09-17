Attribute VB_Name = "XlsUtil"
Function clear_sheet(Optional sheet As Worksheet) As Boolean
clear_sheet = False
If sheet Is Nothing Then
    Set sheet = ActiveSheet
End If

sheet.Cells.Delete
Dim xConnection As QueryTable

For Each xConnection In sheet.QueryTables
    xConnection.Delete
Next xConnection

clear_sheet = True
End Function

Function clear_sheet_connections(Optional sheet As Worksheet) As Boolean
clear_sheet_connections = False
If sheet Is Nothing Then
    Set sheet = ActiveSheet
End If

Dim xConnection As QueryTable
For Each xConnection In sheet.QueryTables
    xConnection.Delete
Next xConnection

clear_sheet_connections = True
End Function

Function write_cell(line As Variant, Optional sheet As Worksheet, Optional ByRef row_cadre = 1, Optional ByRef col_cadre = 1) As Boolean
If sheet Is Nothing Then
    Set sheet = ActiveSheet
End If
write_cell = False
Dim current_cell As range
Set current_cell = sheet.Cells(row_cadre, col_cadre)

current_cell = line

row_cadre = current_cell.row + 1
col_cadre = current_cell.Column + 1
write_cell = True
End Function

Function write_row(line_row As Variant, Optional sheet As Worksheet, Optional ByRef row_cadre = 1, Optional ByRef col_cadre = 1) As Boolean
If sheet Is Nothing Then
    Set sheet = ActiveSheet
End If
write_row = False

Dim current_range As range
Set current_range = range(sheet.Cells(row_cadre, col_cadre), sheet.Cells(row_cadre, col_cadre + UBound(line_row)))
For x = 0 To UBound(line_row)
    current_range(x + 1).value = line_row(x)
Next x

row_cadre = row_cadre + 1
col_cadre = col_cadre + UBound(line_row) + 1

write_row = True
End Function

Function write_range(range_2d As Variant, Optional sheet As Worksheet, Optional ByRef row_cadre = 1, Optional ByRef col_cadre = 1, Optional Dirc = Direction.East) As Boolean
If sheet Is Nothing Then
    Set sheet = ActiveSheet
End If
write_range = False

'find out dimensions
row_shift = UBound(range_2d)
col_shift = 0
For i = 0 To UBound(range_2d)
    If col_shift < UBound(range_2d(i)) Then
        col_shift = UBound(range_2d(i))
    End If
Next i

Dim current_range As range
original_col_cadre = col_cadre
For x = 0 To UBound(range_2d)
    result = write_row(range_2d(x), , row_cadre, col_cadre)
    col_cadre = original_col_cadre
    'Set current_range = range(sheet.Cells(row_cadre + x, col_cadre), sheet.Cells(row_cadre + x, col_cadre + col_shift))
    
Next x

'current_range = range_2d

'row_cadre = current_range.SpecialCells(xlCellTypeLastCell).row + 1
col_cadre = original_col_cadre + col_shift + 1

write_range = True
End Function

Function add_file_connection(filename, name, Optional sheet As Worksheet, Optional ByRef row_cadre As Integer = 1, Optional ByRef col_cadre As Integer = 1) As Boolean
If sheet Is Nothing Then
    Set sheet = ActiveSheet
End If
add_file_connection = False

Dim current_connection As QueryTable
'Dim range_string As String

Set start_range = Cells(row_cadre, col_cadre)

Set current_connection = sheet.QueryTables.Add(Connection:="TEXT;" & filename, Destination:=start_range)
With current_connection
    .name = name
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlOverwriteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 65001
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = True
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False

End With

row_cadre = row_cadre + current_connection.ResultRange.Rows.Count
col_cadre = col_cadre + current_connection.ResultRange.Columns.Count

add_file_connection = True

End Function

Function find_named_range(range_name) As Variant
Set find_named_range = Nothing
Set heh = Utility.get_item_by_property(ActiveWorkbook.Names, "Name", range_name)
If heh Is Nothing Then
    Exit Function
    End If
Set find_named_range = heh
End Function

Function rename_range() As Boolean

End Function

Function reset_spec(address As range, name As String, fullspec() As String, ByRef keys() As Variant, ByRef values() As Variant, ByRef captions() As Variant)

Dim spec_array As Variant
Dim keyvalue_pair As Variant
'Dim keys() As Variant
'Dim values() As Variant
Dim current_cell As range
'get spec
Set spec_array = cSettings("Formats")
spec_array = spec_array(name)

'get initial address
init_row = address.row
init_column = address.Column

For x = 0 To UBound(spec_array)
    keyvalue_pair = Split(spec_array(x), ":")
    ReDim Preserve keys(x)
    ReDim Preserve values(x)
    ReDim Preserve captions(x)
    keys(x) = keyvalue_pair(0)
    values(x) = keyvalue_pair(1)
    captions(x) = keyvalue_pair(2)
    Set current_cell = ActiveSheet.Cells(init_row, init_column + keyvalue_pair(0) - 1)
    current_cell.Clear
    create_dropdown current_cell, fullspec
    current_cell.value = keyvalue_pair(1)
Next x

End Function

Function update_named_range(named_range As range, spec_cell As range, fullspec() As String, string_array() As String)


content_init_column = named_range.Column
spec_row = spec_cell.row
content_init_row = named_range.row
Dim new_array() As Variant


For x = 0 To UBound(string_array)
    ReDim Preserve new_array(x)
    new_array(x) = Utility.parse_csv_line(string_array(x))
Next x

For column_increment = 0 To named_range.Columns.Count - 1
    current_spec_value = ActiveSheet.Cells(spec_row, content_init_column + column_increment).value
    spec_position = Utility.in_array(current_spec_value, fullspec)
    If spec_position > -1 Then
        For y = 0 To UBound(string_array)
            temparray = new_array(y)
            XlsUtil.write_cell temparray(spec_position), , content_init_row + y, content_init_column + column_increment
        Next y
    End If
    
Next column_increment

End Function

Function find_spec_position(data_range As range, spec_dictionary() As String) As range



End Function

Function create_dropdown(cell As range, dropdown As Variant) As Boolean
result = False
On Error GoTo FAIL
dropdown_formula = Join(dropdown, ", ")
With cell.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=dropdown_formula
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With

result = True
FAIL:
On Error GoTo 0
End Function
