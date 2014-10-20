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
col_cadre = current_cell.column + 1
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


Function find_connection(connection_names, file_names, ByRef connection_name, ByRef file_name, ByRef range_name) As range
Dim named_range As range
On Error GoTo FAIL
For x = 0 To UBound(connection_names)
    Set range_name = XlsUtil.find_named_range(connection_names(x))
    If Not range_name Is Nothing Then
        connection_name = connection_names(x)
        file_name = file_names(x)
        Set named_range = range_name.RefersToRange
        Exit For
    End If
Next x

If named_range Is Nothing Then
    'file_path = FileUtil.get_csv_file_path()
    'For x = 0 To UBound(file_names)
    '    If file_path = Utility.get_cwd & file_names(x) Then
    '        connection_name = connection_names(x)
    '        file_name = file_names(x)
    '        Exit For
    '    End If
    'Next x
    MsgBox "В таблице нет ни одной ссылки на файл отчета. Вставьте ведомость из шаблона.", vbExclamation
End If

Set find_connection = named_range
FAIL:
On Error GoTo 0
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
init_column = address.column

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

Sub update_named_range(named_range As range, spec_cell As range, fullspec() As String, string_array() As String)

'0. Check for no data from CSV parsing
If UBound(string_array) < 1 Then
    If string_array(0) = "" Then GoTo NODATA
End If

Dim new_array() As Variant
For x = 0 To UBound(string_array)
    ReDim Preserve new_array(x)
    new_array(x) = Utility.parse_csv_line(string_array(x))
Next x

Dim project_file As String
Dim chapter_name As String
file_names = cSettings("Filenames")
substitutions = cSettings("Substitutions")

'1. Performing substitutions
project_file = Utility.get_cwd & file_names(0)
For Each substitution In substitutions
    chapter_name = substitution(0)
    what_name = substitution(1)
    chapter_index = substitution(2)
    Dim keys_array() As String, chapter_table() As Variant
    result = FileUtil.extract_table(project_file, chapter_name, keys_array, chapter_table)
    
    If result = False Then GoTo NOPROJECTFILE
    
    index = Utility.in_array(what_name, fullspec)
    If index = -1 Then GoTo CONTINUE
    
    For i = 0 To UBound(new_array)
        If UBound(new_array(i)) = -1 Then
            GoTo EXITFOR
        End If
        
        element = new_array(i)(index)
        key_index = Utility.in_array(element, keys_array)
        element = chapter_table(key_index)(chapter_index)
        new_array(i)(index) = element
EXITFOR:
    Next i
CONTINUE:
Next substitution

'Get shortened spec
colcount = named_range.Columns.Count - 1
Dim valid_indices() As Integer
Dim newspec() As String
spec_row = spec_cell.row
spec_col = spec_cell.column
cur_spec_value = Cells(spec_row, spec_col).value
test_position = Utility.in_array(cur_spec_value, fullspec)
If test_position > -1 Then
    val_indices_len = 0
    ReDim Preserve valid_indices(val_indices_len)
    ReDim Preserve newspec(val_indices_len)
    valid_indices(val_indices_len) = test_position
    newspec(val_indices_len) = cur_spec_value
End If

For ii = 1 To colcount
    cur_spec_value = Cells(spec_row, spec_col + ii).value
    test_position = Utility.in_array(cur_spec_value, fullspec)
    If test_position > -1 Then
        val_indices_len = UBound(valid_indices) + 1
        ReDim Preserve valid_indices(val_indices_len)
        ReDim Preserve newspec(val_indices_len)
        valid_indices(val_indices_len) = test_position
        newspec(val_indices_len) = cur_spec_value
    End If
Next ii

'Truncate new_array to actual spec
'respec new_array from fullspec to conform to newspec using valid_indices
'Dim valid_array() As Variant
If Utility.lazy_compare(fullspec, newspec) = True Then GoTo AGGREGATION 'specs are identical, no need to respec
'otherwise, rewrite each line of new_array
For ii = 0 To UBound(new_array)
    old_string = new_array(ii)
    Dim new_string() As Variant
    flag = False
    
    For Each idx In valid_indices
        'compile new_string by taking values according to valid_indices
        If flag = False Then
            new_len = 0
            flag = True
        Else
            new_len = UBound(new_string) + 1
            
        End If
        ReDim Preserve new_string(new_len)
        new_string(new_len) = old_string(idx)
    Next idx
    'replace old_string with new_string
    new_array(ii) = new_string
Next ii
'truncating complete
'replace old spec with new spec
fullspec = newspec

AGGREGATION:
'Perform aggregation if there is QUANTITY classifier in active spec
Dim is_aggregate As Boolean
spec_row = spec_cell.row
'1. figure out if aggregate.
Set temp_setting = cSettings("TypeInfo")
Set name = named_range.name
is_aggregate = temp_setting(name.name)(0)
If is_aggregate = False Then GoTo NOAGGREGATE

'2. add to fullspec as last field
quantity_classifier = cSettings("Aggregating")
quant_index = UBound(fullspec) + 1
ReDim Preserve fullspec(quant_index)
fullspec(quant_index) = quantity_classifier

'3. generate new array
'3a. init with first string extended with quantity=0
Dim aggregate_array() As Variant
ReDim Preserve aggregate_array(0)
Dim init_string() As Variant
init_string = new_array(0)
Dim aggregate_pos As Integer
aggregate_pos = UBound(init_string) + 1

ReDim Preserve init_string(aggregate_pos)
init_string(aggregate_pos) = 0
aggregate_array(0) = init_string

For Each old_string In new_array
    Dim old_string_srsly() As Variant
    old_string_srsly = old_string
    For xx = 0 To UBound(aggregate_array)
        new_string = aggregate_array(xx)
        If Utility.lazy_compare(new_string, old_string) = True Then
            'old_string is same in first elements as new one.
            'increase quantity by 1, discard old_string by continueing
            new_string(aggregate_pos) = new_string(aggregate_pos) + 1
            'no need to look further
            aggregate_array(xx) = new_string
            GoTo EXITSCAN
        End If
    Next xx
    'end of foreach iteration means we have a new element to aggregate to
    'add new line, init with quantity=1
    ReDim Preserve old_string_srsly(aggregate_pos)
    old_string_srsly(aggregate_pos) = 1
    new_line_pos = UBound(aggregate_array) + 1
    ReDim Preserve aggregate_array(new_line_pos)
    aggregate_array(new_line_pos) = old_string_srsly
    
EXITSCAN:
    'found and added new quantity, process next line in original array
Next old_string

'replace new_array with aggregate_array
new_array = aggregate_array

NOAGGREGATE:
'we may or may have not replaced the new_array with aggregate array, but that does not influence further logic
content_init_column = named_range.column
content_init_row = named_range.row
Dim affected_range As range

For column_increment = 0 To named_range.Columns.Count - 1
    current_spec_value = ActiveSheet.Cells(spec_row, content_init_column + column_increment).value
    spec_position = Utility.in_array(current_spec_value, fullspec)
    If spec_position > -1 Then
        For y = 0 To UBound(new_array)
            temparray = new_array(y)
            If UBound(temparray) = -1 Then GoTo EMPTYSTRING
            XlsUtil.write_cell temparray(spec_position), , content_init_row + y, content_init_column + column_increment
EMPTYSTRING:
        Next y
    End If
    max_row = content_init_row + UBound(new_array)
    max_col = content_init_column + column_increment
Next column_increment

Set affected_range = range(ActiveSheet.Cells(content_init_row, content_init_column), ActiveSheet.Cells(max_row, max_col))
affected_range.Select


Dim settings_array() As Variant
settings_array = cSettings("Style Locals")

Set range_style = Utility.get_item_by_property_m(ThisWorkbook.Styles, "Name", settings_array)
'Utility.choose_one_existing settings_array, ThisWorkbook.Styles(0).name, range_name
affected_range.Style = range_style

Exit Sub
NOPROJECTFILE:

MsgBox "Для вывода отчета требуется файл Project.csv, но он отсутствует или повреждён. Выполните экспорт отчета из Tornado.", vbExclamation

NODATA:

End Sub

Function find_spec_position(data_range As range, fullspec() As String) As range
end_row = data_range.row - 1
Dim csheet As Worksheet
Set csheet = ActiveSheet
Dim result As range
Set find_spec_position = Nothing

For x = 1 To end_row
    'iterate through every row segment upper than the named data_range
    For Each cCell In range(csheet.Cells(x, data_range.column), csheet.Cells(x, data_range.column + data_range.Columns.Count - 1))
        If Utility.in_array(cCell.value, fullspec) > -1 Then
            Set find_spec_position = csheet.Cells(x, data_range.column)
            Exit Function
        End If
    Next cCell
Next x

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
