Attribute VB_Name = "UI"
Sub add_ribbon()
    config.production_settings
    
    On Error Resume Next
    Set cbToolbar = Application.CommandBars.Add(config.cSettings("ToolbarName"), msoBarFloating, False, True)
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, config.cSettings("ToolbarName"))
    On Error GoTo 0
        
    With cbToolbar
        Dim icons() As Variant
        Dim templates() As Variant
        icons = config.cSettings("Icons")
        templates = config.cSettings("Templates")
        
        Dim ctButton As CommandBarControl
        For x = LBound(icons) To UBound(icons)
            Caption = icons(x)(0)
            FaceId = icons(x)(1)
            OnAction = icons(x)(2)

            If OnAction = "UI.popup" Then
                Set ctButton = .Controls.Add(Type:=msoControlPopup)
                For Each template In templates
                    Set ctSubButton = ctButton.Controls.Add(Type:=msoControlButton)
                    With ctSubButton
                        .Style = msoButtonCaption
                        .Caption = template(0)
                        .OnAction = OnAction
                        .parameter = template(1)
                        .FaceId = template(3)
                        .BeginGroup = True
                        
                    End With
                Next template
            Else
                Set ctButton = .Controls.Add(Type:=msoControlButton, ID:=2950)
                ctButton.Style = msoButtonIconAndCaption
                ctButton.OnAction = OnAction
                ctButton.FaceId = FaceId
            End If
           
            ctButton.Caption = Caption
            
            
            
        Next
    End With
     
    With cbToolbar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub

Sub remove_ribbon()
    On Error GoTo EXT
    Dim cbToolbar As CommandBar
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, "BetterReports")
    With cbToolbar
        .Visible = False
    End With
    cbToolbar.Delete
EXT:
End Sub

Sub remove_buttons()
    config.production_settings
    
    'On Error GoTo EXT
    Dim cbToolbar As CommandBar
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, "BetterReports")
    If Not cbToolbar Is Nothing Then
        With cbToolbar
            Dim icons() As Variant
            icons = config.cSettings("Icons")
            For x = LBound(icons) To UBound(icons)
                Caption = icons(x)(0)
                FaceId = icons(x)(1)
                OnAction = icons(x)(2)
                
                'Behave wicked smaht when deleting buttons
                Dim existing_button As CommandBarButton
                Set existing_button = Utility.get_item_by_property(.Controls, "OnAction", "'" & Utility.get_fullname() & "'!" & OnAction)
                If Not existing_button Is Nothing Then
                    existing_button.Delete
                End If
            Next
        End With
    End If
    
    'With cbToolbar
    '    .Visible = False
    'End With
    'cbToolbar.Delete
EXT:
End Sub

Sub refresh_buttons()
    config.production_settings
    
    On Error Resume Next
    Set cbToolbar = Application.CommandBars.Add(config.cSettings("ToolbarName"), msoBarFloating, False, True)
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, config.cSettings("ToolbarName"))
    On Error GoTo 0
        
    With cbToolbar
        Dim icons() As Variant
        icons = config.cSettings("Icons")
        Dim templates() As Variant
        templates = config.cSettings("Templates")
        For x = LBound(icons) To UBound(icons)
            Caption = icons(x)(0)
            FaceId = icons(x)(1)
            OnAction = icons(x)(2)
            
            'Behave wicked smaht when adding buttons
            Dim existing_button As CommandBarControl
            Set existing_button = Utility.get_item_by_property(.Controls, "Caption", Caption)
            If Not existing_button Is Nothing Then
                existing_button.Delete
            End If
            
            
            If OnAction = "UI.popup" Then
                Set ctButton = .Controls.Add(Type:=msoControlPopup)
                For Each template In templates
                    Set ctSubButton = ctButton.Controls.Add(Type:=msoControlButton)
                    With ctSubButton
                        .Style = msoButtonIconAndCaption
                        .Caption = template(0)
                        .OnAction = "'" & OnAction & """" & template(1) & """'"
                        '.parameter = 1
                        .FaceId = template(3)
                        .BeginGroup = True
                        
                    End With
                Next template
            Else
                Set ctButton = .Controls.Add(Type:=msoControlButton, ID:=2950)
                ctButton.Style = msoButtonIconAndCaption
                ctButton.OnAction = OnAction
                ctButton.FaceId = FaceId
            End If
           
            ctButton.Caption = Caption
            
            
        Next
    End With
     
    With cbToolbar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub


Sub set_location()
    Dim ThisRng As range
    On Error Resume Next
    Set ThisRng = Application.InputBox("Выберите область отчета", "Выбрать область", Type:=8)
    On Error GoTo 0
    If ThisRng Is Nothing Then Exit Sub
    
    'find name
    
    config.production_settings
    'Dim row_cadre As Integer, column_cadre As Integer
    Dim keys() As Variant, values() As Variant, captions() As Variant
    
    Dim connection_name As String, file_path As String, file_name As String
    'Dim rangename As Variant
    Dim string_array() As String
    
    connection_names = cSettings("Names")
    file_names = cSettings("Filenames")
    
    Dim named_range As range
    Set named_range = Nothing
    Dim range_name As name
    
    Set named_range = XlsUtil.find_connection(connection_names, file_names, connection_name, file_name, range_name)
    If range_name Is Nothing Then Exit Sub
    If range_name = "" Then Exit Sub
    If connection_name = "" Then Exit Sub
    range_name.Delete
'    named_range.Delete
    
    ThisRng.name = connection_name
   
    update
    
End Sub

Sub set_defaults()


'ActiveSheet.Rows.Ungroup
config.production_settings
Dim row_cadre As Integer, column_cadre As Integer
Dim keys() As Variant, values() As Variant, captions() As Variant
Dim named_range As range
Dim range_name As name
Dim result As Boolean
Set named_range = Nothing
Dim connection_name As String, file_path As String
'Dim rangename As Variant
Dim string_array() As String

connection_names = cSettings("Names")
file_names = cSettings("Filenames")
row_cadre = 1
column_cadre = 1

For x = 0 To UBound(connection_names)
    Set range_name = XlsUtil.find_named_range(connection_names(x))
    If Not range_name Is Nothing Then
        Set named_range = ThisWorkbook.ActiveSheet.range(range_name)
        
        connection_name = connection_names(x)
        file_name = file_names(x)
        Exit For
    End If
Next x

If named_range Is Nothing Then
    file_path = FileUtil.get_csv_file_path()
    For x = 0 To UBound(file_names)
        If file_path = Utility.get_cwd & file_names(x) Then
            connection_name = connection_names(x)
            file_name = file_names(x)
            Exit For
        End If
    Next x
End If

If Not range_name Is Nothing Then range_name.Delete
Set named_range = Nothing

If connection_name = "" Then
    MsgBox "Не найден подходящий файл, используется источник по умолчанию: " & file_names(1)
    connection_name = connection_names(1)
    file_name = file_names(1)
End If

Dim fullspec() As String
string_array = FileUtil.get_strings_from_file(Utility.get_cwd & file_name, result, fullspec)
If result = False Then GoTo NOFILEFOUND

'write spec
ActiveSheet.Cells.Clear
Dim spec_cell As range
Set spec_cell = ActiveSheet.Cells(row_cadre, column_cadre)
XlsUtil.reset_spec spec_cell, connection_name, fullspec, keys, values, captions
'hide spec
ActiveSheet.Rows(row_cadre).EntireRow.Hidden = True

'write table caption
row_cadre = row_cadre + 1

'for styling
Set first_cell = ActiveSheet.Cells(row_cadre, column_cadre)

table_caption = cSettings("Captions")(connection_name)
XlsUtil.write_cell table_caption, , row_cadre, column_cadre

column_cadre = 1

'write column captions
max_col_cadre = 1
For x = 0 To UBound(keys)
    static_cadre = row_cadre
    'column_cadre = column_cadre + keys(x) - 1
    XlsUtil.write_cell captions(x), , static_cadre, column_cadre + keys(x) - 1
    If max_col_cadre < column_cadre + keys(x) - 1 Then
        max_col_cadre = column_cadre + keys(x) - 1
    End If
Next x

'create named range
row_cadre = row_cadre + 1
column_cadre = 1
Set named_range = range(Cells(row_cadre, column_cadre), Cells(row_cadre, max_col_cadre))
named_range.name = connection_name

Set last_cell = ActiveSheet.Cells(row_cadre + UBound(string_array), max_col_cadre)

'write data string by string, using update_table
XlsUtil.update_named_range named_range, spec_cell, fullspec, string_array

'Style
Dim style_range As range
Set style_range = range(first_cell, last_cell)
style_range.Style = "Output"

Exit Sub
NOFILEFOUND:
MsgBox "Файл " & "" & " не найден или поврежден", vbExclamation
End Sub


Sub popup(parameter As String)

'MsgBox "hehe"
'MsgBox parameter

ActiveSheet.Cells.Clear
'ActiveSheet.Rows.Ungroup
config.production_settings
Dim row_cadre As Integer, column_cadre As Integer
Dim keys() As Variant, values() As Variant, captions() As Variant
Dim named_range As range
Dim range_name As name
Dim result As Boolean
Set named_range = Nothing
Dim connection_name As String, file_path As String
'Dim rangename As Variant
Dim string_array() As String

connection_names = cSettings("Names")
file_names = cSettings("Filenames")
row_cadre = 1
column_cadre = 1

connection_name = parameter


For x = 0 To UBound(connection_names)
    Set range_name = XlsUtil.find_named_range(connection_names(x))
    If Not range_name Is Nothing Then range_name.Delete
    Set named_range = Nothing
    If connection_name = connection_names(x) Then
        file_name = file_names(x)
    End If
Next x



Dim fullspec() As String
string_array = FileUtil.get_strings_from_file(Utility.get_cwd & file_name, result, fullspec)
If result = False Then Exit Sub

'write spec
Dim spec_cell As range
Set spec_cell = ActiveSheet.Cells(row_cadre, column_cadre)
XlsUtil.reset_spec spec_cell, connection_name, fullspec, keys, values, captions
'hide spec
ActiveSheet.Rows(row_cadre).EntireRow.Hidden = True

'write table caption
row_cadre = row_cadre + 1

'for styling
Set first_cell = ActiveSheet.Cells(row_cadre, column_cadre)

table_caption = cSettings("Captions")(connection_name)
XlsUtil.write_cell table_caption, , row_cadre, column_cadre

column_cadre = 1

'write column captions
max_col_cadre = 1
For x = 0 To UBound(keys)
    static_cadre = row_cadre
    'column_cadre = column_cadre + keys(x) - 1
    XlsUtil.write_cell captions(x), , static_cadre, column_cadre + keys(x) - 1
    If max_col_cadre < column_cadre + keys(x) - 1 Then
        max_col_cadre = column_cadre + keys(x) - 1
    End If
Next x

'create named range
row_cadre = row_cadre + 1
column_cadre = 1
Set named_range = range(Cells(row_cadre, column_cadre), Cells(row_cadre, max_col_cadre))
named_range.name = connection_name

Set last_cell = ActiveSheet.Cells(row_cadre + UBound(string_array), max_col_cadre)

'write data string by string, using update_table
XlsUtil.update_named_range named_range, spec_cell, fullspec, string_array

'Style
Dim style_range As range
Set style_range = range(first_cell, last_cell)
style_range.Style = "Output"


End Sub

Sub snapshot()
    Dim asheet As Worksheet, new_sheet As Worksheet
    Set asheet = ThisWorkbook.ActiveSheet
    Dim new_name As String
    On Error Resume Next
    i = 0
    Do While new_name = ""
        i = i + 1
        Set wsSheet = Nothing
        Set wsSheet = Sheets(asheet.name & " (" & i & ")")
        If wsSheet Is Nothing Then
            new_name = asheet.name & " (" & i & ")"
        End If
    Loop
    
Set new_sheet = Worksheets.Add(After:=asheet)
new_sheet.name = new_name
asheet.Cells.Copy
With new_sheet
    .Cells.PasteSpecial xlValues
    .Cells.PasteSpecial xlFormats
End With
asheet.Activate
Application.CutCopyMode = False
End Sub


Sub update()

config.production_settings
Dim row_cadre As Integer, column_cadre As Integer
Dim keys() As Variant, values() As Variant, captions() As Variant

Dim connection_name As String, file_path As String, file_name As String


connection_names = cSettings("Names")
file_names = cSettings("Filenames")

row_cadre = 1
column_cadre = 1

Dim named_range As range
Set named_range = Nothing
Dim range_name As name
Set range_name = Nothing

Set named_range = XlsUtil.find_connection(connection_names, file_names, connection_name, file_name, range_name)

If range_name Is Nothing Then Exit Sub
If range_name = "" Then Exit Sub
If connection_name = "" Then Exit Sub

Dim fullspec() As String
Dim result As Boolean
Dim string_array() As String
string_array = FileUtil.get_strings_from_file(Utility.get_cwd & file_name, result, fullspec)
If result = False Then Exit Sub

Dim spec_cell As range
Set spec_cell = XlsUtil.find_spec_position(named_range, fullspec)
If spec_cell Is Nothing Then Exit Sub

XlsUtil.update_named_range named_range, spec_cell, fullspec, string_array

End Sub
