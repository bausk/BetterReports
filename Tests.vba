Attribute VB_Name = "Tests"
Sub unittest_getbyproperty()
    'Test: getting getting specific items from a collection by their property
    'Mockup
    On Error Resume Next
    Dim cbToolbar As CommandBar
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, "testing1")
    cbToolbar.Delete
    On Error GoTo 0
    Set cbToolbar = Application.CommandBars.Add("testing1", msoBarFloating, False, True)
    
    With cbToolbar
        Set ctTestButton = .Controls.Add(Type:=msoControlButton, ID:=2950)
    End With
    
    With ctTestButton
        .Style = msoButtonIconAndCaption
        .Caption = "Тестовая кнопка"
        .FaceId = 37
        .OnAction = "TestAction"
    End With

    'Testing
    Set testItem = Utility.get_item_by_property(cbToolbar.Controls, "Caption", "Тестовая кнопка")
    'assert testItem is same as ctTestButton
    If testItem.InstanceId <> ctTestButton.InstanceId Then
        MsgBox "Test failed"
    End If
    
    'Cleanup
    cbToolbar.Delete

End Sub

Sub unittest_config()
'Test: example on how to get config properties

config.get_settings
Dim testconfig As Collection

Set testconfig = config.cSettings

If config.cSettings("ToolbarName") <> "BetterReports" Then
    MsgBox "Test Failed"
End If

If config.cSettings("Icons")(0)(0) <> "Обновить &отчет!" Then
    MsgBox "Test Failed"
End If

If config.cSettings("Filenames")(0) <> "Project.csv" Then
    MsgBox "Test Failed"
End If

End Sub

Sub test_addconnection()
Dim row_cadre As Integer, col_cadre As Integer, filename As String
row_cadre = 4
col_cadre = 5

config.mock_settings 2
dirname = Utility.get_cwd()
Dim result As Boolean
filename = config.cSettings("Filenames")(0)

'Call being tested is add_file_connection() function
'also provides example of usage
XlsUtil.clear_sheet
result = XlsUtil.add_file_connection(dirname & filename, "test_connection", , row_cadre, col_cadre)
'End of test fragment
Dim sheet As Worksheet

If result = False Or StrComp(range("F5").value, """Erledigt""") <> 0 Then
    MsgBox "Test failed: value " & range("F5").value & " is not equal to ""Erledigt"" or file not found"
End If

If row_cadre <> 6 Or col_cadre <> 7 Then
    MsgBox "Test failed: either of the pointers to the next writable cell is wrong"
End If

XlsUtil.clear_sheet_connections
End Sub

Sub test_writers()

'Test the write_cell, write_row and write_range routines that use cadre pointer

Dim row_cadre As Integer
row_cadre = 1
col_cadre = 1
line = "Ooga booga"
For x = 1 To 5
    result = XlsUtil.write_cell(line, , row_cadre)
Next x

row = Array("Faith", "Plus", "One")
result = XlsUtil.write_row(row, , row_cadre, col_cadre)

range_2d = Array( _
    Array("I", "can"), _
    Array("We", "can", "be", "heroes"), _
    Array("They", "can") _
    )
For x = 1 To 3
    result = XlsUtil.write_range(range_2d, , row_cadre, col_cadre)
Next x
original_col_position = col_cadre
For x = 1 To 3
    result = XlsUtil.write_range(range_2d, , row_cadre, col_cadre)
    col_cadre = original_col_position
Next x
End Sub

Sub test_dataset_1()
'Made after writers in XlsUtils were ready for prime time
'Testing CSV import from a dataset in local folder and putting them into a table
' use mockup config
Dim row_cadre As Integer
row_cadre = 1

config.mock_settings 1
dirname = Utility.get_cwd()
Dim result As Boolean

XlsUtil.clear_sheet
'Testing fragment
'For every file in config, write header and connection into current sheet, sequentially
For x = LBound(config.cSettings("Filenames")) To UBound(config.cSettings("Filenames"))
    filename = config.cSettings("Filenames")(x)
    XlsUtil.write_cell config.cSettings("Filenames")(x), , row_cadre
    XlsUtil.add_file_connection dirname & filename, filename, , row_cadre
Next x

'End of testing fragment
End Sub

Sub test_parse_csv_line()

Dim line As String
line = """D:\Documents\Dropbox\ASCONProjects\BetterReports\Test Set 1\Simple Test 1.tor"",""20121024"",""2014-9-9T12:41:47"",""2014-9-11T14:55:26"""
result = Utility.parse_csv_line(line, ",", """")
'MsgBox result(3)
If result(3) <> "2014-9-11T14:55:26" Then MsgBox "Test 1 failed"

line = ""","""","""
result = Utility.parse_csv_line(line)
If result(0) <> ",""," Then MsgBox "Test 2 failed"

line = ",""Field 1"""","",""""""Field 2"","""""""""
result = Utility.parse_csv_line(line)
If result(1) <> "Field 1""," Or result(3) <> """" Then MsgBox "Test 3 failed"

line = """B,C"",""D"""",E"""""",F,"
result = Utility.parse_csv_line(line)
If result(1) <> "D"",E""" Or result(2) <> "F" Then MsgBox "Test 4 failed"

line = ""","","""",""A"",""B,C"",""D"""",E"""""",F,"
result = Utility.parse_csv_line(line)
'MsgBox ":" & result(1) & ": :" & result(2) & ": :" & result(3) & ": :" & result(4)


End Sub

Sub test_properties()

PropType = Utility.get_property_type("huehuehuehuehue")
Set Prop1 = ActiveWorkbook.CustomDocumentProperties.Add(name:="Range", LinkToContent:=False, Type:=PropType, value:="Huhuehue")
End Sub

Sub test_find_named_range()

Dim r As range
Set r = XlsUtil.find_named_range("Sample")
MsgBox r.name

End Sub


Sub production()
'Testing getting data from production folder (i.e. current file folder)

'Not facking needed for ahnythin' good
End Sub
