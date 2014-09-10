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


Sub test_dataset_1()
'Testing CSV import from a dataset in local folder and putting them into a table
'Where to begin
row_cadre = 1
' use mockup config
config.mock_settings 1

Dim wb As Workbook
Dim ws As Worksheet

Set wb = ActiveWorkbook
Set ws = ActiveSheet
'Set ws = wb.Sheets.Item(1)


For x = LBound(config.cSettings("Filenames")) To UBound(config.cSettings("Filenames"))
    XlsUtil.clear_sheet ws
    XlsUtil.writeline row_cadre, config.cSettings("Filenames")(x)
    XlsUtil.add_file_connection row_cadre, config.cSettings("Filenames")(x)
Next x

Set TextCell = ws.Cells.Item(1, 1)

TextCell.value = "dfdsfdsf"


End Sub

Sub production()
'Testing getting data from production folder (i.e. current file folder)

'Not facking needed for ahnythin' good
End Sub
