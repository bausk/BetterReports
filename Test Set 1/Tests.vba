Attribute VB_Name = "Tests"
Sub test1()
    
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
