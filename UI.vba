Attribute VB_Name = "UI"
Sub add_ribbon()
    config.production_settings
    
    On Error Resume Next
    Set cbToolbar = Application.CommandBars.Add(config.cSettings("ToolbarName"), msoBarFloating, False, True)
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, config.cSettings("ToolbarName"))
    On Error GoTo 0
        
    With cbToolbar
        Dim icons() As Variant
        icons = config.cSettings("Icons")
        For x = LBound(icons) To UBound(icons)
            Caption = icons(x)(0)
            FaceId = icons(x)(1)
            OnAction = icons(x)(2)
            
            Set ctButton = .Controls.Add(Type:=msoControlButton, ID:=2950)
            
            ctButton.Style = msoButtonIconAndCaption
            ctButton.Caption = Caption
            ctButton.FaceId = FaceId
            ctButton.OnAction = OnAction
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
        For x = LBound(icons) To UBound(icons)
            Caption = icons(x)(0)
            FaceId = icons(x)(1)
            OnAction = icons(x)(2)
            
            'Behave wicked smaht when adding buttons
            Dim existing_button As CommandBarButton
            Set existing_button = Utility.get_item_by_property(.Controls, "Caption", Caption)
            If Not existing_button Is Nothing Then
                existing_button.Delete
            End If
            Set ctButton = .Controls.Add(Type:=msoControlButton, ID:=2950)
            
            ctButton.Style = msoButtonIconAndCaption
            ctButton.Caption = Caption
            ctButton.FaceId = FaceId
            ctButton.OnAction = OnAction
        Next
    End With
     
    With cbToolbar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub

Sub update()
    MsgBox "Update: " & ThisWorkbook.FullName
End Sub

Sub set_location()
    Dim ThisRng As range
    Set ThisRng = Application.InputBox("Select a range", "Get Range", Type:=8)
End Sub

Sub set_defaults()

config.mock_settings 1
Dim Rng As range
Set Rng = Nothing
Dim connection_name As String
Dim rangename As Variant

For Each rangename In cSettings("Names")
    Set Rng = XlsUtil.find_named_range(rangename)
    If Not Rng Is Nothing Then
        connection_name = rangename
        Exit For
    End If
Next rangename
If Rng Is Nothing Then
    connection_name = FileUtil.get_csv_file_path()
End If
MsgBox connection_name

End Sub

Sub set_source()


End Sub


Sub snapshot()
    MsgBox "Snapshot: " & ThisWorkbook.FullName
End Sub
