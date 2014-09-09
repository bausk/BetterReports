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
    On Error GoTo EXT
    Dim cbToolbar As CommandBar
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, "BetterReports")
    With cbToolbar
        .Visible = False
    End With
    cbToolbar.Delete
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

Sub Snapshot()
    MsgBox "Snapshot: " & ThisWorkbook.FullName
End Sub
