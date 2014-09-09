Attribute VB_Name = "UI"
Sub add_ribbon()
    
    On Error Resume Next
    Dim localconfig As Collection
    
    Set localconfig = config.get_config()
    
    Set cbToolbar = Application.CommandBars.Add(localconfig("ToolbarName"), msoBarFloating, False, True)
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, localconfig("ToolbarName"))
    On Error GoTo 0
        
    With cbToolbar
        Dim icons() As Variant
        icons = config.get_icons()
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
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, "BetterReports")
    With cbToolbar
        .Visible = False
    End With
    
    
End Sub
