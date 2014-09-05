Attribute VB_Name = "UI"
Sub add_ribbon()
    
    On Error Resume Next
    Set cbToolbar = Application.CommandBars.Add("BetterReports", msoBarFloating, False, True)
    Set cbToolbar = Utility.get_item_by_name(Application.CommandBars, "BetterReports")
    On Error GoTo 0
        
    With cbToolbar
        Set ctButtonRefresh = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonMove = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonDefaults = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonSnapshot = .Controls.Add(Type:=msoControlButton, ID:=2950)
    End With
    
    With ctButtonRefresh
        .Style = msoButtonIconAndCaption
        .Caption = "Обновить &Отчет"
        .FaceId = 37
        .OnAction = "Update"
    End With
    
    With ctButtonMove
        .Style = msoButtonIconAndCaption
        .Caption = "Выбрать &Место"
        .FaceId = 231
        .OnAction = "SetLocation"
    End With
    
    With ctButtonDefaults
        .Style = msoButtonIconAndCaption
        .Caption = "По &умолчанию"
        .FaceId = 232
        .OnAction = "SetDefaults"
    End With
    
    With ctButtonSnapshot
        .Style = msoButtonIconAndCaption
        .Caption = "С&нимок"
        .FaceId = 3633
        .OnAction = "Snapshot"
    End With
    
    With cbToolbar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub

Sub remove_ribbon()
    Set cbToolbar = Application.CommandBars.Add(csToolbarName, msoBarTop, False, True)
    
    With cbToolbar
        Set ctButtonRefresh = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonMove = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonDefaults = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonSnapshot = .Controls.Add(Type:=msoControlButton, ID:=2950)
    End With
    
    With ctButtonRefresh
        .Style = msoButtonIconAndCaption
        .Caption = "Обновить &Отчет"
        .FaceId = 37
        .OnAction = "Update"
    End With
    
    With ctButtonMove
        .Style = msoButtonIconAndCaption
        .Caption = "Выбрать &Место"
        .FaceId = 231
        .OnAction = "SetLocation"
    End With
    
    With ctButtonDefaults
        .Style = msoButtonIconAndCaption
        .Caption = "По &умолчанию"
        .FaceId = 232
        .OnAction = "SetDefaults"
    End With
    
    With ctButtonSnapshot
        .Style = msoButtonIconAndCaption
        .Caption = "С&нимок"
        .FaceId = 3633
        .OnAction = "Snapshot"
    End With
    
    With cbToolbar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub
