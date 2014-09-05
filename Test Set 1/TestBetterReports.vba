Attribute VB_Name = "TestBetterReports"
Function get_connection_filenames()
    get_connection_filenames = Array("Project.csv", "Rooms.csv")
End Function

Sub create_ribbon()
    Set cbToolbar = Application.CommandBars.Add(csToolbarName, msoBarTop, False, True)
    
    With cbToolbar
        Set ctButtonRefresh = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonMove = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonDefaults = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButtonSnapshot = .Controls.Add(Type:=msoControlButton, ID:=2950)
    End With
    
    With ctButtonRefresh
        .Style = msoButtonIconAndCaption
        .Caption = "�������� &�����"
        .FaceId = 37
        .OnAction = "Update"
    End With
    
    With ctButtonMove
        .Style = msoButtonIconAndCaption
        .Caption = "������� &�����"
        .FaceId = 231
        .OnAction = "SetLocation"
    End With
    
    With ctButtonDefaults
        .Style = msoButtonIconAndCaption
        .Caption = "�� &���������"
        .FaceId = 232
        .OnAction = "SetDefaults"
    End With
    
    With ctButtonSnapshot
        .Style = msoButtonIconAndCaption
        .Caption = "�&�����"
        .FaceId = 3633
        .OnAction = "Snapshot"
    End With
    
    With cbToolbar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub

Sub load_csv()
    Dim fStr As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Cancel Selected"
            Exit Sub
        End If
        'fStr is the file path and name of the file you selected.
        fStr = .SelectedItems(1)
    End With

    With ThisWorkbook.Sheets(1).QueryTables.Add(Connection:= _
    "TEXT;" & fStr, Destination:=Range("$A$1"))
        .Name = "CAPTURE"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
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
End Sub

Sub DropDownList()
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="eh, meh"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub


Sub UpdateAllConnections()

    For Each cn In ThisWorkbook.Connections
        cn.Delete
    Next cn

    Dim arrConNames(1) As String
    arrConNames(0) = "Project.csv"
    arrConNames(1) = "Rooms.csv"

    Dim indCon As Integer

    For indCon = LBound(arrConNames) To UBound(arrConNames)
        UpdateConnections arrConNames(indCon)
    Next
End Sub

Sub UpdateConnections(ConName As String)
    FilePath = ThisWorkbook.Path
    ResultPath = FilePath
    ThisWorkbook.Worksheets(1).Select
    ActiveSheet.Cells.Clear
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & ResultPath & "\" & ConName, Destination:=Range( _
        "$A$1"))
        .Name = ConName
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
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub Update()
    With ThisWorkbook.Sheets(1).QueryTables.Add(Connection:="TEXT;" & "mindata.csv", Destination:=Range("$B$1"))
        .Name = "CAPTURE"
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
End Sub

Sub SetLocation()
MsgBox "����������� � ����� ���������"
End Sub


Sub SetDefaults()
MsgBox "������ ����� ������������!"
End Sub


Sub Snapshot()
MsgBox "������ ������"
End Sub

