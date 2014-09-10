Attribute VB_Name = "TestBetterReports"
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
        .name = "CAPTURE"
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
        .name = ConName
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
        .name = "CAPTURE"
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
