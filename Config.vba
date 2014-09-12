Attribute VB_Name = "Config"
'Option Explicit
Public cSettings As Collection
Public Dir As Direction
'DayOff = workdayconstant

Public Sub production_settings()

'Dim connection_filenames, icons As Variant
Set cSettings = New Collection

connection_filenames = Array("Project.csv", "Rooms.csv")
cSettings.Add connection_filenames, "Filenames"

icons = Array( _
    Array("Обновить &отчет", 37, "Update"), _
    Array("Выбрать &место", 231, "SetLocation"), _
    Array("По &умолчанию", 3633, "SetDefaults"), _
    Array("С&нимок", 280, "UI.Snapshot") _
    )
cSettings.Add icons, "Icons"

cSettings.Add "BetterReports", "ToolbarName"

End Sub

Public Sub mock_settings(i As Integer)

'Dim connection_filenames, icons As Variant
Set cSettings = New Collection

If i = 1 Then
    connection_filenames = Array("Test Set 1\Project.csv", "Test Set 1\Rooms.csv", "Test Set 1\Doors.csv", "Test Set 1\Windows.csv")
    cSettings.Add connection_filenames, "Filenames"
    
    icons = Array( _
        Array("Обновить &отчет", 37, "Update"), _
        Array("Выбрать &место", 231, "SetLocation"), _
        Array("По &умолчанию", 3633, "SetDefaults"), _
        Array("С&нимок", 280, "Snapshot") _
        )
    cSettings.Add icons, "Icons"
    
    cSettings.Add "BetterReports", "ToolbarName"
ElseIf i = 2 Then
    connection_filenames = Array("Test Set 2\mindata.csv")
    cSettings.Add connection_filenames, "Filenames"
End If

End Sub

