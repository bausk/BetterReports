Attribute VB_Name = "Config"
'Option Explicit
Public cSettings As Collection

Public Sub get_settings()

'Dim connection_filenames, icons As Variant
Set cSettings = New Collection

connection_filenames = Array("Project.csv", "Rooms.csv")
cSettings.Add connection_filenames, "Filenames"

icons = Array( _
    Array("�������� &�����", 37, "Update"), _
    Array("������� &�����", 231, "SetLocation"), _
    Array("�� &���������", 3633, "SetDefaults"), _
    Array("�&�����", 280, "Snapshot") _
    )
cSettings.Add icons, "Icons"

cSettings.Add "BetterReports", "ToolbarName"

End Sub

