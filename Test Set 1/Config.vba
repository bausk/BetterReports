Attribute VB_Name = "Config"
Function get_connection_filenames()
    get_connection_filenames = Array("Project.csv", "Rooms.csv")
End Function

Function get_icons()
    get_icons = Array( _
    Array("�������� &�����", 37, "Update"), _
    Array("������� &�����", 231, "SetLocation"), _
    Array("�� &���������", 3633, "SetDefaults"), _
    Array("�&�����", 280, "Snapshot") _
    )
End Function
