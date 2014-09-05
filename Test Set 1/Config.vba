Attribute VB_Name = "Config"
Function get_connection_filenames()
    get_connection_filenames = Array("Project.csv", "Rooms.csv")
End Function

Function get_icons()
    get_icons = Array( _
    Array("Обновить &Отчет", 37, "Update"), _
    Array("Выбрать &Место", 231, "SetLocation"), _
    Array("По &умолчанию", 3633, "SetDefaults"), _
    Array("С&нимок", 280, "Snapshot") _
    )
End Function
