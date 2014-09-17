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

connection_names = Array("TornadoProject", "TornadoRooms")
cSettings.Add connection_names, "Names"

Dim report_formats As Collection
Set report_formats = New Collection
report_formats.Add Array("1:NUMBER:����� ���������", "2:NAME:������������", "4:AREA:�������, �.��."), "TornadoRooms"
report_formats.Add Array("1:NUMBER", "2:NAME", "4:AREA"), "TornadoDoors"
report_formats.Add Array("1:NUMBER", "2:NAME", "4:AREA"), "TornadoWindows"
cSettings.Add report_formats, "Formats"

Dim report_headings As Collection
Set report_headings = New Collection
report_headings.Add "����������� ���������", "TornadoRooms"
report_headings.Add "����������� ������� �������", "TornadoDoors"
report_headings.Add "����������� ������� �������", "TornadoWindows"
cSettings.Add report_headings, "Captions"

icons = Array( _
    Array("������� &�����", 231, "UI.set_location"), _
    Array("�������� &�����", 37, "UI.update"), _
    Array("�� �������", 3633, "UI.set_defaults"), _
    Array("�&�����", 280, "UI.snapshot") _
    )
cSettings.Add icons, "Icons"

templates = Array( _
    Array("������� &�����", 231, "UI.set_location"), _
    Array("�������� &�����", 37, "UI.update"), _
    Array("�� �������", 3633, "UI.set_defaults"), _
    Array("�&�����", 280, "UI.snapshot") _
    )
cSettings.Add templates, "Templates"


cSettings.Add "BetterReports", "ToolbarName"

End Sub

Public Sub mock_settings(i As Integer)

'Dim connection_filenames, icons As Variant
Set cSettings = New Collection

If i = 1 Then
    connection_filenames = Array("Test Set 1\Project.csv", "Test Set 1\Rooms.csv", "Test Set 1\Doors.csv", "Test Set 1\Windows.csv")
    cSettings.Add connection_filenames, "Filenames"
    
    connection_names = Array("TornadoProject", "TornadoRooms", "TornadoDoors", "TornadoWindows")
    cSettings.Add connection_names, "Names"

    Dim report_formats As Collection
    Set report_formats = New Collection
    report_formats.Add Array("1:NUMBER:����� ���������", "2:NAME:������������", "4:AREA:�������, �.��."), "TornadoRooms"
    report_formats.Add Array("1:NUMBER", "2:NAME", "4:AREA"), "TornadoDoors"
    report_formats.Add Array("1:NUMBER", "2:NAME", "4:AREA"), "TornadoWindows"
    cSettings.Add report_formats, "Formats"
    
    Dim report_headings As Collection
    Set report_headings = New Collection
    report_headings.Add "����������� ���������", "TornadoRooms"
    report_headings.Add "����������� ������� �������", "TornadoDoors"
    report_headings.Add "����������� ������� �������", "TornadoWindows"
    cSettings.Add report_headings, "Captions"
    
icons = Array( _
    Array("������� &�����", 231, "UI.set_location"), _
    Array("�������� &�����", 37, "UI.update"), _
    Array("�� &�������", 3633, "UI.popup"), _
    Array("�&�����", 280, "UI.snapshot") _
    )
cSettings.Add icons, "Icons"


templates = Array( _
    Array("��������� ���������", "TornadoRooms", 1, 8) _
    )
cSettings.Add templates, "Templates"
    
    
    cSettings.Add "BetterReports", "ToolbarName"
ElseIf i = 2 Then
    connection_filenames = Array("Test Set 2\mindata.csv")
    cSettings.Add connection_filenames, "Filenames"
End If

End Sub

