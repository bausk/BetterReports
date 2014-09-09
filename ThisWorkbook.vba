VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()
    UI.refresh_buttons
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    UI.remove_buttons
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    save_code_modules
End Sub

Private Sub Workbook_Deactivate()
    'UI.remove_buttons
End Sub
