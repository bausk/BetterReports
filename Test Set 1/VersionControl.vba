Attribute VB_Name = "VersionControl"
Sub save_code_modules()
    
    'This code Exports all VBA modules
    Dim i%, sName$
    Dim FolderName As String
    FolderName = Utility.get_cwd()
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.Name
                .VBComponents(i%).Export FolderName & sName$ & ".vba"
            End If
        Next i
    End With
End Sub

