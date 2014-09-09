Attribute VB_Name = "Utility"
Function get_cwd()
    get_cwd = ThisWorkbook.Path & "\"
End Function

Function get_fullname()
    get_fullname = ThisWorkbook.FullName
End Function


Function get_item_by_name(iterable As Object, name As String)
    For x = 1 To iterable.Count
        If name = iterable.Item(x).name Then
            Set get_item_by_name = iterable.Item(x)
            Exit For
        End If
    Next x
End Function

Function get_item_by_property(iterable As Object, propertyName As String, value As Variant)
    Set get_item_by_property = Nothing
    For x = 1 To iterable.Count
        If value = CallByName(iterable.Item(x), propertyName, VbGet) Then
            Set get_item_by_property = iterable.Item(x)
            Exit For
        End If
    Next x
End Function

