Attribute VB_Name = "Utility"
Function get_cwd()
    get_cwd = ThisWorkbook.Path & "\"
End Function

Function get_fullname()
    get_fullname = ThisWorkbook.FullName
End Function

Function get_item_by_name(iterable As Object, name)
    For x = 1 To iterable.Count
        If name = iterable.Item(x).name Then
            Set get_item_by_name = iterable.Item(x)
            Exit For
        End If
    Next x
End Function

Function get_item_by_property(iterable As Object, PropertyName As String, value As Variant)
    Set get_item_by_property = Nothing
    For x = 1 To iterable.Count
        If value = CallByName(iterable.Item(x), PropertyName, VbGet) Then
            Set get_item_by_property = iterable.Item(x)
            Exit For
        End If
    Next x
End Function

'Split a string into an array based on a Delimiter and a Text Identifier
Function parse_csv_line(input_line As String, Optional delimeter As String = ",", Optional text_identifier As String = """") As Variant
Dim result() As String
Dim input_array() As String
Dim in_text_flag As Boolean
Dim schedule_flag_toggle As Boolean
Dim i As Long, n As Long
Dim tempstring As String, tmp As String
input_array() = Split(input_line, delimeter)

If UBound(input_array) = -1 Then
    parse_csv_line = Array()
    Exit Function
End If

first_init = True
For Each x In input_array

    If InStr(1, x, text_identifier) = 0 Then
        schedule_flag_toggle = False
    Else
        If in_text_flag Then
            'inside text, with identifier
            tempstring = Replace(x, text_identifier & text_identifier, "")
            If InStr(1, tempstring, text_identifier) <> 0 Then
                schedule_flag_toggle = True
                x = Left(x, Len(x) - 1)
            Else
                schedule_flag_toggle = False
            End If
        Else
            'out of text, with identifier
            tempstring = Replace(x, text_identifier & text_identifier, "")
            x = Right(x, Len(x) - 1)
            If InStr(1, tempstring, text_identifier) <> 0 And InStr(1, tempstring, text_identifier) = InStrRev(tempstring, text_identifier) Then
                schedule_flag_toggle = True
            Else
                x = Left(x, Len(x) - 1)
            End If
        End If
        x = Replace(x, text_identifier & text_identifier, text_identifier)
    End If
    
    If in_text_flag Then
        result(UBound(result)) = result(UBound(result)) & delimeter & x
    Else
        If first_init Then
            ReDim Preserve result(0)
            first_init = False
        Else
            ReDim Preserve result(UBound(result) + 1)
        End If
        result(UBound(result)) = x
    End If
    
    If schedule_flag_toggle Then
        in_text_flag = Not in_text_flag
        schedule_flag_toggle = False
    End If
Next x

parse_csv_line = result()
End Function

Function get_custom_property(PropertyName As String, Optional WhatWorkbook As Workbook) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetProperty
' This procedure returns the value of a DocumentProperty named in
' PropertyName. It will examine BuiltinDocumentProperties,
' or CustomDocumentProperties, or both. The parameters are:
'
'   PropertyName        The name of the property to return.
'
'   PropertySet         One of PropertyLocationBuiltIn,
'                       PropertyLocationCustom, or PropertyLocationBoth.
'                       This specifies the property set to search.
'
'   WhatWorkbook        A reference to the workbook whose properties
'                       are to be examined. If omitted or Nothing,
'                       ThisWorkbook is used.
'
' The function will return:
'
'   The value of property named by PropertyName, or
'
'   #VALUE if the PropertySet parameter is not valid (test with IsError), or
'
'   Null if the property could not be found (test with IsNull)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim WB As Workbook
Dim Props1 As Office.DocumentProperties
Dim Props2 As Office.DocumentProperties
Dim Prop As Office.DocumentProperty

'''''''''''''''''''''''''''''''''''''''''
' Set the workbook whose properties we
' will search.
'''''''''''''''''''''''''''''''''''''''''
If WhatWorkbook Is Nothing Then
    Set WB = ThisWorkbook
Else
    Set WB = WhatWorkbook
End If

Set Props1 = WB.CustomDocumentProperties

On Error Resume Next
'''''''''''''''''''''''''''''''''''''''''
' Search either BuiltIn or Custom.
'''''''''''''''''''''''''''''''''''''''''
Set Prop = Props1(PropertyName)
If Err.Number <> 0 Then
    ''''''''''''''''''''''''''''''''''
    ' Not found in one set. See if
    ' we need to look in the other.
    ''''''''''''''''''''''''''''''''''
    GetProperty = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Property found. Return the value.
''''''''''''''''''''''''''''''''''''
GetProperty = Prop.value

End Function

Function in_array(value, list_array() As String) As Integer
in_array = -1
For x = 0 To UBound(list_array)
    If value = list_array(x) Then
        in_array = x
        Exit Function
    End If
Next x
End Function

Function get_property_type(V As Variant) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetPropertyType
' This tests the data type of V and returns the appropriate Property type.
' Returns a member of the VbVarType group or NULL if an illegal type (e.g,
' an Object) is found. Be sure to test the return value with IsNull.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Select Case VarType(V)
    Case vbArray, vbDataObject, vbEmpty, vbError, vbNull, vbObject, _
        vbUserDefinedType, vbCurrency, vbDecimal
        ''''''''''''''''''''''''''''''''''''
        ' Illegal types. Return NULL.
        ''''''''''''''''''''''''''''''''''''
        get_property_type = Null
        Exit Function
    ''''''''''''''''''''''''''''''''''
    ' All numeric types are rolled up
    ' into Floats. Strings and Booleans
    ' get their own types.
    ''''''''''''''''''''''''''''''''''
    Case vbString
        get_property_type = msoPropertyTypeString
    Case vbBoolean
        get_property_type = msoPropertyTypeBoolean
    Case Else
        get_property_type = msoPropertyTypeFloat
End Select

End Function

'?? utf-8 ? Unicode
Function UTF8_16(s)
    UTF8_16 = ""
    Dim i, j, j2, ch, k1, k2, k3, m
    i = 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        j = CLng(Asc(ch))
        If j >= 128 Then
            If j < 224 Then
                '2 ?????
                k1 = j Mod 32
                i = i + 1
                ch = Mid(s, i, 1)
                j2 = CLng(Asc(ch))
                k2 = j2 Mod 64
                'ChrW - ?????? ?? UTF-16 ????????
                UTF8_16 = UTF8_16 & ChrW(k2 + k1 * 64)
            Else
                '3 ?????
                k1 = j Mod 16
                i = i + 1
                ch = Mid(s, i, 1)
                j2 = CLng(Asc(ch))
                k2 = j2 Mod 64
                i = i + 1
                ch = Mid(s, i, 1)
                j2 = CLng(Asc(ch))
                k3 = j2 Mod 64
                UTF8_16 = UTF8_16 & ChrW(k3 + (k2 + k1 * 64) * 64)
            End If
        Else
            UTF8_16 = UTF8_16 & ch
        End If
        i = i + 1
    Loop
End Function

Function match_fields(string_array() As Variant, index As Integer, column() As String) As Variant()

End Function

