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

'Split a string into an array based on a Delimiter and a Text Identifier
Function parse_csv_line(sInput As String, Optional m_Delim As String = ",", _
                                  Optional m_TextIdentifier As String = """") As Variant
   'Dim vArr As Variant
   Dim sArr() As String
   Dim bInText As Boolean
   Dim i As Long, n As Long
   Dim sTemp As String, tmp As String

   If sInput = "" Or InStr(1, sInput, m_Delim) = 0 Then
      'zero length string, or delimiter not present
      'dump all input into single-element array (minus Text Identifier)
      ReDim sArr(0)
      sArr(0) = Replace(sInput, m_TextIdentifier, "")
      ParseLineToArray = sArr()
   Else
      If InStr(1, sInput, m_TextIdentifier) = 0 Then
         'no text identifier so just split and return
         sArr() = Split(sInput, m_Delim)
         ParseLineToArray = sArr()
      Else
         'found the text identifier, so do it the long way
         bInText = False
         sTemp = ""
         n = 0

         For i = 1 To Len(sInput)
            tmp = Mid(sInput, i, 1)
            If tmp = m_TextIdentifier Then
               'just toggle the flag - don't add to string
               bInText = Not bInText
            Else
               If tmp = m_Delim Then
                  If Not bInText Then
                     'delimiter not within quoted text, so add next array member
                     ReDim Preserve sArr(n)
                     sArr(n) = sTemp
                     sTemp = ""
                     n = n + 1
                  Else
                     sTemp = sTemp & tmp
                  End If
               Else
                  sTemp = sTemp & tmp
               End If           'character is a delimiter
            End If              'character is a quote marker
         Next i

         ReDim Preserve sArr(n)
         sArr(n) = sTemp

         parse_csv_line = sArr()
      End If 'has any quoted text
   End If 'parseable

End Function

'Split a string into an array based on a Delimiter and a Text Identifier
Function parse_csv_line_2(input_line As String, Optional delimeter As String = ",", Optional text_identifier As String = """") As Variant
Dim result() As String
Dim input_array() As String
Dim in_text_flag As Boolean
Dim schedule_flag_toggle As Boolean
Dim i As Long, n As Long
Dim tempstring As String, tmp As String
input_array() = Split(input_line, delimeter)

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

'For i = 0 To UBound(result)
'    result(i) = trim_ends(result(i))
'Next i

parse_csv_line_2 = result()
End Function

Function escape_outtext(line, escape As String) As String
    tempstring = Replace(line, """""", "")
    If InStr(1, tempstring, """") = 1 Then
        
    If Left(line, 1) = """" Then
        
    End If
    
    escape_right = StrReverse(Replace(StrReverse(line), escape & escape, escape))
End Function


Function escape_left(line, escape As String) As String
    escape_left = Replace(line, escape & escape, escape)
End Function

Function escape_right(line, escape As String) As String
    tempstring = Replace(line, """""", "")
    If InStr(1, tempstring, """") = 1 Then
        
    If Left(line, 1) = """" Then
        line = Right(line, Len(line) - 1)
    End If
    
    escape_right = StrReverse(Replace(StrReverse(line), escape & escape, escape))
End Function


Function trim_ends(line) As String
If Left(line, 1) = """" Then
    line = Right(line, Len(line) - 1)
    If Right(line, 1) = """" Then
        line = Left(line, Len(line) - 1)
    End If
End If
trim_ends = line
End Function
