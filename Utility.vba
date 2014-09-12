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
Function ParseLineToArray(sInput As String, m_Delim As String, _
                                  m_TextIdentifier As String) As Variant
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

         ParseLineToArray = sArr()
      End If 'has any quoted text
   End If 'parseable

End Function

