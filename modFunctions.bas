Attribute VB_Name = "modFunctions"
Option Explicit

Public Function fncCustDocPropExitst(strCustDocPropName As String) As Boolean
'This function checks whether a specific CustomDocProperty exists or not.
Dim prop As DocumentProperty
Dim blnReturnvalue As Boolean
blnReturnvalue = False
For Each prop In ActiveDocument.CustomDocumentProperties
    If LCase(prop.Name) = LCase(strCustDocPropName) Then
        blnReturnvalue = True
        Exit For
    End If
Next
fncCustDocPropExitst = blnReturnvalue
End Function



Function fncNormalizePhone(ByVal strRaw As String) As String



'====================================================================
' fncNormalizePhone
'
' Takes a raw phone number string (Swiss format) and returns it
' formatted as "+41 XX XXX XX XX".
'
' Handles input like:
'   +41795076580
'   +41 79 507 65 80
'   044 403 28 86
'
' If formatting fails, returns the original string.
'
' Returns: Formatted phone string or original string on error
'====================================================================

    On Error GoTo ErrHandler
    
    Dim strClean As String       ' cleaned phone number (only digits + "+")
    Dim i As Long                ' loop counter
    Dim ch As String             ' single character
    
    ' --- Step 1: keep only digits and "+" ---
    strClean = ""
    For i = 1 To Len(strRaw)
        ch = Mid(strRaw, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "+" Then
            strClean = strClean & ch
        End If
    Next i
    
    ' --- Step 2: convert leading 0 to +41 for Swiss numbers ---
    If Left(strClean, 1) = "0" Then
        strClean = "+41" & Mid(strClean, 2)
    ElseIf Left(strClean, 1) <> "+" Then
        ' assume national number without leading 0, prepend +41
        strClean = "+41" & strClean
    End If
    
    ' --- Step 3: format as +41 XX XXX XX XX ---
    If Left(strClean, 3) = "+41" And Len(strClean) = 12 Then
        fncNormalizePhone = Left(strClean, 3) & " " & Mid(strClean, 4, 2) & " " & _
                         Mid(strClean, 6, 3) & " " & Mid(strClean, 9, 2) & " " & Mid(strClean, 11, 2)
    Else
        ' fallback: return original string if format not recognized
        fncNormalizePhone = strRaw
    End If
    
    Exit Function
    
ErrHandler:
    ' return original input if any error occurs
    fncNormalizePhone = strRaw
End Function






