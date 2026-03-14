Attribute VB_Name = "modProperties"
Option Explicit

Property Get prpRegKeyValue(strKeyName As String) As String
'Key / sub key (official: "section") are fixed baked in:
'
'Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings\
'VB AND VBA PROGRAM SETTINGS
'+---HubspotConnector
'    +---API
'
prpRegKeyValue = GetSetting("HubspotConnector", "API", strKeyName, "")
End Property

Property Let prpRegKeyValue(strKeyName As String, strValue As String)
'
'Use
'prpRegKeyValue("Keyname")="Keyvalue"
'to test
'
'Key / sub key (official: "section") are fixed baked in:
'VB AND VBA PROGRAM SETTINGS
'+---HubspotConnector
'    +---API

    SaveSetting "HubspotConnector", "API", strKeyName, strValue
End Property

'Similar to a function, this property queries the name of a bookmark.
Property Get prpBookmark(strName As String) As String
    If ActiveDocument.Bookmarks.Exists(strName) Then
        prpBookmark = ActiveDocument.Bookmarks(strName).Range.Text
    Else
        prpBookmark = ""
    End If
End Property

'This property sets the content of a bookmark. As it disappears, it is immediately set again.
Property Let prpBookmark(strName As String, strText As String)

Dim rngprpBookmark As Range
        
    If ActiveDocument.Bookmarks.Exists(strName) Then
        Set rngprpBookmark = ActiveDocument.Bookmarks(strName).Range
        rngprpBookmark.Text = strText
        rngprpBookmark.Bookmarks.Add strName
    Else
        MsgBox "The bookmark '" & strName & "' is missing!", vbCritical, "Syntax Error"
    End If

End Property

Property Get CustDocProperties(strCustDocPropName As String) As String

CustDocProperties = ActiveDocument.CustomDocumentProperties(strCustDocPropName)

End Property


Property Let CustDocProperties(strCustDocPropName As String, strCustDocPrpContent As String)
With ActiveDocument.CustomDocumentProperties
    .Add Name:=strCustDocPropName, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=strCustDocPrpContent
End With
End Property

