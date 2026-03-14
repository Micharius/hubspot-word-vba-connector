Attribute VB_Name = "modPrepareDoc"
Option Explicit

Sub PrepareDocAndHubspot()
    Dim objTbl As Table
    
    Dim strInputBox As String
    
    Dim objCell As Cell
    
    Dim objRngTitle As Object
    Dim objrngSubject As Object
    Dim objrngVersion As Object
    Dim objrngAuthor As Object
    Dim objRngDate As Object
    Dim objRngConCompany As Object
    Dim objRngCustAddress1 As Object
    Dim objRngCustAddress2 As Object
    Dim objRngConName As Object
    Dim objRngConEmail As Object
    Dim objRngConPhone As Object
    Dim objRngConMobile As Object
    Dim objRngCustWebAddress As Object
    
    Dim strRngTitle As String
    Dim strRngSubject As String
    Dim strRngVersion As String
    
    On Error GoTo ErrorHandler
    
    
    
    If prpRegKeyValue("API-KEY") = "" Then 'API-KEY is the name of the registry value
        
        If MsgBox("No HubSpot API token found - do you want to enter it?", vbQuestion + vbYesNo, pcMsgBoxTitel) = vbYes Then
           
           strInputBox = InputBox("Please paste your API token here", pcMsgBoxTitel)
        
            If strInputBox = "" Then
                MsgBox "This program is no working withouth a valid HubSpot API token - Abort!", vbInformation, pcMsgBoxTitel
                Exit Sub
            End If
        
           prpRegKeyValue("API-KEY") = strInputBox
        
        Else
            
            MsgBox "This program is no working withouth a valid HubSpot API token - Abort!", vbInformation, pcMsgBoxTitel
            Exit Sub
        
        End If
    
    End If
    
    frmTLContactImport.Show
    
    Exit Sub
    
ErrorHandler:
    

            MsgBox "Fatal error no" & Err.Number & "-" & Err.Description & " in PrepareDocAndHubspot", vbCritical, pcMsgBoxTitel



End Sub


Sub InsertBookmarkAtCursor(strBookmarkName As String)
    On Error GoTo ErrHandler

    ' Check if the bookmark name is valid
    If Trim(strBookmarkName) = "" Then Exit Sub

    ' Check if a bookmark with the same name already exists
    If ActiveDocument.Bookmarks.Exists(strBookmarkName) Then Exit Sub

    ' Insert the bookmark at the current selection
    ActiveDocument.Bookmarks.Add Name:=strBookmarkName, Range:=Selection.Range
    Exit Sub

ErrHandler:
    MsgBox "Fatal error in sub 'InsertBookmarkAtCursor'", vbCritical, pcMsgBoxTitel
End Sub


Sub KeyTool()

Dim strToken As String
Dim strNewToken As String

strToken = prpRegKeyValue("API-KEY") 'read value from registry

strNewToken = InputBox("View and edit your HubSpot API token", pcMsgBoxTitel, strToken)

If strToken = "" Then
    MsgBox "HubSpot API token removed or cancelled", vbInformation, pcMsgBoxTitel

    If strToken <> strNewToken Then
        prpRegKeyValue("API-KEY") = strNewToken
        MsgBox "HubSpot API token updated!", vbInformation, pcMsgBoxTitel
        Exit Sub
    
    Else
    
    
    End If

End If

End Sub



' Inserts a DOCPROPERTY field that displays the document property.
 'If wrapInCC = True, the field is packed into a text content control.
Sub InsertCorePropertyCC(objRange As Range, strPropName As String, strCaption As String)

    Dim oXmlPart As CustomXMLPart
    Dim oCC As ContentControl
    Dim strXPath As String
    Dim strNS As String
    Dim blnFound As Boolean
    Dim strTextContent As String

    blnFound = False
    
    'The text of the objRange has to be saved, it'll we be lost else
    strTextContent = objRange.Text

    ' core-properties XML-Part finden
    For Each oXmlPart In ActiveDocument.CustomXMLParts
        If InStr(1, oXmlPart.NamespaceURI, "/package/2006/metadata/core-properties", vbTextCompare) > 0 Then
            blnFound = True
            Exit For
        End If
    Next

    If Not blnFound Then
        MsgBox "no core-properties XML-Part found.", vbCritical
        Exit Sub
    End If

    ' define matching XPath
    Select Case LCase(strPropName)
        Case "title":            strXPath = "/cp:coreProperties/dc:title"
        Case "subject":          strXPath = "/cp:coreProperties/dc:subject"
        Case "author", "creator": strXPath = "/cp:coreProperties/dc:creator"
        Case "description":      strXPath = "/cp:coreProperties/dc:description"
        Case "keywords":         strXPath = "/cp:coreProperties/cp:keywords"
        Case "lastmodifiedby":   strXPath = "/cp:coreProperties/cp:lastModifiedBy"
        Case "created":          strXPath = "/cp:coreProperties/dcterms:created"
        Case "modified":         strXPath = "/cp:coreProperties/dcterms:modified"
        Case "contentstatus":    strXPath = "/cp:coreProperties/cp:contentStatus" 'Achtung, eingefügt wird das 'ContentStatus' mit Space dazwischen!
        Case Else
            MsgBox "Property '" & strPropName & "' wird nicht in core-properties gespeichert.", vbCritical
            Exit Sub
    End Select

    ' Namespaces
    strNS = "xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' " & _
            "xmlns:dc='http://purl.org/dc/elements/1.1/' " & _
            "xmlns:dcterms='http://purl.org/dc/terms/'"

    ' Insert Content-Control
    Set oCC = ActiveDocument.ContentControls.Add(wdContentControlText, objRange)
    oCC.Title = strCaption

    ' set XML-Mapping
    If Not oCC.XMLMapping.SetMapping(strXPath, strNS, oXmlPart) Then
        MsgBox "Mapping fehlgeschlagen.", vbCritical
    End If
    oCC.Range.Text = strTextContent

End Sub


