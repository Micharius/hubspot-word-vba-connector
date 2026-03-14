VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTLContactImport 
   Caption         =   "Hubspot API Connector - Select your Contact"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   OleObjectBlob   =   "frmTLContactImport.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTLContactImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit




Private Sub btnReload_Click()
UserForm_Initialize
End Sub



Private Sub lstTLContact_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'Enter your own company url if you want to be able to doubleclick on an entry and open the contact in your webbrowser

'    Dim strHubSpotURL As String
'    Dim lngiRow As Long
'    Dim strContactID As String
'
'    strHubSpotURL = "https://app.hubspot.com/contacts/XXXXX/record/0-1"
'
'    lngiRow = Me.lstTLContact.ListIndex
'
'        ' Check if first row was selected (column heads, no actual data)
'        If lngiRow = 0 Then
'            MsgBox "Invalid row, please don't select title row", vbExclamation, pcMsgBoxTitel
'            Exit Sub
'        End If
'
'    'Use Column 11 (it start at 0) to get the contact ID
'    strContactID = Mid(Me.lstTLContact.List(lngiRow, 9), 5) 'Due to an ugly workaround, the zip code and Contact ID is combined in a string, therefore this cod
'
'    strHubSpotURL = strHubSpotURL & "/" & strContactID
'
'    ActiveDocument.FollowHyperlink Address:=strHubSpotURL
    
End Sub

Private Sub UserForm_Initialize()

On Error GoTo ErrorHandler


    Dim strInput As String
    Dim arrName As Variant
    Dim strVorname As String
    Dim strNachname As String
    Dim objResults As Object
    Dim objItem As Object
    Dim objHttp2 As Object, strResp2 As String, assocJson As Object, objAssoc As Object
    Dim objHttp3 As Object, strResp3 As String, objJson3 As Object
    Dim iRow As Long
    Dim intZ As Integer
    
    ' === Access Token ===
    Dim strToken As String
    strToken = prpRegKeyValue("API-KEY") 'by this point, we already have verified that the key exists!
    
    ' === User Input ===
    strInput = InputBox("Please enter name and first name separated by space:", "Search contact")
    If strInput = "" Then
        Unload Me
        End
    End If
    
    arrName = Split(Trim(strInput), " ")
    If UBound(arrName) < 1 Then
        MsgBox "Please enter name and first Name separated by space", vbExclamation
        UserForm_Initialize
    End If
    
    strVorname = arrName(1)
    strNachname = arrName(0)
    
    
    Set objResults = fncSearchHubspotContacts(strVorname, strNachname, strToken)
    
    If objResults.Count = 0 Then 'switch First/LastName and make another pass!
        strVorname = arrName(0)
        strNachname = arrName(1)
        Set objResults = fncSearchHubspotContacts(strVorname, strNachname, strToken)
        
        If objResults.Count = 0 Then
            MsgBox "Sorry, no contacts found!", vbExclamation, "Abort"
            Exit Sub
        End If
                
    End If
       
    ' === ListBox füllen ===
    Me.lstTLContact.Clear
    iRow = 0
    
    For Each objItem In objResults
        Dim strContactID As String
        Dim strFirstName As String
        Dim strLastName As String
        Dim strEmail As String
        Dim strAddress As String
        Dim strZipCode As String
        Dim strCity As String
        Dim strPhone As String
        Dim strMobilePhone As String
        Dim strCompany As String
        
        strContactID = objItem("id")
        strFirstName = fncNz(objItem("properties")("firstname"))
        strLastName = fncNz(objItem("properties")("lastname"))
        strEmail = fncNz(objItem("properties")("email"))
        strAddress = fncNz(objItem("properties")("address"))
        strCity = fncNz(objItem("properties")("city"))
        strPhone = fncNz(objItem("properties")("phone"))
        strMobilePhone = fncNz(objItem("properties")("mobilephone"))
        
        
        ' === Associations: Contact -> Company ===
        Set objHttp2 = CreateObject("MSXML2.XMLHTTP")
        objHttp2.Open "GET", "https://api.hubapi.com/crm/v4/objects/contacts/" & strContactID & "/associations/companies", False
        objHttp2.setRequestHeader "Authorization", "Bearer " & strToken
        objHttp2.send
        strResp2 = objHttp2.responseText
        Set assocJson = JsonConverter.ParseJson(strResp2)
        
        If assocJson.Exists("results") Then
        strCompany = ""
        
            For Each objAssoc In assocJson("results")
                Dim strCompanyId As String
                Dim strDomain As String
                
                strCompanyId = objAssoc("toObjectId")
                
                Set objHttp3 = CreateObject("MSXML2.XMLHTTP")
                objHttp3.Open "GET", "https://api.hubapi.com/crm/v3/objects/companies/" & strCompanyId & "?properties=name,city,zip,address,domain", False
                objHttp3.setRequestHeader "Authorization", "Bearer " & strToken
                objHttp3.send
                strResp3 = objHttp3.responseText
                Set objJson3 = JsonConverter.ParseJson(strResp3)
                
                strCompany = fncNz(objJson3("properties")("name"))
                strDomain = fncNz(objJson3("properties")("domain"))
                
                ' if city, zip code and address couldn't extracted from the contact we try to take it from the company
                
                If strCity = "" Then
                    strCity = fncNz(objItem("properties")("city"))
                End If
                
                If strZipCode = "" Then
                    strZipCode = fncNz(objItem("properties")("zip"))
                End If
                
                If strAddress = "" Then
                    strAddress = fncNz(objItem("properties")("address"))
                End If
                
                
                Exit For 'Take only 1st company
            Next objAssoc
                
        End If
        
        ' === ListBox: Add rows ===
              
    Me.lstTLContact.AddItem strLastName, iRow        '1
    Me.lstTLContact.Column(1, iRow) = strFirstName   '2
    Me.lstTLContact.Column(2, iRow) = strEmail       '3
    Me.lstTLContact.Column(3, iRow) = strCompany     '4
    Me.lstTLContact.Column(4, iRow) = strCity        '5
    Me.lstTLContact.Column(5, iRow) = strDomain      '6, hidden
    Me.lstTLContact.Column(6, iRow) = strAddress     '7, hidden
    Me.lstTLContact.Column(7, iRow) = strMobilePhone '8, hidden
    Me.lstTLContact.Column(8, iRow) = strPhone       '9, hidden
    Me.lstTLContact.Column(9, iRow) = strZipCode & strContactID       '10, hidden - ugly workaround due to Word's inability to address more than 10 columsn in a list box :-(
        
        
        iRow = iRow + 1
    Next objItem
    
    ' === Add title row ===
    Me.lstTLContact.AddItem "Name", 0
    Me.lstTLContact.Column(1, 0) = "First name"
    Me.lstTLContact.Column(2, 0) = "Email"
    Me.lstTLContact.Column(3, 0) = "Company"
    Me.lstTLContact.Column(4, 0) = "City"
Exit Sub


ErrorHandler:

    If Err.Number = 91 Then
        MsgBox "Error " & Err.Number & " - Your API token might be invalid!", vbCritical, pcMsgBoxTitel
    
    ElseIf Err.Number = -214669721 Then
    
        MsgBox "Error " & Err.Number & " - HubSpot API couldn't be reached!", vbCritical, pcMsgBoxTitel
    Else
        MsgBox "Error " & Err.Number & "-" & Err.Description & "' in Form Sub lstTLContact - abort!", vbCritical, "Aliens in sector"
    End If

Exit Sub

End Sub

Private Function fncSearchHubspotContacts(strVorname As String, strNachname As String, strToken As String) As Object

Dim objHttp As Object
Dim strResponse As String
Dim objJson As Object
Dim objResults As Object

    Dim strBody As String
    strBody = "{""filterGroups"":[{""filters"":[{""propertyName"":""firstname"",""operator"":""EQ"",""value"":""" & strVorname & """}, " & _
    "{""propertyName"":""lastname"",""operator"":""EQ"",""value"":""" & strNachname & """}]}]," & _
    """properties"":[""firstname"",""lastname"",""email"",""address"",""zip"",""city"",""mobilephone"",""phone""],""limit"":50}"
    
    Set objHttp = CreateObject("MSXML2.XMLHTTP")
    objHttp.Open "POST", "https://api.hubapi.com/crm/v3/objects/contacts/search", False
    objHttp.setRequestHeader "Authorization", "Bearer " & strToken
    objHttp.setRequestHeader "Content-Type", "application/json"
    objHttp.send strBody
    
    strResponse = objHttp.responseText
    Set objJson = JsonConverter.ParseJson(strResponse)
    
    If Not objJson.Exists("results") Then Exit Function
    Set fncSearchHubspotContacts = objJson("results")

End Function

Private Sub bntTLContEinfuegen_Click()

On Error GoTo ErrorHandler

    Dim strLastName As String
    Dim strFirstName As String
    Dim strEmail As String
    Dim strCompany As String
    Dim strCity As String
    Dim strDomain As String
    Dim strAddress As String
    Dim strMobilePhone As String
    Dim strPhone As String
    Dim strZipCode As String
    Dim lngiRow As Long

    ' Check if even something selected
    If Me.lstTLContact.ListIndex = -1 Then
        MsgBox "Please select a contact!", vbExclamation, pcMsgBoxTitel
        Exit Sub
    End If
    
    lngiRow = Me.lstTLContact.ListIndex
    
    ' Check if first row was selected (column heads, no actual data)
    If lngiRow = 0 Then
        MsgBox "Invalid row – please select a contact!.", vbExclamation, pcMsgBoxTitel
        Exit Sub
    End If

    ' Fill variables with columns valures
    strLastName = Me.lstTLContact.Column(0, lngiRow)
    strFirstName = Me.lstTLContact.Column(1, lngiRow)
    strEmail = Me.lstTLContact.Column(2, lngiRow)
    strCompany = Me.lstTLContact.Column(3, lngiRow)
    strCity = Me.lstTLContact.Column(4, lngiRow)
    strDomain = Me.lstTLContact.Column(5, lngiRow)
    strAddress = Me.lstTLContact.Column(6, lngiRow)
    strMobilePhone = Me.lstTLContact.Column(7, lngiRow)
    strPhone = Me.lstTLContact.Column(8, lngiRow)
    strZipCode = Left(Me.lstTLContact.Column(9, lngiRow), 4) '4 digits from right due to a miserabel workaround used in the Form_initalize routine...

'I use this for my templates
    'ActiveDocument.BuiltInDocumentProperties("Keywords") = strCompany
    
'    If fncCustDocPropExitst("prpCustCompany") = True Then
'        ActiveDocument.CustomDocumentProperties("prpCustCompany") = strCompany
'        ActiveDocument.Fields.Update
'    End If
    
    prpBookmark("bkmConName") = strFirstName & " " & strLastName
    prpBookmark("bkmConEmail") = strEmail
    prpBookmark("bkmCustAddress1") = strAddress
    prpBookmark("bkmCustAddress2") = strZipCode & " " & strCity
    prpBookmark("bkmCustWebAddress") = "www." & strDomain
    
    If strPhone <> "" Then
        prpBookmark("bkmConPhone") = fncNormalizePhone(strPhone)
    Else
        prpBookmark("bkmConPhone") = "-"
    End If
    
    
    If strMobilePhone <> "" Then
        prpBookmark("bkmConMobile") = fncNormalizePhone(strMobilePhone)
    Else
        prpBookmark("bkmConMobile") = "-"
    End If
    
    
    
    'refresh fields because othe cust. DocProperties
    'ActiveDocument.Fields.Update
    
    Unload Me
    
    Exit Sub

ErrorHandler:

MsgBox "Fehler '" & Err.Number & "-" & Err.Description & "' in Sub bntTLContEinfuegen - abort!", vbCritical, "Game over"

End Sub



Private Sub btnTLConImportCancel_Click()
Unload Me
End Sub

Private Function fncNz(v As Variant, Optional strDefault As String = "") As String
    If IsNull(v) Or v = vbNullString Then
        fncNz = strDefault
    Else
        fncNz = CStr(v)
    End If
End Function

