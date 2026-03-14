VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKeyTool 
   Caption         =   "HubSpot API Token"
   ClientHeight    =   1635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   OleObjectBlob   =   "frmKeyTool.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmKeyTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

End Sub

Private Sub btnCancel_Click()
    If MsgBox("Update / Save API token?", vbYesNo + vbQuestion, pcMsgBoxTitel) = vbNo Then
        Unload Me
    Else
        prpRegKeyValue("API-KEY") = txtAPIKey
        Unload Me
    End If
End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()

Dim strAPIKey As String

strAPIKey = prpRegKeyValue("API-KEY")

    If strAPIKey = "" Then
        Me.txtAPIKey = "- ('empty' or no token saved) - "
    Else
        Me.txtAPIKey = strAPIKey
    
    End If

End Sub
