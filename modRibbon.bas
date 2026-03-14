Attribute VB_Name = "modRibbon"
Option Explicit

Sub HubiCallback(control As IRibbonControl)
    Select Case control.ID
    
    Case "btnHubSpotContactsFromClipboard": PrepareDocAndHubspot
    Case "btnInfo": MsgBox pcMsgBoxTitel, vbInformation, "Information"
    Case "btnKeyTool": CallfrmKeyTool
    
    End Select

End Sub



Sub CallfrmKeyTool()
frmKeyTool.Show
End Sub
