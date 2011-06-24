VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__sub_audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub button_reports_Click()
On Error GoTo Err_button_reports_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "_sub_sales"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_button_reports_Click:
    Exit Sub

Err_button_reports_Click:
    MsgBox Err.Description
    Resume Exit_button_reports_Click
    
End Sub
