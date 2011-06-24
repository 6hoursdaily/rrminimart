VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sales Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  
  Me.RecordSource = "SELECT * FROM SalesDtls WHERE Status = 'REG' AND DtlsDate LIKE '" & strSQL & "*'"
  DoCmd.RunMacro "export_form_data"
  DoCmd.Close
End Sub
