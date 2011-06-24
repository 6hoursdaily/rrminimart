VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sales_reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim toggleRecords As Boolean
Dim strSalesStatus As String

Private Sub Form_Load()
  
  Me.RecordSource = "SELECT * FROM SalesDtls WHERE Status = 'REG' AND DtlsDate LIKE '" & strSQL & "*'"
  Me.reportdate = strSQL
  Me.reportdate.Visible = True
  Me.headerusername.Caption = "by: " & strUser
  Me.footerusername.Caption = "by: " & strUser
  toggleRecords = False
  strSalesStatus = "REG"

End Sub
Private Sub btn_print_Click()
On Error GoTo Err_btn_print_Click

    
  'DoCmd.PrintOut
  DoCmd.OpenForm "sales_reports", acPreview

Exit_btn_print_Click:
    Exit Sub

Err_btn_print_Click:
    MsgBox Err.Description
    Resume Exit_btn_print_Click
    
End Sub

Private Sub button_toggle_view_Click()
On Error GoTo Err_button_toggle_view_Click

  If toggleRecords = False Then
  
    strSalesStatus = "VOD"
    toggleRecords = True
    Me.button_toggle_view.Caption = "See SALES"
  
  Else
    
    strSalesStatus = "REG"
    toggleRecords = False
    Me.button_toggle_view.Caption = "See VOID"
    
  End If

  Me.RecordSource = "SELECT * FROM SalesDtls WHERE Status = '" & strSalesStatus & "' AND DtlsDate LIKE '" & strSQL & "*'"


Exit_button_toggle_view_Click:
    Exit Sub

Err_button_toggle_view_Click:
    MsgBox Err.Description
    Resume Exit_button_toggle_view_Click
    
End Sub
Private Sub button_export_Click()
On Error GoTo Err_button_export_Click

  Me.RecordSource = "SELECT * FROM SalesDtls WHERE Status = '" & strSalesStatus & "' AND DtlsDate LIKE '" & strSQL & "*'"

Exit_button_export_Click:
    Exit Sub

Err_button_export_Click:
    MsgBox Err.Description
    Resume Exit_button_export_Click
    
End Sub

Private Sub btn_view_item_Click()
  
  Me.RecordSource = "SELECT * FROM SalesDtls WHERE Status = '" & strSalesStatus & "' AND DtlsDate LIKE '" & strSQL & "*' AND ProdCode = '" & Me.ProdCode & "'"

End Sub

Private Sub button_show_all_Click()
On Error GoTo Err_button_show_all_Click

  Me.RecordSource = "SELECT * FROM SalesDtls WHERE Status = 'REG' AND DtlsDate LIKE '" & strSQL & "*'"

Exit_button_show_all_Click:
    Exit Sub

Err_button_show_all_Click:
    MsgBox Err.Description
    Resume Exit_button_show_all_Click
    
End Sub
