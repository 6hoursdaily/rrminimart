VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_xsales_reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmd_tot_sales_Click()
On Error GoTo Err_cmd_tot_sales_Click

    Dim stDocName As String

    stDocName = "Total_sales"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_cmd_tot_sales_Click:
    Exit Sub

Err_cmd_tot_sales_Click:
    MsgBox Err.Description
    Resume Exit_cmd_tot_sales_Click
    
End Sub
Private Sub cmd_export_to_excel_Click()
On Error GoTo Err_cmd_export_to_excel_Click

    Dim stDocName As String

    stDocName = "export_total_sales"
    DoCmd.RunMacro stDocName

Exit_cmd_export_to_excel_Click:
    Exit Sub

Err_cmd_export_to_excel_Click:
    MsgBox Err.Description
    Resume Exit_cmd_export_to_excel_Click
    
End Sub
