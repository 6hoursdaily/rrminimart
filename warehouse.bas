VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_warehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  objSubformControl = "viewer"
End Sub

Private Sub button_itemmaintenance_Click()
On Error GoTo Err_button_itemmaintenance_Click

  objSubformControl = "viewer"
  stDocName = "item_maintenance"
  Call format_viewer("warehouse_product_search", 0, 1)

Exit_button_itemmaintenance_Click:
    Exit Sub

Err_button_itemmaintenance_Click:
    MsgBox Err.Description
    Resume Exit_button_itemmaintenance_Click
    
End Sub
Private Sub button_startinginventory_Click()
On Error GoTo Err_button_startinginventory_Click

  objSubformControl = "viewer"
  stDocName = "starting_inventory"
  Call format_viewer("warehouse_product_search", 0, 1)

Exit_button_startinginventory_Click:
    Exit Sub

Err_button_startinginventory_Click:
    MsgBox Err.Description
    Resume Exit_button_startinginventory_Click
    
End Sub
Private Sub button_physicalinventory_Click()
On Error GoTo Err_button_physicalinventory_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "physical_inventory"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_button_physicalinventory_Click:
    Exit Sub

Err_button_physicalinventory_Click:
    MsgBox Err.Description
    Resume Exit_button_physicalinventory_Click
    
End Sub
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
Private Sub btn_departments_Click()
On Error GoTo Err_btn_departments_Click

    Call open_Form("department_manager")

Exit_btn_departments_Click:
    Exit Sub

Err_btn_departments_Click:
    MsgBox Err.Description
    Resume Exit_btn_departments_Click
    
End Sub
Private Sub button_forms_Click()
On Error GoTo Err_button_forms_Click

    objSubformControl = "viewer"
    Call format_viewer("_sub_forms_commands", 0, 1)

Exit_button_forms_Click:
    Exit Sub

Err_button_forms_Click:
    MsgBox Err.Description
    Resume Exit_button_forms_Click
    
End Sub

