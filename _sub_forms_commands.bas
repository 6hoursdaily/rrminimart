VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__sub_forms_commands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  Me.combo_department_name.RowSource = "SELECT DISTINCT name FROM departments ORDER by name"
End Sub

Private Sub button_physical_inventory_Click()
On Error GoTo Err_button_physical_inventory_Click

    strSQLWhere = Me.combo_department_name.Value
    stDocName = "physical_inventory"
    DoCmd.OpenReport stDocName, acPreview

Exit_button_physical_inventory_Click:
    Exit Sub

Err_button_physical_inventory_Click:
    MsgBox Err.Description
    Resume Exit_button_physical_inventory_Click
    
End Sub

