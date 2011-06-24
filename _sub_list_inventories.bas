VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__sub_list_inventories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub button_edit_Click()
On Error GoTo Err_button_edit_Click

    Me.User_ID = strUserID
    strPIID = Me.ID
    
    'save the physical inventory record
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    DoCmd.OpenForm "inventories", acNormal
    
    DoCmd.Close acForm, "physical_inventory"

    'Form_inventories.RecordSource = "SELECT * FROM query_inventories" WHERE physical_inventory_ID = " & Me.ID

Exit_button_edit_Click:
    Exit Sub

Err_button_edit_Click:
    MsgBox Err.Description
    Resume Exit_button_edit_Click
    
End Sub
