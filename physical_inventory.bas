VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_physical_inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub button_create_Click()
On Error GoTo Err_button_create_Click

    Me.User_ID = strUserID

    'save the physical inventory record
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    DoCmd.OpenForm "inventories", acNormal
    
    strPIID = Me.ID
    DoCmd.Close acForm, "physical_inventory"

    'Form_inventories.RecordSource = "SELECT * FROM query_inventories" 'WHERE physical_inventory_ID = " & Me.ID
    'Form_inventories.RecordSource = "SELECT sku_master.ID, sku_master.sku, sku_master.product_name, sku_master.uom, (SELECT inventories.physical_inventory_ID FROM inventories WHERE inventories.physical_inventory_ID = 24) AS physical_inventory_ID, (SELECT inventories.boxes FROM inventories WHERE inventories.physical_inventory_ID = 24) AS boxes, (SELECT inventories.pieces FROM inventories WHERE inventories.physical_inventory_ID = 24 AND inventories.sku_master_id = sku_master.ID) AS pieces, (SELECT inventories.total FROM inventories WHERE inventories.physical_inventory_ID = 24 AND inventories.sku_master_id = sku_master.ID) AS total FROM sku_master LEFT JOIN inventories ON sku_master.ID = inventories.sku_master_id"
    
Exit_button_create_Click:
    Exit Sub

Err_button_create_Click:
    MsgBox Err.Description
    Resume Exit_button_create_Click
    
End Sub

Private Sub Form_Load()
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    Me.Date.SetFocus
End Sub
Private Sub button_return_Click()
On Error GoTo Err_button_return_Click

    DoCmd.Close
    DoCmd.OpenForm "_main_menu", acNormal

Exit_button_return_Click:
    Exit Sub

Err_button_return_Click:
    MsgBox Err.Description
    Resume Exit_button_return_Click
    
End Sub
