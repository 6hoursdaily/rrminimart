VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_warehouse_starting_inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  Me.RecordSource = strSQLEdit
  
  strSQLInventory = "SELECT TOP 1 * FROM starting_inventory WHERE sku_master_id = '" & Me.ProdCode & "'"
  
  If IsNull(Me.unit_per_box) = True Then
    strInventoryCount = 0
  Else
    strInventoryCount = Me.unit_per_box
  End If
  
  'strCode = Me.ProdCode
  Me.viewer.SourceObject = "warehouse_starting_inventory_count"
End Sub

Private Sub product_name_Click()
  
  strSQLEdit = "SELECT TOP 1 ProdCode, ProdNameLong FROM Products WHERE ProdCode = '" & strCode & "'"
  Call format_viewer("warehouse_product_view", 0, 1)

End Sub

