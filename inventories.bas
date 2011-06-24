VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_inventories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Click()
    Me.boxes.SetFocus
End Sub

Private Sub Form_Load()
    
    'strSQLSearch = "SELECT sku_master.ID, sku_master.sku, sku_master.product_name, sku_master.uom, (SELECT inventories.physical_inventory_ID FROM inventories WHERE inventories.physical_inventory_ID = " & strPIID & ") AS physical_inventory_ID, (SELECT inventories.boxes FROM inventories WHERE inventories.physical_inventory_ID = " & strPIID & " AND inventories.sku_master_id = sku_master.ID) AS boxes, (SELECT inventories.pieces FROM inventories WHERE inventories.physical_inventory_ID = " & strPIID & " AND inventories.sku_master_id = sku_master.ID) AS pieces, (SELECT inventories.total FROM inventories WHERE inventories.physical_inventory_ID = " & strPIID & " AND inventories.sku_master_id = sku_master.ID) AS total FROM sku_master LEFT JOIN inventories ON sku_master.ID = inventories.sku_master_id"
    strSQLSearch = "SELECT sku_master.ID, sku_master.sku, sku_master.product_name, sku_master.uom, inventories.physical_inventory_ID, inventories.boxes, inventories.pieces, inventories.total FROM sku_master LEFT JOIN inventories ON sku_master.ID = inventories.sku_master_id WHERE inventories.physical_inventory_ID = " & strPIID & " OR inventories.physical_inventory_ID IS NULL"
    Me.RecordSource = strSQLSearch
    Me.input_name.SetFocus

End Sub
Private Sub button_search_Click()
On Error GoTo Err_button_search_Click

    strSQLBase = "SELECT sku_master.ID, sku_master.sku, sku_master.product_name, sku_master.uom, inventories.physical_inventory_ID, inventories.boxes, inventories.pieces, inventories.total FROM sku_master LEFT JOIN inventories ON sku_master.ID = inventories.sku_master_id"
    'strSQLBase = "SELECT sku_master.ID, sku_master.sku, sku_master.product_name, sku_master.uom, inventories.physical_inventory_ID, inventories.boxes, inventories.pieces, inventories.total FROM sku_master LEFT JOIN inventories ON sku_master.ID = inventories.sku_master_id WHERE inventories.physical_inventory_ID = " & strPIID & " OR inventories.physical_inventory_ID IS NULL"
    
    objSearchForm = "inventories"
    objInputCode = "input_code"
    objInputName = "input_name"
    colS1 = "sku_master.sku"
    colS2 = "sku_master.product_name"
    
    search_records
        
    Me.boxes.SetFocus

Exit_button_search_Click:
    Exit Sub

Err_button_search_Click:
    MsgBox Err.Description
    Resume Exit_button_search_Click
    
End Sub

Private Sub button_enter_Click()
On Error GoTo Err_button_enter_Click

    Set db = CurrentDb() 'You may also use: OpenDatabase("MyDatabase.mdb")
    strSQL = "SELECT * FROM inventories WHERE sku_master_id = " & Me.ID & " AND physical_inventory_ID = " & strPIID
    Set rst = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    
    'check if the record already exists
    If rst.BOF And rst.EOF Then
    
        rst.AddNew
    
    Else
        
        rst.MoveFirst
        rst.Edit
        
    End If
    
    rst!physical_inventory_ID = strPIID
    rst!sku_master_id = Me.ID
    rst!boxes = Me.boxes
    rst!pieces = Me.pieces
    rst!total = Me.total
    rst.Update
    rst.Close
    
    Me.RecordSource = strSQLSearch
    Me.boxes = 0
    Me.pieces = 0
    Me.total = 0

    Me.input_name.SetFocus
        
Exit_button_enter_Click:
    Set rst = Nothing    'Deassign all objects.
    Set db = Nothing
    Exit Sub

Err_button_enter_Click:
    
    MsgBox Err.Description
    Resume Exit_button_enter_Click
    
End Sub

Private Sub pieces_Exit(Cancel As Integer)
    If IsNull(Me.boxes) Or Me.boxes = "" Then
        Me.total = Me.pieces
    Else
        Me.total = (Me.uom * Me.boxes) + Me.pieces
    End If
    
End Sub
Private Sub button_choose_Click()
On Error GoTo Err_button_choose_Click

    strSQLSearch = strSQLBase & " WHERE sku_master.sku = '" & Me.sku & "'"
    Form_inventories.RecordSource = strSQLSearch
    Me.boxes.SetFocus

Exit_button_choose_Click:
    Exit Sub

Err_button_choose_Click:
    MsgBox Err.Description
    Resume Exit_button_choose_Click
    
End Sub
Private Sub button_go_to_main_menu_Click()
On Error GoTo Err_button_go_to_main_menu_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "physical_inventory"
    DoCmd.Close acForm, "inventories"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_button_go_to_main_menu_Click:
    Exit Sub

Err_button_go_to_main_menu_Click:
    MsgBox Err.Description
    Resume Exit_button_go_to_main_menu_Click
    
End Sub

Private Sub product_name_Click()
    Me.boxes.SetFocus
End Sub
