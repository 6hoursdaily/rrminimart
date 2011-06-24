VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_warehouse_starting_inventory_count"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  Me.RecordSource = strSQLInventory
  
  If Me.boxes > 0 Or Me.pieces > 0 Then
    Me.total = (strInventoryCount * Me.boxes) + Me.pieces
  Else
    Me.total = Me.pieces
  End If

End Sub

Private Sub boxes_LostFocus()
  If Me.boxes > 0 Then
    Me.total = strInventoryCount * Me.boxes
  End If
End Sub

Private Sub pieces_LostFocus()
  If Me.boxes > 0 Or Me.pieces > 0 Then
    Me.total = (strInventoryCount * Me.boxes) + Me.pieces
  Else
    Me.total = Me.pieces
  End If
End Sub

Private Sub btn_save_Click()
  DoCmd.Save
End Sub

