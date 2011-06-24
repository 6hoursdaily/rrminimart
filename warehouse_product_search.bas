VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_warehouse_product_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    
  Me.label_product_code.Visible = False
  Me.label_product_name.Visible = False
  Me.ProdCode.Visible = False
  Me.ProdNameLong.Visible = False
    
End Sub

Private Sub button_edit_Click()
On Error GoTo Err_button_edit_Click
  
  objSubformControl = "info"
  
  strCode = Me.ProdCode

  Select Case stDocName
    Case "item_maintenance"
      strSQLEdit = "SELECT TOP 1 ProdCode, ProdNameLong FROM Products WHERE ProdCode = '" & strCode & "'"
      Call format_viewer("warehouse_product_view", 0, 1)
  
    Case "starting_inventory"
      strSQLEdit = "SELECT TOP 1 * FROM views_product_info WHERE ProdCode = '" & strCode & "'"
      Call format_viewer("warehouse_starting_inventory", 0, 1)
  End Select
  
Exit_button_edit_Click:
    Exit Sub

Err_button_edit_Click:
    MsgBox Err.Description
    Resume Exit_button_edit_Click
    
End Sub

Private Sub button_search_Click()
  
  'declare search
  strName = Me.input_string
  strSQLSearch = "SELECT ProdCode, ProdNameLong FROM Products WHERE ProdNameLong LIKE '*" & strName & "*'"
  
  Me.RecordSource = strSQLSearch
  
  'show fields
  Me.label_product_code.Visible = True
  Me.label_product_name.Visible = True
  Me.ProdCode.Visible = True
  Me.ProdNameLong.Visible = True
  
  'set the focus back to search box
  Me.input_string.SetFocus
  Me.input_string.SelStart = Len(Me.input_string)
  

End Sub
