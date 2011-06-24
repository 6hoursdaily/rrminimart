VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_warehouse_product_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  
  Me.RecordSource = strSQLEdit
  
  'call details view
  strSQLDetail = "SELECT TOP 1 * FROM item_master_query WHERE ProdCode = '" & Me.ProdCode & "'"
  Me.viewer.SourceObject = "warehouse_product_detail_edit"
  
End Sub

Private Sub btn_change_Click()
  
  strSQL = Me.ProdCode
  DoCmd.OpenForm "warehouse_product_name_edit"
  'MsgBox "should open form to change name", vbOKOnly, "dev note"
  
End Sub

Private Sub button_update_Click()
On Error GoTo Err_button_update_Click
    
    'DoCmd.Save acForm, "warehouse_product_detail_edit"
    DoCmd.Save
    
    'Call format_viewer("_sub_product_search", 0)
    'Forms(objForm).Controls(objControl).Form.ScrollBars = 2

Exit_button_update_Click:
    Exit Sub

Err_button_update_Click:
    MsgBox Err.Description
    Resume Exit_button_update_Click
    
End Sub


