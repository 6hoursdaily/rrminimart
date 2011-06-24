VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_warehouse_item_maintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Load()
  
  strOrderBy = " ORDER BY ProdNameLong"
  objControl = "viewer"
  strSQLSearch = ""
  'Me.SetFocus
  'Call toggle_viewer(True)
  Call format_viewer("_sub_product_search", 0)
  
End Sub

Private Sub button_import_Click()
On Error GoTo Err_button_import_Click

    subformwidth = Form__sub_import_items.Width
    subformheight = Form__sub_import_items.Detail.Height
    
    Call format_viewer("_sub_import_items", 1)

Exit_button_import_Click:
    Exit Sub

Err_button_import_Click:
    MsgBox Err.Description
    Resume Exit_button_import_Click
    
End Sub
Private Sub button_edit_Click()
On Error GoTo Err_button_edit_Click

    subformwidth = Form__sub_product_search.Width + 720
    subformheight = 5000
    
    strSQLSearch = "SELECT * FROM Products" + strOrderBy
    Call format_viewer("_sub_product_search", 1)
    Forms(objForm).Controls(objControl).Form.ScrollBars = 2

Exit_button_edit_Click:
    Exit Sub

Err_button_edit_Click:
    MsgBox Err.Description
    Resume Exit_button_edit_Click
    
End Sub
Private Sub button_create_Click()
On Error GoTo Err_button_create_Click


    subformwidth = Form__sub_create_product.Width + 720
    subformheight = 5000
    
    Call format_viewer("_sub_create_product", 1)
    Forms(objForm).Controls(objControl).Form.ScrollBars = 2

Exit_button_create_Click:
    Exit Sub

Err_button_create_Click:
    MsgBox Err.Description
    Resume Exit_button_create_Click
    
End Sub
Private Sub button_go_back_Click()
On Error GoTo Err_button_go_back_Click


    DoCmd.Close
    DoCmd.OpenForm "_main_menu", acNormal
    objControl = "submenu"

Exit_button_go_back_Click:
    Exit Sub

Err_button_go_back_Click:
    MsgBox Err.Description
    Resume Exit_button_go_back_Click
    
End Sub
