VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_warehouse_product_name_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim checkInputChange As Boolean

Private Sub Form_Load()
  
  checkInputChange = False
  
  Me.a_save.Hyperlink.Address = "#"
  Me.btn_save.Hyperlink.Address = "#"
  
  Me.a_save.Visible = False
  Me.btn_save.Visible = False
  
  Me.input_name.SetFocus
  Me.input_name.SelStart = 0
  
  Me.RecordSource = "SELECT * FROM Products WHERE ProdCode = '" & strSQL & "'"

End Sub

Private Sub input_name_KeyPress(KeyAscii As Integer)
  
  If checkInputChange = False Then
  
    Me.a_save.Visible = True
    Me.btn_save.Visible = True
    
    checkInputChange = True
  
  End If

End Sub

Private Sub btn_save_Click()
  
  DoCmd.Save
  DoCmd.Close

End Sub

Private Sub btn_cancel_Click()
  DoCmd.Close , , acSaveNo
End Sub
