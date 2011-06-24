VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_starting_inventory_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
  strSQL = "new"
End Sub

Private Sub Form_Load()
    strOrderBy = " ORDER BY ProdNameLong"
    Me.input_name.SetFocus
    Me.RecordSource = "SELECT * FROM query_starting_inventory" & strOrderBy
End Sub

Private Sub pieces_Exit(Cancel As Integer)
    'Form_starting_inventory_view.total.Value = Form_starting_inventory_view.computed_total.Value
    Me.input_name.SetFocus
End Sub
Private Sub button_search_Click()
On Error GoTo Err_button_search_Click

  Dim strName As String

  If Not IsNull(Me.input_code) Then
    strCode = Me.input_code
  End If
  
  If Not IsNull(Me.input_name) Then
    strName = Me.input_name
  End If
  
  If strCode = "" And strName = "" Then
    Me.RecordSource = "SELECT * FROM query_starting_inventory"
    Exit Sub
  ElseIf (Not IsNull(strCode) = True) And (Not IsNull(strName) = True) Then
    strSQLSearch = "SELECT * FROM query_starting_inventory WHERE ProdCode LIKE '" & strCode & "*' AND ProdNameLong LIKE '*" & strName & "*'"
  ElseIf Not IsNull(strCode) Then
    strSQLSearch = "SELECT * FROM query_starting_inventory WHERE ProdCode LIKE '" & strCode & "*'"
  ElseIf Not IsNull(strName) Then
    strSQLSearch = "SELECT * FROM query_starting_inventory WHERE ProdNameLong LIKE '*" & strName & "*'"
  Else
    MsgBox "Please provide details to search"
    Exit Sub
  End If
  
  strSQLSearch = strSQLSearch & strOrderBy
  Me.RecordSource = strSQLSearch
  Me.input_name.SetFocus
    

Exit_button_search_Click:
    Exit Sub

Err_button_search_Click:
    MsgBox Err.Description
    Resume Exit_button_search_Click
    
End Sub
Private Sub button_reset_Click()
On Error GoTo Err_button_reset_Click

    Me.RecordSource = "SELECT * FROM query_starting_inventory"

Exit_button_reset_Click:
    Exit Sub

Err_button_reset_Click:
    MsgBox Err.Description
    Resume Exit_button_reset_Click
    
End Sub
Private Sub button_return_Click()
On Error GoTo Err_button_return_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "_main_menu"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "starting_inventory_view"
Exit_button_return_Click:
    Exit Sub

Err_button_return_Click:
    MsgBox Err.Description
    Resume Exit_button_return_Click
    
End Sub
