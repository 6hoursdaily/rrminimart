VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_warehouse_product_detail_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Function update_logs()
  
  If IsNull(Me.created_at) Then
    Me.created_at = Now()
  End If
  
  Me.updated_at = Now()
  Me.User_ID = strUserID

End Function

Function changeDetailView()

  If IsNull(Me.department_id) Then
    Me.text_department.Visible = False
    Me.combo_department_name.Visible = True
  Else
    Me.text_department.Visible = True
    Me.text_department.SetFocus
    Me.combo_department_name.Visible = False
  End If
  
  If IsNull(Me.subdepartment_id) Then
    Me.text_subdepartment.Visible = False
    Me.combo_subdepartment_name.Visible = True
  Else
    Me.text_subdepartment.Visible = True
    Me.text_subdepartment.SetFocus
    Me.combo_subdepartment_name.Visible = False
  End If

End Function

Private Sub Form_Load()

  Me.RecordSource = strSQLDetail
  
  strOrderBy = " ORDER by name"
  
  Me.combo_department_name.RowSource = "SELECT name, ID FROM departments" & strOrderBy
  Me.combo_department_name.BoundColumn = 2
  
  Call changeDetailView
  
End Sub

Private Sub combo_department_name_AfterUpdate()
  
  'define first the department_id value for assignment to the strSQLWhere variable
  Me.department_id.Value = Me.combo_department_name
  
  'filter the sub-department list for easier picking
  strSQLWhere = " WHERE department_id = " & Me.department_id
  Me.combo_subdepartment_name.RowSource = "SELECT name, ID FROM sub_departments" & strSQLWhere & strOrderBy
  Me.combo_subdepartment_name.BoundColumn = 2
  
  Call update_logs

End Sub

Private Sub combo_subdepartment_name_AfterUpdate()
  
  Me.subdepartment_id.Value = Me.combo_subdepartment_name
  Call update_logs
  
End Sub

Private Sub text_department_Click()
  
  Me.combo_department_name.Visible = True
  Me.combo_subdepartment_name.Visible = True
  Me.combo_department_name.SetFocus
  Me.combo_department_name.Dropdown
  Me.text_department.Visible = False
  Me.text_subdepartment.Visible = False

End Sub

Private Sub text_subdepartment_Click()

  strSQLWhere = " WHERE department_id = " & Me.department_id
  
  Me.combo_department_name.Visible = True
  Me.combo_subdepartment_name.Visible = True
  Me.combo_subdepartment_name.RowSource = "SELECT name, ID FROM sub_departments" & strSQLWhere & strOrderBy
  Me.combo_subdepartment_name.BoundColumn = 2
  
  Me.combo_subdepartment_name.SetFocus
  Me.combo_subdepartment_name.Dropdown
  Me.text_department.Visible = False
  Me.text_subdepartment.Visible = False

End Sub

Private Sub unit_per_box_AfterUpdate()
  
  Call update_logs

End Sub

Private Sub uom_AfterUpdate()
  
  Call update_logs

End Sub
Private Sub box_price_AfterUpdate()
  
  Call update_logs

End Sub
