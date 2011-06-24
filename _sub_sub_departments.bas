VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__sub_sub_departments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  
  If strSQL = "new" Then
    Me.RecordSource = "SELECT * FROM sub_departments"
  Else
    Me.RecordSource = strSQL
  End If

End Sub
Private Sub button_edit_Click()
On Error GoTo Err_button_edit_Click

  strSQLEdit = strSQL & " WHERE sub_departments.ID = " & Me.ID
  Call edit_thisForm("_sub_sub_departments", "_sub_sub_department_editor")
  
  'just hiding the dropdown so its not automatically editable
  Forms(objForm).Controls(objControl).Form.Controls("department_name_view").Visible = True

Exit_button_edit_Click:
    Exit Sub

Err_button_edit_Click:
    MsgBox Err.Description
    Resume Exit_button_edit_Click
    
End Sub

Private Sub department_name_Click()
  
  strSQLWhere = " WHERE sub_departments.department_id = " & Me.department_id
  'strSQL = "SELECT sub_departments.User_ID AS User_ID, sub_departments.name, sub_departments.description, sub_departments.ID, sub_departments.department_id AS department_id, sub_departments.created_at AS created_at, sub_departments.updated_at AS updated_at, departments.name AS department_name FROM sub_departments LEFT JOIN departments ON sub_departments.department_id = departments.ID " & strSQLWhere
  strSQL = strSQL + strSQLWhere
  Me.RecordSource = strSQL
  'Call Form_Load
End Sub
