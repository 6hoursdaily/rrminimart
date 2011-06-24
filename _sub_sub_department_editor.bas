VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__sub_sub_department_editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  
  Call load_thisForm("_sub_sub_department_editor", "sub_departments")
  
  If strSQL = "new" Then
    Me.RecordSource = "SELECT * FROM sub_departments"
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
  Else
    Me.RecordSource = strSQLEdit
  End If

End Sub

Private Sub department_name_view_Click()

  Me.department_name.SetFocus
  Me.department_name.Dropdown
  Me.department_name_view.Visible = False
  
End Sub

Private Sub department_name_AfterUpdate()

  Me.department_id = Me.department_name.Value

End Sub

Private Sub button_save_Click()
On Error GoTo Err_button_save_Click
  
  If IsNull(Me.department_id) Or Me.department_id = "" Then
    MsgBox "Please choose a department to associate this sub-department", vbExclamation, "Choose Department"
    Exit Sub
  End If
  
  If IsNull(Me.sub_department_name) Or Me.sub_department_name = "" Then
    Me.sub_department_name = "Sub-Department Name"
  End If
  
  If IsNull(Me.sub_department_desc) Or Me.sub_department_desc = "" Then
    Me.sub_department_desc = "Description"
  End If
  
  If IsNull(Me.created_at) Or Me.created_at = "" Then
    Me.created_at = Now()
  End If
  
  Me.updated_at = Now()
  Me.User_ID = strUserID
  DoCmd.Save acDefault
  
  strSQL = "SELECT sub_departments.User_ID AS User_ID, sub_departments.name, sub_departments.description, sub_departments.ID, sub_departments.department_id AS department_id, sub_departments.created_at AS created_at, sub_departments.updated_at AS updated_at, departments.name AS department_name FROM sub_departments LEFT JOIN departments ON sub_departments.department_id = departments.ID"
  
  Call format_viewer("_sub_sub_departments", 0)
  
Exit_button_save_Click:
    Exit Sub

Err_button_save_Click:
    MsgBox Err.Description
    Resume Exit_button_save_Click
    
End Sub
Private Sub xbutton_save_Click()
On Error GoTo Err_button_save_Click
  
  Dim str_name As String
  Dim str_description  As String
  Dim str_department_id As Integer
  Dim str_created_at As Integer
  
  str_department_id = Me.department_id
  
  If Me.sub_department_name = Null Or Me.sub_department_name = "" Then
    str_name = "Department Name"
  Else
    str_name = Me.sub_department_name
  End If
  
  If IsNull(Me.sub_department_desc) Or Me.sub_department_desc = "" Then
    str_description = "Description"
  Else
    str_description = Me.sub_department_desc
  End If
  If IsNull(Me.created_at) Then
    str_created_at = 0
  Else
    str_created_at = 1
  End If
  
  Forms("department_manager").Controls("button_create").SetFocus
  Me.Visible = False
  Me.RecordSource = ""
  
  Set db = CurrentDb()
  
  If strSQL = "new" Then
    
    Set rst = db.OpenRecordset("sub_departments", dbOpenDynaset)
    rst.AddNew
  
  Else
    
    Set rst = db.OpenRecordset(strSQLEdit, dbOpenDynaset)
    rst.MoveFirst
    rst.Edit
  
  End If
  
  rst!name = str_name
  rst!Description = str_description
  If str_created_at = 0 Then
    rst!created_at = Now()
  End If
  rst!updated_at = Now()
  rst!department_id = str_department_id
  rst!User_ID = strUserID
  
  rst.Update
  rst.Close
  
  strSQL = "SELECT sub_departments.User_ID AS User_ID, sub_departments.name, sub_departments.description, sub_departments.ID, sub_departments.department_id AS department_id, sub_departments.created_at AS created_at, sub_departments.updated_at AS updated_at, departments.name AS department_name FROM sub_departments LEFT JOIN departments ON sub_departments.department_id = departments.ID"
  Call format_viewer("_sub_sub_departments", 1)
  
Exit_button_save_Click:
    Exit Sub

Err_button_save_Click:
    MsgBox Err.Description
    Resume Exit_button_save_Click
    
End Sub
