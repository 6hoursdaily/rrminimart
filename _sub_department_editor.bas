VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__sub_department_editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Load()
  
  Call load_thisForm("_sub_department_editor", "departments")

End Sub

Private Sub department_desc_GotFocus()
  If Me.Controls("department_desc") = "" Then
    Me.Controls("department_desc") = Me.Controls("department_name")
  End If
End Sub

Private Sub button_save_Click()
On Error GoTo Err_button_save_Click

  If IsNull(Me.department_name) Or Me.department_name = "" Then
    Me.department_name = "Department Name"
  End If
  If IsNull(Me.department_desc) Or Me.department_desc = "" Then
    Me.department_desc = "Description"
  End If
  If IsNull(Me.created_at) Or Me.created_at = "" Then
    Me.created_at = Now()
  End If
  
  Me.updated_at = Now()
  Me.User_ID = strUserID
  
  DoCmd.Save acDefault
  
  strSQL = "SELECT * FROM departments"
  Call format_viewer("_sub_departments", 0)

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
  Dim str_created_at As Integer
  
  If Me.department_desc = Null Then
    str_name = "Department Name"
  Else
    str_name = Me.department_name
  End If
  
  If IsNull(Me.department_desc) Or Me.department_desc = "" Then
    str_description = "Description"
  Else
    str_description = Me.department_desc
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
    
    Set rst = db.OpenRecordset("departments", dbOpenDynaset)
    rst.AddNew
  
  Else
    
    Set rst = db.OpenRecordset(strSQL, dbOpenDynaset)
    rst.MoveFirst
    rst.Edit
  
  End If
  
  rst!name = str_name
  rst!Description = str_description
  If str_created_at = 0 Then
    rst!created_at = Now()
  End If
  rst!updated_at = Now()
  rst!User_ID = strUserID
  
  rst.Update
  rst.Close
  
  strSQL = "SELECT * FROM departments"
  Call format_viewer("_sub_departments", 0)

Exit_button_save_Click:
    Exit Sub

Err_button_save_Click:
    MsgBox Err.Description
    Resume Exit_button_save_Click
    
End Sub
