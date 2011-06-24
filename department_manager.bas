VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_department_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim strEditorForm As String

Private Sub Form_Load()
On Error GoTo Err_Form_Load

  Me.SetFocus
  objControl = "viewer"

  Call option_department_GotFocus

Exit_Form_Load:
    Exit Sub

Err_Form_Load:
    MsgBox Err.Description
    Resume Exit_Form_Load

End Sub

Private Sub option_department_GotFocus()

  strEditorForm = "departments"
  strSQL = "SELECT * FROM departments"
  Call format_viewer("_sub_departments", 0)
  Me.button_create.Caption = "Create Department"

End Sub

Private Sub option_subdepartment_GotFocus()

  strEditorForm = "sub_departments"
  strSQL = "SELECT sub_departments.User_ID AS User_ID, sub_departments.name, sub_departments.description, sub_departments.ID, sub_departments.department_id AS department_id, sub_departments.created_at AS created_at, sub_departments.updated_at AS updated_at, departments.name AS department_name FROM sub_departments LEFT JOIN departments ON sub_departments.department_id = departments.ID"
  Call format_viewer("_sub_sub_departments", 0)
  Me.button_create.Caption = "Create Sub-Department"

End Sub
Private Sub button_create_Click()

  strSQL = "new"
  
  Me.viewer.SourceObject = ""

  If strEditorForm = "departments" Then
    Call format_viewer("_sub_department_editor", 0)
    Forms(objForm).Controls(objControl).Form.Controls("department_name").SetFocus
  Else
    Call format_viewer("_sub_sub_department_editor", 0)
    Forms(objForm).Controls(objControl).Form.Controls("department_name_view").Visible = False
  End If
  

End Sub
Private Sub button_back_Click()
On Error GoTo Err_button_back_Click

  Call open_Form("_main_menu", "department_manager")

Exit_button_back_Click:
    Exit Sub

Err_button_back_Click:
    MsgBox Err.Description
    Resume Exit_button_back_Click
    
End Sub
