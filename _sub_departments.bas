VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__sub_departments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  If strSQL = "new" Or strSQL = "" Then
    Me.RecordSource = "SELECT * FROM departments"
  Else
    Me.RecordSource = strSQL
  End If
End Sub

Private Sub button_edit_Click()
On Error GoTo Err_button_edit_Click

  strSQL = "SELECT * FROM departments WHERE departments.ID = " & Me.ID
  Call edit_thisForm("_sub_departments", "_sub_department_editor")
  
Exit_button_edit_Click:
    Exit Sub

Err_button_edit_Click:
    MsgBox Err.Description
    Resume Exit_button_edit_Click
    
End Sub
