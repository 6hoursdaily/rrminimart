VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  
  Me.btn_start.Hyperlink.Address = "#"
  Me.input_password = "Password"
  Me.input_password.InputMask = ""
  
  If isLoggedin = False Then
    Forms(objForm).Controls("label_login").Caption = "Login"
  Else
    Forms(objForm).Controls("label_login").Caption = "Logout"
  End If
    
End Sub

Private Sub btn_start_Click()
On Error GoTo Err_button_login_Click
    
  strSQL = "SELECT * from Users WHERE User = '" & Me.input_user & "'"
  
  If DevMode = True Then
    'development db:
    Set db = OpenDatabase("C:\development\RR Minimart Server.mdb")
  
  Else
    'production db:
    Set db = OpenDatabase("\\POS01-HP\RR Mini Mart\Management System\RR Minimart Server.mdb") 'CurrentDb() 'You may also use: OpenDatabase("MyDatabase.mdb")
  
  End If
  
  Set rst = db.OpenRecordset(strSQL, dbOpenDynaset)  'db.OpenRecordSet("MyTable", dbOpenDynaset)
  Me.RecordSource = strSQL
  
  If IsNull(Me.username) Or Me.username = "" Then
      
    MsgBox "Username not found.", vbOKOnly, "Oops, something went wrong."
  
  ElseIf (Me.input_user = Me.username) And (Me.input_password = Me.password) Then
    
    rst.MoveFirst
    rst.Edit
    rst!last_login = Now()
    rst.Update
    
    rst.Close
    
    strUser = Form_login.input_user
    strUserID = Me.ID
    strRoleID = Me.role_id
        
    isLoggedin = True

    Call route("welcome")
    Call shownav
    Call mainframe_styles
    
    If strUser <> "" Then
      Forms(objForm).Controls("label_user_display").Caption = "Welcome " & strUser
    End If
    
    If isLoggedin = True Then
    
      Forms(objForm).Controls("label_login").Caption = "Logout"
      Forms(objForm).Controls("label_login").Visible = True
      Forms(objForm).Controls("label_login").ControlTipText = "Logout"
    
    End If
  
  Else
  
    MsgBox "Your username and password is incorrect.", vbOKOnly, "Oops, something went wrong."
  
  End If

Exit_button_login_Click:
    Exit Sub

Err_button_login_Click:
    MsgBox Err.Description
    Resume Exit_button_login_Click

End Sub

Private Sub input_password_GotFocus()
  Me.input_password = ""
  Me.input_password.InputMask = "PASSWORD"
End Sub
