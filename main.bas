VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()

  'set mode
  DevMode = True
  
  isLoggedin = False
  
  'initialize app
  If DevMode = True Then
    'DoCmd.Maximize
    Call init
    Call main_styles
      
  Else
    DoCmd.Maximize
    Call init
    Call main_styles
  
  End If
  
  Me.SetFocus
  Me.Detail.BackColor = 3618615
  
  If strAppInit = "" Then
    strAppInit = "init"
    Me.label_user_display.Caption = "Welcome Guest"
  Else
    Me.label_user_display.Caption = "Welcome " & strUser
  End If
  
  If isLoggedin = False Then
      Me.label_login.Caption = "Login"
  End If
  
  Me.label_login.Hyperlink.Address = "#"
  Me.img_logo.Hyperlink.Address = "#"
  
  'setup content holder variables
  objControl = "mainframe"
  Call resizethisForm(8.5, 4.5)
  
  'boot app content
  Call format_viewer("home", 1)
  Call hidenav

End Sub

Private Sub Form_Resize()
  'Call main_styles
End Sub

Private Sub img_logo_Click()
  
  Me.Detail.BackColor = 3618615
  Call route("home")
  Call hidenav

End Sub

Private Sub label_login_Click()
  
  Me.Detail.BackColor = 3618615
  
  strUser = ""
  
  Call resizethisForm(8.5, 4.5)
  Call format_viewer("login", 1, 0)
  
  Me.label_user_display.Caption = "Welcome Guest"
  
  If isLoggedin = True Then
    isLoggedin = False
    Me.label_login.Visible = False
  End If
  
  Call hidenav

End Sub

Private Sub inactive_03_Click()

  Call changenavitem("active_03", "inactive_03", "sales")

End Sub

Private Sub inactive_05_Click()
  
  'Call changenavitem("active_05", "inactive_05", "_sub_purchasing")
  MsgBox "Development on-going.", vbOKOnly, "We're making this better"

End Sub

Private Sub inactive_07_Click()
  
  Call changenavitem("active_07", "inactive_07", "warehouse")
  'MsgBox "Development on-going.", vbOKOnly, "We're making this better"

End Sub

Private Sub inactive_09_Click()
  
  'Call changenavitem("active_09", "inactive_09", "_sub_audit")
  MsgBox "Development on-going.", vbOKOnly, "We're making this better"

End Sub


