VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  If isLoggedin = True Then
    Me.img_signin.Visible = False
  Else
    Me.img_signin.Visible = True
  End If
End Sub

Private Sub img_signin_Click()
  Forms(objForm).Controls("label_login").Visible = False
  Call route("login")
End Sub
