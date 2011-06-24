VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  
  Me.notification.Visible = True
  Me.TimerInterval = 5000
  Me.img_sales.Hyperlink.Address = "Open Sales"

End Sub

Private Sub img_sales_Click()
  
  Call changenavitem("active_03", "inactive_03", "sales")

End Sub
