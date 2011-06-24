Option Compare Database
Option Explicit

Public ctlWidth As Integer
Public ctlHeight As Integer

Function control_set_left(controlName As String)
    
  thisfrm.Controls(controlName).Left = 720
  
End Function

Function control_set_right(controlName As String)
    
  ctlWidth = thisfrm.Controls(controlName).Width
  thisfrm.Controls(controlName).Left = frmWidth - ctlWidth - 720
  
End Function

Function control_set_center(controlName As String)
  
  ctlWidth = thisfrm.Controls(controlName).Width
  thisfrm.Controls(controlName).Left = (frmWidth / 2) - (ctlWidth / 2)


End Function