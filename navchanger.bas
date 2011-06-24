Option Compare Database
Option Explicit

Function changenavitem(navshow As String, navhide As String, targetForm As String)

  If objForm = "" Then
    objForm = "main"
  End If
  
  Call checkUserRole("nav_change")
  
  Forms(objForm).Controls(navshow).Visible = True
  Forms(objForm).Controls(navhide).Visible = False
  
  Forms(objForm).Detail.BackColor = 16777215
  
  Call route(targetForm)
  
End Function

Function shownav()
  
  If objForm = "" Then
    objForm = "main"
  End If
  
  Forms(objForm).Controls("inactive_03").Hyperlink.Address = "#"
  Forms(objForm).Controls("inactive_05").Hyperlink.Address = "#"
  Forms(objForm).Controls("inactive_07").Hyperlink.Address = "#"
  Forms(objForm).Controls("inactive_09").Hyperlink.Address = "#"

  Forms(objForm).Controls("inactive_03").Visible = True
  
  Call checkUserRole("main_nav")
  
End Function

Function hidenav()

  Forms(objForm).Controls("inactive_03").Visible = False
  Forms(objForm).Controls("inactive_05").Visible = False
  Forms(objForm).Controls("inactive_07").Visible = False
  Forms(objForm).Controls("inactive_09").Visible = False
  Forms(objForm).Controls("active_03").Visible = False
  Forms(objForm).Controls("active_05").Visible = False
  Forms(objForm).Controls("active_07").Visible = False
  Forms(objForm).Controls("active_09").Visible = False

End Function