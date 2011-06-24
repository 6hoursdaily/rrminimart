Option Compare Database
Option Explicit

Function main_styles()
  
  Set thisfrm = Forms("main")
  
  Call control_set_right("label_user_display")
  Call control_set_right("label_login")
  Call control_set_center("mainframe")
  
End Function

Function mainframe_styles()

  Set thisfrm = Forms("main")

  Call control_set_left("mainframe")

End Function