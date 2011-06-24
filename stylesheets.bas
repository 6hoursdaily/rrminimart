Option Compare Database
Option Explicit

Function main_styles()
  
  Set thisfrm = Forms("main")
  
  'login controls
  Call control_set_right("label_user_display")
  Call control_set_right("label_login")
  Call control_set_center("mainframe")
  
  'navigation
  Forms(objForm).Controls("inactive_03").Left = 1 * 1440
  Forms(objForm).Controls("inactive_05").Left = 1.7083 * 1440
  Forms(objForm).Controls("inactive_07").Left = 2.7396 * 1440
  Forms(objForm).Controls("inactive_09").Left = 3.75 * 1440
  Forms(objForm).Controls("active_03").Left = 1 * 1440
  Forms(objForm).Controls("active_05").Left = 1.7083 * 1440
  Forms(objForm).Controls("active_07").Left = 2.7396 * 1440
  Forms(objForm).Controls("active_09").Left = 3.75 * 1440
  
End Function

Function mainframe_styles()

  Set thisfrm = Forms("main")

  Call control_set_left("mainframe")

End Function