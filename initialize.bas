Option Compare Database
Option Explicit

Public myApp As Access.Application
Public DevMode As Boolean
Public frmWidth As Integer
Public frmHeight As Integer
Public thisfrm As Access.Form

Function init()

  Set thisfrm = Forms("main")

  frmWidth = thisfrm.InsideWidth
  frmHeight = thisfrm.InsideHeight

End Function