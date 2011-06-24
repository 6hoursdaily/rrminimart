Option Compare Database
Option Explicit

Function resizethisForm(thiswidth As Long, Optional thisheight As Long)
  subformwidth = thiswidth * 1440
  
  If Not IsNull(thisheight) Or thisheight > 0 Then
    subformheight = thisheight * 1440
  End If
End Function