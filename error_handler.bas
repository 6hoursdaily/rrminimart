Option Compare Database
Option Explicit

Function trackError()
  
  MsgBox Err.Description
  Resume Exit_trackError

Exit_trackError:
  Exit Function

End Function