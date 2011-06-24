Option Compare Database
Option Explicit

Public activateTimer As Boolean
Public timerForm As String 'declare target form to use timer on form's onLoad event
'Public timerControl As String

Public Declare Function SetTimer Lib "user32" ( _
    ByVal HWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "user32" ( _
    ByVal HWnd As Long, _
    ByVal nIDEvent As Long) As Long

Public TimerID As Long
Public TimerSeconds As Single

Function StartTimer(Optional repeatFlag As Integer, Optional timerDelay As Integer)
  
  TimerSeconds = 60 'set how many seconds before the timer refreshes
  TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc)

End Function

Function EndTimer()
    On Error Resume Next
    KillTimer 0&, TimerID
End Function

Function TimerProc(ByVal HWnd As Long, ByVal uMsg As Long, _
        ByVal nIDEvent As Long, ByVal dwTimer As Long)
On Error GoTo Err_TimerProc
    
  'set timer actions here
  'Forms(objForm).Controls(objControl).Form.Controls(timerForm).Form.RecordSource = strSQLSearch
  objSubformControl = "viewer"
  Call format_viewer("sales_viewer", 1, 1)
    
  'MsgBox "it works"
    
Exit_TimerProc:
    Exit Function

Err_TimerProc:
    MsgBox Err.Description
    Resume Exit_TimerProc
End Function