Option Compare Database
Option Explicit

Function cmdTestLink_Click()
On Error GoTo Err_cmdTestLink_Click
Const conLINKED_TABLE As String = "tblDepartment"
 
'A Linked Table will have a Connect Strain whose Length is > 0
If Len(CurrentDb.TableDefs(conLINKED_TABLE).Connect) > 0 Then
  'OK, we know that conLINKED_TABLE is a Linked Table, but is the Link valid?
  'The next line of code will generate Errors 3011 or 3024 if it isn't
  CurrentDb.TableDefs(conLINKED_TABLE).RefreshLink
  'If you get to this point, you have a valid, Linked Table
  '...normal code processing here
Else
  'An Internal Table will have a Connect String Length of 0
  MsgBox "[" & conLINKED_TABLE & "] is a Non-Linked Table", vbInformation, "Internal Table"
End If
 
Exit_cmdTestLink_Click:
  Exit Function
 
Err_cmdTestLink_Click:
  Select Case Err.Number
    Case 3265
      MsgBox "[" & conLINKED_TABLE & "] does not exist as either an Internal or Linked Table", _
             vbCritical, "Table Missing"
    Case 3011, 3024     'Linked Table does not exist or DB Path not valid
      MsgBox "[" & conLINKED_TABLE & "] is not a valid, Linked Table", vbCritical, "Link Not Valid"
    Case Else
      MsgBox Err.Description & Err.Number, vbExclamation, "Error in cmdTestLink_Click()"
  End Select
    Resume Exit_cmdTestLink_Click

End Function