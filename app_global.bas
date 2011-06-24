Option Compare Database
Option Explicit

'---- Global Variables ----

'Application init configuration
Public strAppInit As String

'Form and Criteria targets
Public stDocName As String
Public stLinkCriteria As String
Public objForm As String 'Form name for viewer
Public objControl As String 'Form control name
Public objSubformControl As String
Public strViewer As String
Public frmCurrentForm As Form
Public resetSubform As Integer 'boolean check if subform should change dimensions

'DAO recordset handlers
Public db As DAO.Database
Public rst As DAO.Recordset

'recordset filters
Public strSQL As String
Public strUser As String
Public strUserID As Variant
Public strSQLSearch As String
Public strCode As String
Public strSQLEdit As String
Public strSQLEditName As String
Public strSQLDetail As String
Public strSQLWhere As String
Public strOrderBy As String
Public strGroupBy As String
Public strInventoryCode As Integer
Public strSQLInventory As String
Public strInventoryCount As Integer
Public strPIID As Integer

'chart recordset filters
Public strSQLChart As String
Public strSQLRank As String

'error handlers
Public errMessage As String

'form size handlerss
Public subformwidth As Integer
Public subformheight As Integer

Function initActiveForm()
  
  objForm = Screen.ActiveForm.name
  
  If objControl = "" Then 'reloading the form is the variable controls does not exist so that its reintialized
    DoCmd.Close acForm, objForm
    DoCmd.OpenForm objForm
  End If

End Function

Function format_viewer(form_to_load As Variant, resetSubform As Integer, Optional within_subform As Integer)
      
  objForm = Screen.ActiveForm.name
  
  If objControl = "" Then 'reloading the form is the variable controls does not exist so that its reintialized
    DoCmd.Close acForm, objForm
    DoCmd.OpenForm objForm
  End If
  'To use:
      'step 1: within your current form
      'define objControl in your Form_Load event
      'define subformwidth and subformheight if necessary
      
      'step 2: calling this function
      'define the form_to_load into subform
      'define boolean resetSubform if subform dimensions should be reset
      'optionally define within_subform if subform is within another subform of the active form
  
  'check if subform should be resized
  If resetSubform = 0 Then 'no resize, just plain load the subform content
    'check if subform is under another subform
    If within_subform = 1 Then
        Forms(objForm).Controls(objControl).Form.Controls(objSubformControl).SourceObject = form_to_load
        Forms(objForm).Controls(objControl).Form.Controls(objSubformControl).Visible = True
    Else
        Forms(objForm).Controls(objControl).SourceObject = form_to_load
        Forms(objForm).Controls(objControl).Visible = True
    End If
  Else
    If within_subform = 1 Then
      Forms(objForm).Controls(objControl).Form.Controls(objSubformControl).SourceObject = form_to_load
      Forms(objForm).Controls(objControl).Form.Controls(objSubformControl).Visible = True
      Forms(objForm).Controls(objControl).Form.Controls(objSubformControl).Width = subformwidth
      Forms(objForm).Controls(objControl).Form.Controls(objSubformControl).Height = subformheight
    
    Else
      Forms(objForm).Controls(objControl).SourceObject = form_to_load
      Forms(objForm).Controls(objControl).Visible = True
      Forms(objForm).Controls(objControl).Width = subformwidth
      Forms(objForm).Controls(objControl).Height = subformheight

    End If
    
  End If
    
End Function

Function toggle_viewer(set_visibility As String)
  
  objForm = Screen.ActiveForm.name
  
  If objControl = "" Then 'reloading the form is the variable controls does not exist so that its reintialized
    DoCmd.Close acForm, objForm
    DoCmd.OpenForm objForm
  End If
  
  'handles controls within the active form
  Forms(objForm).Controls(objControl).Visible = set_visibility
    
End Function

Function toggle_subviewer(set_visibility As Boolean)

    'Handles controls within a subform in the active form
    Forms(objForm).Controls(objControl).Form.Controls(objSubformControl).Visible = set_visibility
    
End Function

Function return_to(return_form As String, Optional form_to_close As String)
  
  'simplifying form routing
  DoCmd.OpenForm return_form, , , stLinkCriteria 'return_form as the form you're going back to
  DoCmd.Close acForm, form_to_close 'form_to_close as the form you need to close, usually the form you're on

End Function
Function open_Form(form_to_open As String, Optional form_to_close As String)
  
  'simplifying form routing
  If form_to_close <> "" Then
    DoCmd.Close acForm, form_to_close
  End If
  If form_to_open = "_main_menu" Then
      objControl = "submenu"
  End If
  DoCmd.OpenForm form_to_open, , , stLinkCriteria 'return_form as the form you're going back to

End Function

Function load_thisForm(thisForm As String, form_resource As String)
 
  If strSQL = "new" Then
    'Forms(thisForm).RecordSource = form_resource
    Forms(objForm).Controls(objControl).Form.RecordSource = form_resource
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
  Else
    Forms(objForm).Controls(objControl).Form.RecordSource = strSQL
  End If

End Function

Function edit_thisForm(closethisForm As String, targetEditor)
  
  DoCmd.Close acForm, closethisForm
  Call format_viewer(targetEditor, 0)

End Function