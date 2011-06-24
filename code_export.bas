Option Compare Database
Option Explicit

Public Sub backupModules()

   On Error GoTo ErrHandler
   
   Dim sPath As String
   
   sPath = "F:\rrminimart"
   MsgBox "All VBA code exported = " & exportModules(sPath)
   
   Exit Sub

ErrHandler:

   MsgBox "Error in backupModules( )." & vbCrLf & vbCrLf & _
       "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
   Err.Clear

End Sub

Public Function exportModules(sDestPath As String) As Boolean

   On Error GoTo ErrHandler
   
   Dim recSet As DAO.Recordset
   Dim frm As Form
   Dim rpt As Report
   Dim sqlStmt As String
   Dim sObjName As String
   Dim idx As Long
   Dim fOpenedRecSet As Boolean
   
   '--------------------------------------------------------------
   '  Ensure that there's a backslash at the end of the path.
   '--------------------------------------------------------------
   
   If (Mid$(sDestPath, Len(sDestPath), 1) <> "\") Then
       sDestPath = sDestPath & "\"
   End If
   
   '--------------------------------------------------------------
   '  Export standard modules and classes.
   '--------------------------------------------------------------
   
   sqlStmt = "SELECT [Name] " & _
       "FROM  MSysObjects " & _
       "WHERE ([Type] = -32761);"
   
   Set recSet = CurrentDb().OpenRecordset(sqlStmt)
   fOpenedRecSet = True
   
   If (Not (recSet.BOF And recSet.EOF)) Then
       recSet.MoveLast
       recSet.MoveFirst
   
       For idx = 1 To recSet.RecordCount
           SaveAsText acModule, recSet.Fields(0).Value, _
               sDestPath & recSet.Fields(0).Value & ".bas"
           
           If (Not (recSet.EOF)) Then
               recSet.MoveNext
           End If
       Next idx
   End If

   '--------------------------------------------------------------
   '  Export form modules.
   '--------------------------------------------------------------
   
   sqlStmt = "SELECT [Name] " & _
       "FROM  MSysObjects " & _
       "WHERE ([Type] = -32768);"
   
   Set recSet = CurrentDb().OpenRecordset(sqlStmt)
   fOpenedRecSet = True
   
   If (Not (recSet.BOF And recSet.EOF)) Then
       recSet.MoveLast
       recSet.MoveFirst
   
       For idx = 1 To recSet.RecordCount
           sObjName = recSet.Fields(0).Value
           DoCmd.OpenForm sObjName, acDesign
           Set frm = Forms(sObjName)
           
           If (frm.HasModule) Then
               DoCmd.OutputTo acOutputModule, "Form_" & _
                   sObjName, acFormatTXT, sDestPath & _
                   sObjName & ".bas"
           End If
           
           DoCmd.Close acForm, sObjName
           
           If (Not (recSet.EOF)) Then
               recSet.MoveNext
           End If
       Next idx
   End If

   '--------------------------------------------------------------
   '  Export report modules.
   '--------------------------------------------------------------

   sqlStmt = "SELECT [Name] " & _
       "FROM  MSysObjects " & _
       "WHERE ([Type] = -32764);"
   
   Set recSet = CurrentDb().OpenRecordset(sqlStmt)
   fOpenedRecSet = True
   
   If (Not (recSet.BOF And recSet.EOF)) Then
       recSet.MoveLast
       recSet.MoveFirst
   
       For idx = 1 To recSet.RecordCount
           sObjName = recSet.Fields(0).Value
           DoCmd.OpenReport sObjName, acDesign
           Set rpt = Reports(sObjName)
           
           If (rpt.HasModule) Then
               DoCmd.OutputTo acOutputModule, "Report_" & _
                   sObjName, acFormatTXT, sDestPath & _
                   sObjName & ".bas"
           End If
           
           DoCmd.Close acReport, sObjName
           
           If (Not (recSet.EOF)) Then
               recSet.MoveNext
           End If
       Next idx
   End If

   exportModules = True           ' Success.
   
CleanUp:

   If (fOpenedRecSet) Then
       recSet.Close
       fOpenedRecSet = False
   End If
   
   Set frm = Nothing
   Set rpt = Nothing
   Set recSet = Nothing
   
   Exit Function

ErrHandler:

   MsgBox "Error in exportModules( )." & vbCrLf & vbCrLf & _
       "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
   Err.Clear
   exportModules = False              ' Failed.
   GoTo CleanUp

End Function       '  exportModules( )