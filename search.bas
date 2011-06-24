Option Compare Database
Option Explicit

Public strSQLBase As String 'query base
Public objSearchForm As String 'require form name
Public objInputCode As String 'text field for product code entry
Public objInputName As String 'text field for product name entry
Public searchCode As String
Public searchName As String
Public colS1 As String 'column to search
Public colS2 As String 'column to search

Function search_records()

    'check form controls if they have user input
    If Not IsNull(Forms(objSearchForm).Controls(objInputCode)) Then
        searchCode = Forms(objSearchForm).Controls(objInputCode)
    Else
        searchCode = ""
    End If
    
    If Not IsNull(Forms(objSearchForm).Controls(objInputName)) Then
        searchName = Forms(objSearchForm).Controls(objInputName)
    Else
        searchName = ""
    End If
    
    'main search logic
    If (searchCode = "" And searchName = "") Or (IsNull(searchCode) And IsNull(searchName)) Then
        strSQLSearch = strSQLBase
    ElseIf (Not IsNull(searchCode) = True) And (Not IsNull(searchName) = True) Then
        strSQLSearch = strSQLBase & " WHERE " & colS1 & " LIKE '" & searchCode & "*' AND " & colS2 & " LIKE '*" & searchName & "*'"
    ElseIf Not IsNull(searchCode) Then
        strSQLSearch = strSQLBase & " WHERE " & colS1 & " LIKE '" & searchCode & "*'"
    ElseIf Not IsNull(searchName) Then
        strSQLSearch = strSQLBase & " WHERE " & colS2 & " LIKE '*" & searchName & "*'"
    Else
        MsgBox "Please provide details to search"
        Exit Function
    End If

    Forms(objSearchForm).RecordSource = strSQLSearch

End Function