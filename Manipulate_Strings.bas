Attribute VB_Name = "Manipulate_Strings"
'***************************************************************************
'Module: Manipulate_Strings
'Procedures:  Function removeExtension: Removes any extensions from a file name
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/08/2019  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: removeExtension
'Purpose: Removes any extensions from a file name
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/08/2019  Chris Kinion      Created
'***************************************************************************
Function removeExtension(fileName As String) As String
  If fileName = "" Then
    removeExtension = ""
    Exit Function
  ElseIf InStrRev(fileName, ".") = 0 Then
    removeExtension = fileName
    Exit Function
  Else
    removeExtension = Left(fileName, InStrRev(fileName, ".") - 1)
  End If
End Function
