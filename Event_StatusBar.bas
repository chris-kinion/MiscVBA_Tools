Attribute VB_Name = "Event_StatusBar"
'***************************************************************************
'Module: Event_StatusBar
'Procedures:  UpdateStatusBar
            ' statusBarDisplay
'Changes----------------------------------------------
' Date        Programmer        Change
' 09/12/2018  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'Procedure: UpdateStatusBar
'Purpose: Updates the status bar with a message
'Comments: Useful to ensure doesn't lock up Excel, but can also be written inline
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/12  Chris Kinion      Created
'***************************************************************************
Sub UpdateStatusBar(statusMessage As String)
  Application.StatusBar = statusMessage
  DoEvents
End Sub

'***************************************************************************
'Procedure: statusBarDisplay
'Purpose: Animates status bar as counter increases
'Called by:
'Calls:
'Arguments: lngCounter As Long, strMessage As String
'Comments:  Outputs input string to be changed next iteration
'           Caution: Removes anything to right of "."
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/10/04  Chris Kinion      Created
'***************************************************************************
Function statusBarDisplay(lngCounter As Long, strMessage As String) As String
  If lngCounter Mod 1500 = 0 Then
    strMessage = Split(strMessage, ".")(0)
    Application.StatusBar = strMessage
  ElseIf lngCounter Mod 300 = 0 Then
    strMessage = strMessage & "."
    Application.StatusBar = strMessage
    DoEvents
  End If
  statusBarDisplay = strMessage
End Function

