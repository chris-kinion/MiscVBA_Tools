Attribute VB_Name = "Event_Application"
'***************************************************************************
'Module: Event_Application
'Procedures:  ReduceFunctionality: Limit Excel activity to boost VBA speed
            ' RestoreFunctionality: Return Excel to normal operation
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2018  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: ReduceFunctionality
'Purpose: Limit Excel activity to boost VBA speed
'Changes----------------------------------------------
' Date        Programmer        Change
' 09/12/2018  Chris Kinion      Created
'***************************************************************************
Sub ReduceFunctionality()
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  'Application.DisplayStatusBar = False
End Sub

'***************************************************************************
'Procedure: RestoreFunctionality
'Purpose: Return Excel to normal upon exit / end procedure
'Changes----------------------------------------------
' Date        Programmer        Change
' 09/12/2018  Chris Kinion      Created
'***************************************************************************
Sub RestoreFunctionality()
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  Application.DisplayStatusBar = True
End Sub

