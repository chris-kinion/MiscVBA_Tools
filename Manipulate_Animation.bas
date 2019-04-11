Attribute VB_Name = "Manipulate_Animation"
'***************************************************************************
'Module: Animation
'Procedures: timeFrame: Causes a pause between actions
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: timeFrame
'Purpose: Runs a loop process for a given duration of time
'Comments: Smallest discrete amount of time is 0.01 seconds
'***************************************************************************
Sub timeFrame(pauseTime As Double)
  Dim Start As Double
  Start = Timer
  Do
  DoEvents
  Loop Until (Timer - Start) >= pauseTime
End Sub

