Attribute VB_Name = "Math_Randoms"
'***************************************************************************
'Module: Math_Randoms
'Procedures:  Function indivRandomNumber: Randomly returns 1 at rate of provided percent
            ' Function eventGenerator: Returns a random number between two numbers
'Comments:    Random numbers are generated with equal probability using Excel's
            ' built-in random number generator
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2018  Chris Kinion      Created
'***************************************************************************
Option Explicit

'***************************************************************************
'Procedure: indivRandomNumber
'Purpose: Returns a random number between two whole numbers
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2018  Chris Kinion      Created
'***************************************************************************
Function indivRandomNumber(highNum As Long, lowNum As Long) As Long
 Randomize ' Uses Excel's random number generator
 indivRandomNumber = Int((highNum - lowNum + 1) * Rnd + lowNum)
End Function

'***************************************************************************
'Procedure: eventGenerator
'Purpose:   Randomly returns 1 at rate of provided percent
'Calls:     indivRandomNumber
'Comments:  If percent in < 0 returns 0; if percent in > 100 returns 1
'           Requires UDF indivRandomNumber
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2018  Chris Kinion      Created
'***************************************************************************
Function eventGenerator(outOf100 As Long) As Long
 Dim tempNumber As Long ' Error check
 If outOf100 > 100 Then
   eventGenerator = 1
 ElseIf outOf100 <= 0 Then
   eventGenerator = 0
 End If
 
 tempNumber = indivRandomNumber(100, 0) ' UDF indivRandomNumber
 If tempNumber < outOf100 Then
   eventGenerator = 1
 Else
   eventGenerator = 0
 End If
End Function
