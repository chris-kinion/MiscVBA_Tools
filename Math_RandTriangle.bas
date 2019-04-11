Attribute VB_Name = "Math_RandTriangle"
'***************************************************************************
'Module: Math_RandTriangle
'Procedures:  RandTri:      Return a random long data type along a triangular distribution given
                          ' a minimum value. Assumes mode is 10% more than the input, max is 20%
                          ' more than the input.
'             RandTriDist:  Return a random long data type along a random triangular distribution given a min, max and mode
'Comments: All procedures are functions
'Changes----------------------------------------------
' Date        Programmer        Change
' 03/19/2019  Chris Kinion      Created
'***************************************************************************
Option Base 1
Option Explicit

'***************************************************************************
'Procedure: RandTri
'Purpose: Return a random long data type along a triangular distribution given a minimum value
'Comments:
' Procedure RandTri assumes the input number is the lowest value of a long data type, _
the mode is 10% more than the input, and the max is 20% more than the input. _
The output is a triangular distribution using VBA's uniform random distribution function Rnd() _
and coerced to the long data type. The shape of the distribution may be changed by varying _
the coefficients of dblMaxB (which calculates the maximum) and dblModeC (which calculates the mode). _
See also: https://en.wikipedia.org/wiki/Triangular_distribution
'Changes----------------------------------------------
' Date        Programmer        Change
' 3/19/2019   Chris Kinion      Written
'***************************************************************************
Function RandTri(lngNumIn As Long) As Long
  'Application.Volatile
  Dim dblMinA As Double
  Dim dblMaxB As Double
  Dim dblModeC As Double
  Dim U As Double
  Dim F_c As Double
  
  dblMinA = lngNumIn
  dblMaxB = 1.2 * dblMinA
  dblModeC = 1.1 * dblMinA
  U = Rnd()
  F_c = (dblModeC - dblMinA) / (dblMaxB - dblMinA) 'F(c) = (c-a)/(b-a)
  
  If U > 0 And U < F_c Then ' X = a + sqrt(U(b-a)(c-a))           for 0 < U < F(c)
    RandTri = CLng(dblMinA + Sqr(U * (dblMaxB - dblMinA) * (dblModeC - dblMinA)))
  ElseIf F_c <= U And U < 1 Then ' X = b - sqrt((1-U)(b-a)(b-c))       for F(c) <= U < 1
    RandTri = CLng(dblMaxB - Sqr((1 - U) * (dblMaxB - dblMinA) * (dblMaxB - dblModeC)))
  End If
End Function

'***************************************************************************
'Procedure: RandTriDist
'Purpose: Return a random long data type along a random triangular distribution given a min, max and mode
'Comments:
'The function RandTriDist returns a long data type following a triangular distribution provided _
a minimum, maximum and a mode which is inclusively in-between. Should any of these variables be _
misaligned, the function will return 0.
'Changes----------------------------------------------
' Date        Programmer        Change
' 03/19/2019  Chris Kinion      Written
'***************************************************************************
Function RandTriDist(dblMinA As Double, dblMaxB As Double, dblModeC As Double) As Long
  'Application.Volatile
  Dim U As Double
  Dim F_c As Double
  
  If dblModeC <= dblMaxB And dblMinA <= dblModeC Then
    U = Rnd()
    F_c = (dblModeC - dblMinA) / (dblMaxB - dblMinA) 'F(c) = (c-a)/(b-a)
    If U > 0 And U < F_c Then ' X = a + sqrt(U(b-a)(c-a))           for 0 < U < F(c)
      RandTriDist = CLng(dblMinA + Sqr(U * (dblMaxB - dblMinA) * (dblModeC - dblMinA)))
    ElseIf F_c <= U And U < 1 Then ' X = b - sqrt((1-U)(b-a)(b-c))       for F(c) <= U < 1
      RandTriDist = CLng(dblMaxB - Sqr((1 - U) * (dblMaxB - dblMinA) * (dblMaxB - dblModeC)))
    End If
  Else
    RandTriDist = 0
    'Debug.Print "Function RandTriDist was provided a min of " & dblMinA & ", a max of " & dblMaxB & ", and a mode of " & dblModeC & " which defies logic."
  End If
End Function
