Attribute VB_Name = "Math_Geometry"
'***************************************************************************
'Module: Math_Geometry
'Procedures: DistanceBetween2Points: Euclidean (2D) distance formula
'            PointInPolygon: input (x,y) and polygon, determine if (x,y) is inside polygon
'            ClosePoly: input array of (x,y) and ensures last entry is same as first
'            GetTslope: Returns an inverse (perpendicular) slope
'Comments:  Arrays assume use of Option Base 1
'Changes----------------------------------------------
' Date        Programmer        Change
' 09/01/2018 Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: DistanceBetween2Points
'Purpose: 2D distance formula
'Comments: d=sqrt((y2-y1)^2 + (x2-x1)^2)
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2018  Chris Kinion      Created
'***************************************************************************
Function DistanceBetween2Points(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
 Dim a As Double
 Dim b As Double
 a = x2 - x1
 b = y2 - y1
 DistanceBetween2Points = Sqr((a * a) + (b * b))
End Function

'***************************************************************************
'Procedure: PointInPolygon
'Purpose: input (x,y) and polygon, determine if (x,y) is inside polygon
'Comments:  Variable polyGon is array in format varPolygon(x,y)
'           Requires UDF closePoly(), ' Return 0 FALSE or 1 TRUE
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2018  Chris Kinion      Created
'***************************************************************************
Function PointInPolygon(xHit As Double, yHit As Double, polyGon As Variant) As Integer
 Dim slope As Double ' Declarations
 Dim intercept As Double
 Dim segmentsCrossed As Double
 Dim x1 As Double
 Dim x2 As Double
 Dim y1 As Double
 Dim y2 As Double
 Dim wholePoly() As Variant
 Dim polyCounter As Long
 
 segmentsCrossed = 0
 wholePoly = ClosePoly(polyGon)
 
 For polyCounter = LBound(wholePoly) To UBound(wholePoly) - 1 ' Test
   y2 = wholePoly(polyCounter + 1, 2)
   y1 = wholePoly(polyCounter, 2)
   x2 = wholePoly(polyCounter + 1, 1)
   x1 = wholePoly(polyCounter, 1)
   If x1 > xHit Xor x2 > xHit Then ' test for crossing only if xHit is between vertices of a segment
     slope = (y2 - y1) / (x2 - x1) 'slope of segment
     intercept = y1 - slope * x1 ' intercept of segment
     If slope * xHit + intercept > yHit Then segmentsCrossed = segmentsCrossed + 1
   End If
 Next polyCounter

 PointInPolygon = (segmentsCrossed) Mod 2 ' Return 0 FALSE or 1 TRUE
End Function

'***************************************************************************
'Procedure: ClosePoly
'Purpose: Make the last row of the array the same as the first by adding it to the end IF NOT the same
'Comments: Requires Option Base 1
'          Array is format varPolygon(x,y)
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2018  Chris Kinion      Created
'***************************************************************************
Function ClosePoly(inputPoly As Variant) As Variant
 Dim lineCount As Double
 Dim returnPoly As Variant
 If Not (inputPoly(1, 1) = inputPoly(UBound(inputPoly), 1) _
   And inputPoly(1, 2) = inputPoly(UBound(inputPoly), 2)) Then
     ReDim returnPoly(1 To UBound(inputPoly) + 1, 2)
     For lineCount = 1 To UBound(inputPoly)
       returnPoly(lineCount, 1) = inputPoly(lineCount, 1)
       returnPoly(lineCount, 2) = inputPoly(lineCount, 2)
     Next lineCount
     returnPoly(UBound(returnPoly), 1) = returnPoly(1, 1)
     returnPoly(UBound(returnPoly), 2) = returnPoly(1, 2)
 Else
   returnPoly = inputPoly
 End If
 ClosePoly = returnPoly
End Function

'***************************************************************************
'Procedure: GetTslope
'Purpose: Returns the slope perpendicular to the line defined by (x1,y1) and (x2,y2)
'Comments: Uses "very small numbers" in case of divide by zero
'Changes----------------------------------------------
' Date        Programmer        Change
' 09/01/2018  Chris Kinion      Created
'***************************************************************************
Function GetTslope(x1 As Double, x2 As Double, y1 As Double, y2 As Double) As Double
 Dim denominator As Double
 Dim numerator As Double
 Dim slope As Double
 Dim tslope As Double

 denominator = x2 - x1
   If denominator = 0 Then
     denominator = 1E-46
   End If
 numerator = y2 - y1
   If numerator = 0 Then
     numerator = 1E-46
   End If
 slope = numerator / denominator
 tslope = -1 / slope
 
 GetTslope = tslope
End Function



