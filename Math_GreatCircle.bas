Attribute VB_Name = "Math_GreatCircle"
'***************************************************************************
'Module: Math_GreatCircle
'Procedures: GreatCircleDistanceRadians: Finds great circle distance in radians between two points given latitude and longitude of those points
'            DegreesToRadians: Converts degrees to radians
'            GreatCircleLatitudeDistanceFeet: Get distance between two latitudes given a starting longitude
'            GreatCircleLongitudeDistanceFeet: Get distance between two longitudes given a starting longitude
'            GreatCircleNauticalMilesFromRadians: Converts a distance in radians on a great circle to Nautical Miles
'            GreatCircleMilesFromRadians: Converts a distance in radians on a great circle to Miles
'            GreatCircleFeetFromRadians: Converts a distance in radians on a great circle to Feet
'            ConvertPositionToAxis: Return distance of 1° of latitude or longitude based on a key latitude
'            ReturnXY: Return a latitude/longitude array converted to an x/y coordinate system
'Comments: All procedures in this module are Functions
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2019  Chris Kinion      Created
' 11/21/2018  Chris Kinion      Added Great Circle distance functions
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: GreatCircleDistanceRadians
'Purpose: Finds great circle distance to radians from two points given the
'         latitude and longitude of those points
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 11/21/2018  Chris Kinion      Created
'***************************************************************************
Function GreatCircleDistanceRadians(startLatitude As Double, startLongitude As Double, endLatitude As Double, endLongitude As Double) As Double
  Dim Lat1 As Double
  Dim Lat2 As Double
  Dim Lon1 As Double
  Dim Lon2 As Double
  Lat1 = DegreesToRadians(startLatitude)
  Lat2 = DegreesToRadians(endLatitude)
  Lon1 = DegreesToRadians(startLongitude)
  Lon2 = DegreesToRadians(endLongitude)
  GreatCircleDistanceRadians = 2 * WorksheetFunction.Asin(Sqr((Sin((Lat1 - Lat2) / 2)) ^ 2 + Cos(Lat1) * Cos(Lat2) * (Sin((Lon1 - Lon2) / 2)) ^ 2))
End Function

'***************************************************************************
'Procedure: DegreesToRadians
'Purpose: Converts degrees to radians
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 11/21/2018  Chris Kinion      Created
'***************************************************************************
Function DegreesToRadians(degreesIn As Double) As Double
  DegreesToRadians = (degreesIn / 180) * WorksheetFunction.Pi
End Function

'***************************************************************************
'Procedure: GreatCircleLatitudeDistanceFeet
'Purpose: Get distance between two latitudes given a starting longitude
'Comments: requres GreatCircleDistanceRadians
'Changes----------------------------------------------
' Date        Programmer        Change
' 11/21/2018  Chris Kinion      Created
'***************************************************************************
Function GreatCircleLatitudeDistanceFeet(startLatitude As Double, startLongitude As Double, endLatitude As Double) As Double
  Dim thisDistance As Double
  thisDistance = GreatCircleDistanceRadians(startLatitude, startLongitude, endLatitude, startLongitude)
  thisDistance = thisDistance * 180 * 60 / WorksheetFunction.Pi
  thisDistance = thisDistance * 1.150779 * 5280
  GreatCircleLatitudeDistanceFeet = thisDistance
End Function

'***************************************************************************
'Procedure: GreatCircleLongitudeDistanceFeet
'Purpose: Get distance between two longitudes given a starting longitude
'Comments: requres GreatCircleDistanceRadians
'Changes----------------------------------------------
' Date        Programmer        Change
' 11/21/2018  Chris Kinion      Created
'***************************************************************************
Function GreatCircleLongitudeDistanceFeet(startLatitude As Double, startLongitude As Double, endLongitude As Double) As Double
  Dim thisDistance As Double
  thisDistance = GreatCircleDistanceRadians(startLatitude, startLongitude, startLatitude, endLongitude)
  thisDistance = thisDistance * 180 * 60 / WorksheetFunction.Pi
  thisDistance = thisDistance * 1.150779 * 5280
  GreatCircleLongitudeDistanceFeet = thisDistance
End Function

'***************************************************************************
'Procedure: GreatCircleNauticalMilesFromRadians
'Purpose: Converts a distance in radians on a great circle to Nautical Miles
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 11/21/2018  Chris Kinion      Created
'***************************************************************************
Function GreatCircleNauticalMilesFromRadians(RadiansValue As Double) As Double
  GreatCircleNauticalMilesFromRadians = RadiansValue * 180 * 60 / WorksheetFunction.Pi
End Function

'***************************************************************************
'Procedure: GreatCircleMilesFromRadians
'Purpose: Converts a distance in radians on a great circle to Miles
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 11/21/2018  Chris Kinion      Created
'***************************************************************************
Function GreatCircleMilesFromRadians(RadiansValue As Double) As Double
  GreatCircleMilesFromRadians = 1.150779 * RadiansValue * 180 * 60 / WorksheetFunction.Pi
End Function

'***************************************************************************
'Procedure: GreatCircleFeetFromRadians
'Purpose: Converts a distance in radians on a great circle to Feet
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 11/21/2018  Chris Kinion      Created
'***************************************************************************
Function GreatCircleFeetFromRadians(RadiansValue As Double) As Double
  GreatCircleFeetFromRadians = 5280 * 1.150779 * RadiansValue * 180 * 60 / WorksheetFunction.Pi
End Function

'***************************************************************************
'Procedure: ConvertPositionToAxis
'Purpose:   Return distance of 1° of latitude or longitude based on a key latitude
'Comments:  lat1_long2 input should be 1 to return 1° of latitude or 2 to return 1° of longitude
          ' meter1_feet2 input should be 1 to return distance in meters or 2 to return distance in feet
          ' Conversions for Standard Miles (SM) and Nautical Miles (NM) are commented
          ' Requires Option Base 1
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2019  Chris Kinion      Created
'***************************************************************************
Function ConvertPositionToAxis(inLat As Double, lat1_lon2 As Integer, meter1_feet2 As Integer) As Double
  Dim m1 As Double
  Dim m2 As Double
  Dim m3 As Double
  Dim m4 As Double
  Dim p1 As Double
  Dim p2 As Double
  Dim p3 As Double
  Dim latDegrees As Double
  Dim latRad As Double
  Dim latLen As Double
  Dim longLen As Double
  Dim latMetersConv As Double
  Dim latFeetConv As Double
  Dim longMetersConv As Double
  Dim longFeetConv As Double
  Dim returnPosition() As Double
  
  ' Standard/Nautical Miles for future use
  ' Dim latSMConv As Double
  ' Dim latNMConv As Double
  ' Dim longSMConv As Double
  ' Dim longNMConv As Double
   
  If lat1_lon2 <> (1 Or 2) Then
    ConvertPositionToAxis = 0
  End If
  
  m1 = 111132.92 ' Constants
  m2 = -559.82
  m3 = 1.175
  m4 = -0.0023
  p1 = 111412.84
  p2 = -93.5
  p3 = 0.118
    
  latDegrees = inLat
  latRad = latDegrees * (2 * WorksheetFunction.Pi) / 360
  latLen = m1 + (m2 * Cos(2 * latRad)) + (m3 * Cos(4 * latRad)) + (m4 * Cos(6 * latRad))
  longLen = (p1 * Cos(latRad)) + (p2 * Cos(3 * latRad)) + (p3 * Cos(5 * latRad))
  
  latMetersConv = latLen
  latFeetConv = (latLen / 12) * 39.370079
  'latSMConv = latFeetConv / 5280
  'latNMConv = latSMConv / 1.15077945
  
  longMetersConv = longLen
  longFeetConv = (longLen / 12) * 39.370079
  'longSMConv = longFeetConv / 5280
  'longNMConv = longSMConv / 1.15077945
  
  ReDim returnPosition(2) ' Adjust for units desired
  If meter1_feet2 = 1 Then
    returnPosition(1) = (latMetersConv)
    returnPosition(2) = (longMetersConv)
  ElseIf meter1_feet2 = 2 Then
    returnPosition(1) = (latFeetConv)
    returnPosition(2) = (longFeetConv)
  Else
    returnPosition(1) = 0
    returnPosition(2) = 0
  End If
  
  ConvertPositionToAxis = returnPosition(lat1_lon2)
End Function

'***************************************************************************
'Procedure: ReturnXY
'Purpose: Return a latitude/longitude array converted to an x/y coordinate system
'Inputs:  toConvert: 2-column array, e.g.: toConvert(col_1,col_2)
        ' lat_unit: conversion unit of latitude (see UDF ConvertPositionToAxis)
        ' long_unit: conversion unit of longitude (see UDF ConvertPositionToAxis)
        ' column_lat: 1 if column 1 is latitude, 2 if column 2 is latitude
        ' column_long: 1 if column 1 is longitude, 2 if column 2 is longitude
        ' lat_origin: latitude of origin coordinates
        ' long_origin: longitude of origin coordinates
'Comments: Requires Option Base 1
        ' UDF ConvertPositionToAxis returns conversion unit (feet or meters per 1° of lat/lon)
        ' Not useful for large distances as the conversion factor changes
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2019  Chris Kinion      Created
'***************************************************************************
Function ReturnXY(toConvert As Variant, lat_unit As Double, long_unit As Double, column_lat As Double, column_long As Double, lat_origin As Double, long_origin As Double) As Variant
  Dim newXY() As Variant
  Dim xyCount As Integer
  ReDim newXY(LBound(toConvert) To UBound(toConvert), 2)
  
  For xyCount = LBound(toConvert) To UBound(toConvert)
    newXY(xyCount, column_lat) = lat_unit * (toConvert(xyCount, column_lat) - lat_origin)
    newXY(xyCount, column_long) = long_unit * (toConvert(xyCount, column_long) - long_origin)
  Next xyCount
 
  ReturnXY = newXY
End Function
