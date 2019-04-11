Attribute VB_Name = "Manipulate_Charts"
'***************************************************************************
'Module: Manipulate_Charts
'Procedures:  reviewSeriesCollection
            ' deleteLastSeries
            ' doesThisSeriesExist
            ' countSeriesInMyChart
            ' showSeriesNamesInMyChart
'Comments: These are primarily tools used that print to Immediate Window
'Changes----------------------------------------------
' Date        Programmer        Change
' 05/31/2018  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: reviewSeriesCollection
'Purpose: Displays named series in the chart "myChart" to Immediate Window
'Changes----------------------------------------------
' Date        Programmer        Change
' 05/31/2018  Chris Kinion      Created
'***************************************************************************
Sub reviewSeriesCollection()
  Dim myCounter As Long
  myCounter = 0
  With ActiveSheet
    .Shapes("myChart").Select
  End With ' ActiveSheet
  With ActiveChart
    For Each namedRange In .FullSeriesCollection
      myCounter = myCounter + 1
      Debug.Print namedRange.Name & " " & myCounter
    Next namedRange
    Debug.Print myCounter & " named ranges"
  End With ' ActiveChart
End Sub

'***************************************************************************
'Procedure: deleteLastSeries
'Purpose: Deletes the last added series to chart "myChart"
'Changes----------------------------------------------
' Date        Programmer        Change
' 05/31/2018  Chris Kinion      Created
'***************************************************************************
Sub deleteLastSeries()
  With ActiveSheet
    .Select
    .Range("E6").Select
    .Shapes("myChart").Select
  End With
  
  With ActiveChart
    Dim seriesNum As Long
    seriesNum = .SeriesCollection.Count
      Debug.Print seriesNum & " named ranges"
      Debug.Print "Deleting: '" & .FullSeriesCollection(seriesNum).Name & "'..."
    .FullSeriesCollection(seriesNum).Delete
    seriesNum = .SeriesCollection.Count
      Debug.Print "Deleted. Now only " & seriesNum & " named ranges"
  End With ' ActiveChart
End Sub

'***************************************************************************
'Procedure: doesThisSeriesExist
'Purpose: Looking to affirm series "mySeries" exists in chart "myChart"
'Changes----------------------------------------------
' Date        Programmer        Change
' 05/31/2018  Chris Kinion      Created
'***************************************************************************
Sub doesThisSeriesExist()
  Dim mySearchVar As String
  mySearchVar = "mySeries"
  Dim foundMySeries As Boolean
  foundMySeries = False
  
  With ActiveSheet
    .Select
    .Range("E6").Select
    .Shapes("myChart").Select
  End With ' ActiveSheet
  
  With ActiveChart
    Dim namedRange As Variant
    For Each namedRange In .FullSeriesCollection
      If Not StrComp(namedRange.Name, mySearchVar, vbTextCompare) <> 0 Then
        foundMySeries = True
      End If
      Debug.Print StrComp(namedRange.Name, mySearchVar, vbTextCompare) & ", " & foundMySeries
    Next namedRange
    Debug.Print foundMySeries
  End With ' ActiveChart
End Sub

'***************************************************************************
'Procedure: countSeriesInMyChart
'Purpose: Place in Immediate Window how may series are in a selected chart
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 05/31/2018  Chris Kinion      Created
'***************************************************************************
Sub countSeriesInMyChart()
  Debug.Print ActiveChart.SeriesCollection.Count
End Sub

'***************************************************************************
'Procedure: showSeriesNamesInMyChart
'Purpose: List all names of a charts series
'Changes----------------------------------------------
' Date        Programmer        Change
' 05/31/2018  Chris Kinion      Created
'***************************************************************************
Sub showSeriesNamesInMyChart()
  Dim xyz As Variant
  For Each xyz In ActiveChart.SeriesCollection
    Debug.Print xyz.Name
  Next xyz
End Sub


