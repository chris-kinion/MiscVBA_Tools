Attribute VB_Name = "Sample_ShapeAlignment"
'***************************************************************************
'Module: Sample_ShapeAlignment
'Procedures: formatTutorialShapes: Format Tutorial Navigation Shapes
          '  thisBoxFill: formats shape fill color
'Comments: Example code
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/11/2019  Chris Kinion      Initial amalgamation
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: formatTutorialShapes
'Purpose: Format All Tutorial shapes (includes left arrows, right arrows, text boxes)
'Calls: thisBoxFill
'Comments: Navigation between shapes used macros to hide and display other shapes
'Changes----------------------------------------------
' Date        Programmer        Change
' 08/02/2018  Chris Kinion      Created
'***************************************************************************
Sub formatTutorialShapes()
  Dim myTextBoxNames()
  Dim arrowLeftNames()
  Dim arrowRightNames()
  Dim i As Integer
  Dim paneLaunch As String, paneStart As String, paneExit As String
  
  myTextBoxNames = Array("Pane_Intro", "1_0", "1_1", "1_2", "1_3", "1_4", "1_5", "1_6", "1_7", "2_0", "2_1", "2_2", "2_3", "3_0", "3_1", "3_2", "3_3", "3_4", "3_5", "3_6", "3_7", "3_8", "3_9", "3_10", "3_11", "4_0", "4_1", "4_2", "4_3", "4_4", "4_5", "5_0", "5_1", "5_2", "5_3", "5_4")
  arrowLeftNames = Array("Arrow: Left 1.0", "Arrow: Left 1.1", "Arrow: Left 1.2", "Arrow: Left 1.3", "Arrow: Left 1.4", "Arrow: Left 1.5", "Arrow: Left 1.6", "Arrow: Left 1.7", "Arrow: Left 2.0", "Arrow: Left 2.1", "Arrow: Left 2.2", "Arrow: Left 2.3", "Arrow: Left 3.0", "Arrow: Left 3.1", "Arrow: Left 3.2", "Arrow: Left 3.3", "Arrow: Left 3.4", "Arrow: Left 3.5", "Arrow: Left 3.6", "Arrow: Left 3.7", "Arrow: Left 3.8", "Arrow: Left 3.9", "Arrow: Left 3.10", "Arrow: Left 3.11", "Arrow: Left 4.0", "Arrow: Left 4.1", "Arrow: Left 4.2", "Arrow: Left 4.3", "Arrow: Left 4.4", "Arrow: Left 4.5", "Arrow: Left 5.0", "Arrow: Left 5.1", "Arrow: Left 5.2", "Arrow: Left 5.3")
  arrowRightNames = Array("Arrow: Right 1.1", "Arrow: Right 1.2", "Arrow: Right 1.3", "Arrow: Right 1.4", "Arrow: Right 1.5", "Arrow: Right 1.6", "Arrow: Right 1.7", "Arrow: Right 2.0", "Arrow: Right 2.1", "Arrow: Right 2.2", "Arrow: Right 2.3", "Arrow: Right 3.0", "Arrow: Right 3.1", "Arrow: Right 3.2", "Arrow: Right 3.3", "Arrow: Right 3.4", "Arrow: Right 3.5", "Arrow: Right 3.6", "Arrow: Right 3.7", "Arrow: Right 3.8", "Arrow: Right 3.9", "Arrow: Right 3.10", "Arrow: Right 3.11", "Arrow: Right 4.0", "Arrow: Right 4.1", "Arrow: Right 4.2", "Arrow: Right 4.3", "Arrow: Right 4.4", "Arrow: Right 4.5", "Arrow: Right 5.0", "Arrow: Right 5.1", "Arrow: Right 5.2", "Arrow: Right 5.3", "Arrow: Right 5.4", "Arrow: Right End")
  
  For i = 1 To ActiveSheet.Shapes.Count
    ActiveSheet.Shapes(i).Visible = msoTrue
  Next i
  
  paneExit = "Pane_Exit"
    Sheet1.Shapes(paneExit).Select
    Selection.ShapeRange.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    Sheet1.Shapes(paneExit).Top = Range("G7").Top + 5
    Sheet1.Shapes(paneExit).Left = Range("G7").Left + 36
    thisBoxFill
  paneStart = "Pane_Start"
    Sheet1.Shapes(paneStart).Select
    Selection.ShapeRange.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    Sheet1.Shapes(paneStart).Top = Range("G5").Top + 5
    Sheet1.Shapes(paneStart).Left = Range("G5").Left + 36
    thisBoxFill
  paneLaunch = "Pane_Launch"
    Sheet1.Shapes(paneLaunch).Select
    Selection.ShapeRange.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    Sheet1.Shapes(paneLaunch).Top = Range("G5").Top + 5
    Sheet1.Shapes(paneLaunch).Left = Range("G5").Left + 14
    thisBoxFill
  For i = LBound(myTextBoxNames) To UBound(myTextBoxNames) ' Text Boxes
    Sheet1.Shapes(myTextBoxNames(i)).Select
    Sheet1.Shapes(myTextBoxNames(i)).Top = Range("I1").Top + 2
    Sheet1.Shapes(myTextBoxNames(i)).Left = Range("I1").Left
    Selection.ShapeRange.Height = 108
    Selection.ShapeRange.Width = 180
    thisBoxFill
  Next i
  
  For i = LBound(arrowLeftNames) To UBound(arrowLeftNames) ' Left Arrows
    Sheet1.Shapes(arrowLeftNames(i)).Select
    With Selection
      .Top = Range("G5").Top + 4
      .Left = Range("G5").Left + 10
      .ShapeRange.Height = 35
      .ShapeRange.Width = 45
    End With
    thisBoxFill
  Next i
  
  For i = LBound(arrowRightNames) To UBound(arrowRightNames) ' Right Arrows
    Sheet1.Shapes(arrowRightNames(i)).Select
    With Selection
      .Top = Range("G5").Top + 4
      .Left = Range("G5").Left + 60
      .ShapeRange.Height = 35
      .ShapeRange.Width = 45
    End With
    thisBoxFill
  Next i

End Sub

'***************************************************************************
'Procedure: thisBoxFill
'Purpose: formats shape fill color
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 08/02/2018  Chris Kinion      Created
'***************************************************************************
Sub thisBoxFill()
    With Selection.ShapeRange.Fill
      .Visible = msoTrue
      .ForeColor.ObjectThemeColor = msoThemeColorAccent6
      .ForeColor.TintAndShade = 0
      .ForeColor.Brightness = 0.400000006
      .Transparency = 0
      .Solid
    End With
End Sub

