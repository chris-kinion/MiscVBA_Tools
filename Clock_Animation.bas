Attribute VB_Name = "Clock_Animation"
Sub timeFrame(pauseTime As Double)
  Dim Start As Single
  Start = Timer
  Do
  DoEvents
  Loop Until (Timer - Start) >= pauseTime
End Sub

Sub testClock()
  ThisWorkbook.Worksheets("Clock").Activate
  ActiveSheet.Range("A1").Select
  
  Dim wholeHour As Long
  Dim halfHour As Long
  Dim i As Long
  
  wholeHour = 128336
  halfHour = 128348

  With ActiveSheet
    For i = 0 To 11
      .Range("B1") = Application.Unichar(wholeHour + i)
      Call timeFrame(0.5)
      .Range("B1") = Application.Unichar(halfHour + i)
      Call timeFrame(0.5)
    Next i
  End With
End Sub

Sub testClock2()
  ThisWorkbook.Worksheets("Clock").Activate
  ActiveSheet.Range("A1").Select
  
  Dim wholeHour As Long
  Dim halfHour As Long
  Dim i As Long, j As Long
  
  wholeHour = 128336
  halfHour = 128348

  With ActiveSheet
    For j = 1 To 2
      For i = 0 To 11
        .Range("E1") = Application.Unichar(wholeHour + i)
        Call timeFrame(0.5)
        '.Range("C1") = Application.Unichar(halfHour + i)
        'Call timeFrame(0.5)
      Next i
    Next j
  End With
End Sub
