Attribute VB_Name = "Math_Random_BoxMuller"
Function BoxMullerTransform() As Double
' Converts a pair of uniform random numbers to a normally distributed random number
  Dim x2 As Double
  'x2 = Sqr(x1) ' square root
  'x2 = Log(x1) ' natural log
  'x2 = Rnd() ' Random number between 0 and 1
  'x2 = Application.WorksheetFunction.Pi() ' pi
  x2 = (Sqr(-2 * Log(Rnd()))) * (Cos(2 * Application.WorksheetFunction.Pi() * Rnd()))
  BoxMullerTransform = x2
End Function

' Converts a pair of uniform random numbers to a normally distributed random number
' using the Box-Muller Transform and Excel / VBA uniform random number generator
Function RandomBMT() As Double
  RandomBMT = (Sqr(-2 * Log(Rnd()))) * (Cos(2 * Application.WorksheetFunction.Pi() * Rnd()))
End Function

Sub RecalculateColumnA()
  With ThisWorkbook.ActiveSheet
    .Range("D4") = Rnd()
  End With
End Sub

Sub RecalculateColumnB()
  With ThisWorkbook.ActiveSheet
    .Range("E5") = Rnd()
  End With
End Sub
