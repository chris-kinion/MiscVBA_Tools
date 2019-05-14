Attribute VB_Name = "Math_Random_16bit"
Option Explicit
Option Base 1

Dim s1 As Integer
Dim s2 As Integer
Dim s3 As Integer

Sub verifyS1(ByRef myVar As Integer)
  If Abs(myVar) <= 1 Then myVar = 2
  If Abs(myVar) >= 323621 Then myVar = 323620
End Sub
Sub verifyS2(ByRef myVar As Integer)
  If Abs(myVar) <= 1 Then myVar = 2
  If Abs(myVar) >= 317261 Then myVar = 317260
End Sub
Sub verifyS3(ByRef myVar As Integer)
  If Abs(myVar) <= 1 Then myVar = 2
  If Abs(myVar) >= 316561 Then myVar = 316560
End Sub

' Generates uniformly distributed random numbers between 0 and 1
' Based on: Figure 4. A Portable Generator for 16-bit Computers
' L’Ecuyer. P. Efficient and portable combined random number generators.
' Communications of the ACM 31, 6 (June 1988) p. 748.
' This program uses the Integer data type. Therefore, possible seed values range from –32,768 to 32,767.
Function Uniform() As Double
  Dim Z As Integer
  Dim k As Integer
  
  Call verifyS1(s1)
  Call verifyS2(s2)
  Call verifyS3(s3)
  
  k = s1 \ 206
  s1 = 157 * (s1 - k * 206) - k * 21
  If s1 < 0 Then s1 = s1 + 32363
  
  k = s2 \ 217
  s2 = 146 * (s2 - k * 217) - k * 45
  If s2 < 0 Then s2 = s2 + 31727
  
  k = s3 \ 222
  s3 = 142 * (s3 - k * 222) - k * 133
  If s3 < 0 Then s3 = s3 + 31657
  
  Z = s1 - s2
  If Z > 706 Then Z = Z - 32362
  Z = Z + s3
  If Z < 1 Then Z = Z + 32362
  
  Uniform = Z * 0.000030899
End Function

' Test output for Uniform function
' Caution! This will overwrite data!
Sub testUniform()
  s1 = 100
  s2 = 100
  s3 = 101
  Dim thisWorksheet As Worksheet
  Set thisWorksheet = ThisWorkbook.Worksheets(1)
  
  Dim i As Long
  For i = 1 To 15 ' 15 Examples of random numbers
    With thisWorksheet
      .Cells(i, 1) = i
      .Cells(i, 2) = Uniform
      'Debug.Print Uniform
    End With
  Next i
End Sub

Sub testUniform2()
  Dim thisWorksheet As Worksheet
  Set thisWorksheet = ThisWorkbook.Worksheets(1)
  With thisWorksheet
    On Error Resume Next
    s1 = CInt(.Cells(2, 5).Value)
    s2 = CInt(.Cells(3, 5).Value)
    s3 = CInt(.Cells(4, 5).Value)
    If Err.Number <> 0 Then
      MsgBox "Data Entry Validation", vbExclamation, "Ensure cells E2:E4 are whole numbers between -32,768 and 32,767 not including 0"
      On Error GoTo 0
      Exit Sub
    End If
  End With
  
  Dim i As Long
  Dim maxIterations As Long
  With thisWorksheet
    If .Cells(1, 5).Value >= 1 And Abs(.Cells(1, 5)) < 32767 Then
      maxIterations = CLng(.Cells(1, 5).Value)
    Else
      maxIterations = 1
    End If
  End With
  
  thisWorksheet.Columns("A:B").Clear
  
  For i = 1 To maxIterations ' 15 Examples of random numbers
    With thisWorksheet
      .Cells(i, 1) = i
      .Cells(i, 2) = Uniform
      'Debug.Print Uniform
    End With
  Next i
End Sub
