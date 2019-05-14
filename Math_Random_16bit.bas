Attribute VB_Name = "Math_Random_16bit"
Option Explicit
Option Base 1

Dim s1 As Integer ' Make Public to use across modules
Dim s2 As Integer ' Make Public to use across modules
Dim s3 As Integer ' Make Public to use across modules

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


