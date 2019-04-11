Attribute VB_Name = "Manipulate_Worksheets"
'***************************************************************************
'Module: Manipulate_Worksheets
'Procedures: addBlankRowsMacro: insert blank rows between each row of contiguous data
            ' getA1address: gets the A1 address for various cells
            ' makeThisCellCool: Modify cell features
'Comments: Part tools, part demonstrative
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Amalgamated module content
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: addBlankRowsMacro
'Purpose: insert blank rows between each row of contiguous data
'Comments: starting point must be upper-left corner of data area to spread and at least two columns wide
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Sub addBlankRowsMacro()
  Dim totalTableRows As Long
  Dim i As Long
  totalTableRows = ActiveCell.End(xlDown).Row - ActiveCell.Row ' Number of rows underneath to move
  For i = 1 To totalTableRows - 1
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlToRight)).Select
    On Error Resume Next
      Range(Selection, Selection.End(xlDown)).Select
    If Error <> 0 Then
      Debug.Print "Down selection error"
    End If
    On Error GoTo 0
    Selection.Cut
    ActiveCell.Offset(1, 0).Select
    On Error Resume Next
      ActiveSheet.Paste
    If Err.Number <> 0 Then
      Debug.Print "Pasting error"
    End If
    On Error GoTo 0
  Next i
  
  ActiveCell.Offset(1, 0).Select ' Last row
  Range(Selection, Selection.End(xlToRight)).Select
  Selection.Cut
  ActiveCell.Offset(1, 0).Select
  ActiveSheet.Paste
End Sub

'***************************************************************************
'Procedure: getA1address
'Purpose: gets the A1 address for various cells
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Sub getA1address()
  Dim myCellAddress As Variant, myLastRowAddress  As Variant, myLastColAddress As Variant, myVeryLastRowAddress  As Variant, myFirstBlankAddress  As Variant
  myCellAddress = ActiveCell.Address
  myLastRowAddress = Range("A5").End(xlDown).Address
  myLastColAddress = Range("A5").End(xlToRight).Address
  myVeryLastRowAddress = Range("A1048576").End(xlUp).Address
  myFirstBlankAddress = Range("A1048576").End(xlUp).Offset(1, 0).Address
  
  Debug.Print "myCellAddress " & myCellAddress
  Debug.Print "myLastRowAddress " & myLastRowAddress
  Debug.Print "myLastColAddress " & myLastColAddress
  Debug.Print "myVeryLastRowAddress " & myVeryLastRowAddress
  Debug.Print "myFirstBlankAddress " & myFirstBlankAddress
End Sub

'***************************************************************************
'Procedure: makeThisCellCool
'Purpose: Modify cell features
'Comments: Can just as easily get these features by recording a macro
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Sub makeThisCellCool()
  With ActiveCell.Font
        .Name = "Courier New"
        .Size = 20
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
  End With
End Sub
