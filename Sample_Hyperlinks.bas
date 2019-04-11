Attribute VB_Name = "Sample_Hyperlinks"
'***************************************************************************
'Module: Sample_Hyperlinks
'Procedures:  addMenuHyperlinks
            ' addJustThisOneHyperlink
'Comments: Sample code
          ' I found hyperlinks to be unstable and unreliable within a spreadsheet
'           but work better when linked to webpages, etc
'Changes----------------------------------------------
' Date        Programmer        Change
' 08/01/2018  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: addMenuHyperlinks
'Purpose: Add hyperlink to same cell within same sheet, controls formatting
'Changes----------------------------------------------
' Date        Programmer        Change
' 08/01/2018  Chris Kinion      Created
'***************************************************************************
Sub addMenuHyperlinks()
  Dim myRows As Integer
  Dim myCols As Integer
  Dim myCellAddress As Variant
  Dim myCellValue As Variant
  
  With ActiveSheet
    For myRows = 4 To 6
      For myCols = 2 To 9
        If .Cells(myRows, myCols) <> "" Then
          .Cells(myRows, myCols).Select
          myCellValue = ActiveCell.Value
          myCellAddress = ActiveCell.Address
          .Hyperlinks.Add anchor:=.Cells(myRows, myCols), Address:="", SubAddress:=myCellAddress, TextToDisplay:=myCellValue, ScreenTip:=myCellValue
          With .Cells(myRows, myCols).Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.499984740745262
            .Underline = xlUnderlineStyleNone
          End With ' cells.font
        End If
      Next myCols
    Next myRows
  End With ' ActiveSheet
End Sub

'***************************************************************************
'Procedure: addJustThisOneHyperlink
'Purpose: Add and format a hyperlink
'Changes----------------------------------------------
' Date        Programmer        Change
' 08/01/2018  Chris Kinion      Created
'***************************************************************************
Sub addJustThisOneHyperlink()
  Dim myCellAddress As Variant
  Dim myCellValue As Variant
  
  myCellValue = ActiveCell.Value
  myCellAddress = ActiveCell.Address
  
  With ActiveSheet
    .Hyperlinks.Add anchor:=Selection, Address:="", SubAddress:=myCellAddress, TextToDisplay:=myCellValue, ScreenTip:=myCellValue
  End With
  
  With ActiveCell.Font
    .Underline = xlUnderlineStyleNone
    .Color = RGB(0, 0, 0)
  End With
End Sub
