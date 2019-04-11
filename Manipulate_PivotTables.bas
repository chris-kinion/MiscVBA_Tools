Attribute VB_Name = "Manipulate_PivotTables"
'***************************************************************************
'Module: Manipulate_PivotTables
'Procedures: removeAllPivots: Remove all pivot tables from active worksheet
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 08/27/2018  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: removeAllPivots
'Purpose:  Remove all pivot tables from active worksheet
'Changes----------------------------------------------
' Date        Programmer        Change
' 08/27/2018  Chris Kinion      Created
'***************************************************************************
Sub removeAllPivots()
  Dim pvtTable As PivotTable
  Dim wksPivot As Worksheet
  Set wksPivot = ThisWorkbook.ActiveSheet
  For Each pvtTable In wksPivot.PivotTables ' Remove old pivot tables
    pvtTable.TableRange2.Clear
  Next pvtTable
End Sub
