Attribute VB_Name = "Event_Worksheet"
'***************************************************************************
'Module: Event_Worksheet
'Procedures:  DisplayPageFromTop
            ' findRowOfValue
            ' clearThenHideSheet
            ' FoundSheetName
            ' WriteOrClearSheet
            ' ContinueWithDeletion
            ' DeleteSheetByName
            ' DeleteThisSheet
            ' wksExists
            ' AddBlankSheetByName
            ' FreezeFirstRow
            ' FreezeFirstColumn
            ' FreezeFirstRowAndColumn
            ' ShowAllSheets
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Created, amalgamated proceedures
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: DisplayPageFromTop
'Purpose: Brings to view of top left of page
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/12  Chris Kinion      Created
'***************************************************************************
Sub DisplayPageFromTop(wksDestinationSheet As Worksheet)
  wksDestinationSheet.Activate
  wksDestinationSheet.Range("A1").Select
  ActiveWindow.ScrollRow = 1
End Sub

'***************************************************************************
'Procedure: findRowOfValue
'Purpose: Return the row number of a number found in a range
'Inputs:  findVal: long (number for which searching)
        ' rangeVal: string (area to search)
        ' searchSheet: worksheet to search
'Comments: Returns last row number if cannot find the value
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/18/2018  Chris Kinion      Created
'***************************************************************************
Function findRowOfValue(findVal As Long, rangeVal As String, searchSheet As Worksheet) As Long
  With searchSheet
    On Error Resume Next
    findRowOfValue = .Range(rangeVal).Find(findVal).Row
    If Err.Number <> 0 Then
      Debug.Print "Value " & findVal & ", Range " & rangeVal & " at sheet " & searchSheet
      findRowOfValue = searchSheet.Rows.Count
    End If
  End With
End Function

'***************************************************************************
'Procedure: clearThenHideSheet
'Purpose: clear and hide a sheet
'Changes----------------------------------------------
' Date        Programmer        Change
' 10/4/2018   Chris Kinion      Created
'***************************************************************************
Sub clearThenHideSheet(wks As Worksheet)
  wks.Cells.Clear
  wks.Visible = xlSheetHidden
End Sub

'***************************************************************************
'Procedure: FoundSheetName(strName As String) As Boolean
'Purpose: Return if found exact string match for sheet name
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/11/2019  Chris Kinion      Created
'***************************************************************************
Function FoundSheetName(strName As String) As Boolean
  Dim wksSheet As Worksheet
  FoundSheetName = False
  With ThisWorkbook
    For Each wksSheet In .Worksheets
      If wksSheet.Name = strName Then
        FoundSheetName = True
      End If
    Next wksSheet
  End With
End Function

'***************************************************************************
'Procedure: WriteOrClearSheet(ByVal strNewName As String)
'Purpose: Clear existing sheet or write new one given sheet name
'Calls: FoundSheetName
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/11/2019  Chris Kinion      Created
'***************************************************************************
Sub WriteOrClearSheet(ByVal strNewName As String)
  With ThisWorkbook
    If FoundSheetName(strNewName) Then
      .Worksheets(strNewName).Cells.Clear
    Else
      .Sheets.Add(after:=.Sheets(.Sheets.Count)).Name = strNewName
    End If
  End With
End Sub

'***************************************************************************
'Procedure: ContinueWithDeletion() As Boolean
'Purpose: Returns a Boolean to confirm a user wants to delete old data sheets
'Comments:  Can re-purpose with any message
          ' Can code inline since vbYes = 6, vbNo = 7
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/12/2019  Chris Kinion      Created
'***************************************************************************
Function ContinueWithDeletion() As Boolean
  Dim intDecision As Integer
  Dim strContinueMessage As String
  
  ContinueWithDeletion = False
  strContinueMessage = "Do you wish to delete sheets with old base data? This cannot be undone." ' vbyesno
  
  intDecision = MsgBox(strContinueMessage, vbYesNo, "Warning!")
  If intDecision = 6 Then
    ContinueWithDeletion = True
  End If
  Debug.Print intDecision
End Function

'***************************************************************************
'Procedure: DeleteSheetByName(ByVal strPart As String, blnContinue As Boolean)
'Purpose: Deletes sheets which contain the input string
'Comments:  Turns off application display alerts temporarily
          ' blnContinue could be linked to a variable that allows sheet deletion without warning
          ' Intended to remove automatically generated worksheets that were written to a specific format
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/12/2019  Chris Kinion      Created
'***************************************************************************
Sub DeleteSheetByName(ByVal strPart As String, blnContinue As Boolean)
  If blnContinue Then
    Dim wksSheet As Worksheet
    Application.DisplayAlerts = False
    With ThisWorkbook
      For Each wksSheet In .Worksheets
        If InStr(wksSheet.Name, strPart) <> 0 Then
          wksSheet.Delete
        End If
      Next wksSheet
    End With
    Application.DisplayAlerts = True
  End If
End Sub

'***************************************************************************
'Procedure: DeleteThisSheet
'Purpose: Deletes specified worksheet
'Comments:  Turns off application display alerts temporarily
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/12/2019  Chris Kinion      Created
'***************************************************************************
Sub DeleteThisSheet(thisSheet As Worksheet)
  If thisSheet Is Nothing Then
  Else
    Application.DisplayAlerts = False
    thisSheet.Delete
    Application.DisplayAlerts = True
  End If
End Sub

'***************************************************************************
'Procedure: wksExists¤
'Purpose: Check if sheet exists by sheet name
'Comments: Returns True if sheet is found, False if not found
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/01  Chris Kinion      Written
'***************************************************************************
Function wksExists(sheetName As String) As Boolean
  Dim thisSheet As Worksheet
  On Error Resume Next
    Set thisSheet = ThisWorkbook.Sheets(sheetName)
  On Error GoTo 0

  If thisSheet Is Nothing Then
     wksExists = False
     Exit Function
  Else
    wksExists = True
  End If
End Function

'***************************************************************************
'Procedure: AddBlankSheetByName
'Purpose: Add blank sheet or clear existing sheet by name
'Calls: wksExists
'Comments: gwksSheet refers to a potential global worksheet variable
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/01  Chris Kinion      Written
'***************************************************************************
Sub AddBlankSheetByName(strSheetName As String)
  If wksExists(strSheetName) Then
    ThisWorkbook.Worksheets(strSheetName).Activate
    Cells.Delete
  Else
    'Worksheets.Add(after:=gwksSheet).Name = strSheetName
    Worksheets.Add(Count:=1).Name = strSheetName
    Dim wksNewSheet As Worksheet
    Set wksNewSheet = ThisWorkbook.Sheets(strSheetName)
    wksNewSheet.Activate
  End If
End Sub

'***************************************************************************
'Procedure: FreezeFirstRow(ByVal sheetName As String)
'Purpose: Freeze Pane of first row of given sheet in this workbook
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/14/2019  Chris Kinion      Created
'***************************************************************************
Sub FreezeFirstRow(ByVal sheetName As String)
  With ThisWorkbook
    .Activate
    .Worksheets(sheetName).Select
  End With
  With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
    .FreezePanes = True
  End With
End Sub

'***************************************************************************
'Procedure: FreezeFirstColumn(ByVal sheetName As String)
'Purpose: Freeze Pane of first column of given sheet in this workbook
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/14/2019  Chris Kinion      Created
'***************************************************************************
Sub FreezeFirstColumn(ByVal sheetName As String)
  With ThisWorkbook
    .Activate
    .Worksheets(sheetName).Select
  End With
  With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 0
    .FreezePanes = True
  End With
End Sub

'***************************************************************************
'Procedure: FreezeFirstRowAndColumn(ByVal sheetName As String)
'Purpose: Freeze first row and column of given sheet in this workbook
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/21/2019  Chris Kinion      Created
'***************************************************************************
Sub FreezeFirstRowAndColumn(ByVal sheetName As String)
  With ThisWorkbook
    .Activate
    .Worksheets(sheetName).Select
  End With
  With ActiveWindow
    .FreezePanes = False
    .SplitColumn = 1
    .SplitRow = 1
    .FreezePanes = True
  End With
End Sub

'***************************************************************************
'Procedure: ShowAllSheets
'Purpose: Show all hidden sheets
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/21/2019  Chris Kinion      Created
'***************************************************************************
Sub ShowAllSheets()
  Dim wks As Worksheet
  For Each wks In ThisWorkbook.Worksheets
    wks.Visible = xlSheetVisible
  Next wks
End Sub
