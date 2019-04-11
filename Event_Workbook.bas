Attribute VB_Name = "Event_Workbook"
'***************************************************************************
'Module: Event_Workbook
'Procedures:  IsWorkBookOpen
            ' OpenWorkBook
            ' CloseWorkBook
            ' LoadFileUsingFilter
            ' SaveThisSheet
            ' UnhideWorkbooks
'Changes----------------------------------------------
' Date        Programmer        Change
' 09/12/2018  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: IsWorkBookOpen
'Purpose: return boolean if file is open
'Arguments: testBook As Workbook
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/12  Perduco #130      Created
'***************************************************************************
Function IsWorkBookOpen(testBook As Workbook) As Boolean
  On Error Resume Next
  IsWorkBookOpen = (Not testBook Is Nothing)
  On Error GoTo 0
End Function

'***************************************************************************
'Procedure: OpenWorkBook
'Purpose: Opens and returns a workbook provided the complete path of the workbook
'Calls: UpdateStatusBar
'Arguments: strPath As String
'Comments: gblnFatalError is a global Boolean that tracks error status
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/10/04  Perduco #130      Created
'***************************************************************************
Function OpenWorkBook(strPath As String) As Workbook
  Dim strBookName As String
  Application.AutoRecover.Enabled = False
  strBookName = Mid(strPath, InStrRev(strPath, "\", -1, vbTextCompare) + 1)
  Call UpdateStatusBar("Opening " & strBookName)
  On Error Resume Next
  Set OpenWorkBook = Workbooks.Open(strPath)
  If Err.Number <> 0 Then
    MsgBox "Unable to open file " & strBookName
    'gblnFatalError = True
    Exit Function
  End If
  On Error GoTo 0
End Function

'***************************************************************************
'Procedure: CloseWorkBook
'Purpose: Closes a workbook without changes if it is open
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/10/04  Perduco #130      Created
'***************************************************************************
Sub CloseWorkbook(wkbkOpen As Workbook)
  On Error Resume Next
    If IsWorkBookOpen(wkbkOpen) Then ' Close First book
      wkbkOpen.Close savechanges:=False
    End If
  On Error GoTo 0
End Sub

'***************************************************************************
'Procedure: LoadFileUsingFilter
'Purpose: Opens only files with specific extensions
'Comments: This example only allows DIF and PDF files to be seen/selected
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/10/04  Perduco #130      Created
'***************************************************************************
Sub LoadFileUsingFilter()
  Dim intChoice As Integer
  Dim strPath As String
 
  With Application.FileDialog(msoFileDialogOpen)
  Call .Filters.Clear ' Remove any previous filters
  Call .Filters.Add("DIF Files", "*.dif") ' Filter
  Call .Filters.Add("PDF Files", "*.pdf") ' Filter
  .Title = "Pick a file" ' Title on dialog
  intChoice = .Show ' Dialog visible to user
  .AllowMultiSelect = False ' Allows only one selection
  If intChoice <> 0 Then ' Conditional for when not cancelled
    strPath = .SelectedItems(1)
    Debug.Print strPath
    Debug.Print "Choice of " & intChoice
  End If
  End With
End Sub

'***************************************************************************
'Procedure: SaveThisSheet
'Purpose: Saves a worksheet in a specific file format
'Comments:  This saves a worksheet as a .Dif, but can be changed to any supported format:
          ' https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/10/04  Perduco #130      Created
'***************************************************************************
Sub SaveThisSheet(thisSheet As Worksheet)
  If Not thisSheet Is Nothing Then
    Dim pathOne As String
    pathOne = ActiveWorkbook.Path & "\" & "Test " & Format(Now, "yyyy-mm-dd") & " " & thisSheet.Name
    thisSheet.Copy ' .Move will remove the sheet from the current workbook
    With ActiveWorkbook
      .SaveAs fileName:=pathOne, FileFormat:=xlDIF ' Path (including file name) and file type - see list of options
      .Close ' Otherwise new workbook stays open
    End With
  Else
  End If
End Sub

'***************************************************************************
'Procedure: UnhideWorkbooks
'Purpose: Make all hidden workbooks visible
'Changes----------------------------------------------
' Date        Programmer        Change
' 05/31/2018  Chris Kinion      Created
'***************************************************************************
Public Sub UnhideWorkbooks()
  Dim i As Long
  For i = 1 To Workbooks.Count
    If Application.Windows(i).Visible = False Then
      Application.Windows(i).Visible = True
    End If
  Next
End Sub

