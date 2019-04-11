Attribute VB_Name = "Sample_FileExport"
'***************************************************************************
'Module: Sample_FileExport
'Procedures:  wksFinalPSWtoPDF
            ' wksFinalPSWtoXLSX
            ' makeValidFileName
'Comments: Exports a worksheet to PDF or xls
            ' References UDF wksExists found in module Event_Worksheet
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/01  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: wksFinalPSWtoPDF
'Function: Exports sheet Final PSW to pdf file
'Called by: checkAndExportPSW (initSheets)
'Calls: wksExists, makeValidFileName
'Arguments: strFileName as String
'Comments: based on printSummaryToPDF() from module ExportAndPrint
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/01  Chris Kinion      Written
'***************************************************************************
Sub wksFinalPSWtoPDF(Optional strFileName As String = "summary")
  Dim strFilePath As String
  Dim strSavePath As String
  
  strFileName = makeValidFileName(strFileName)
  strSavePath = Application.ThisWorkbook.Path & "\Projects\pdfs\"
  strFilePath = Application.ThisWorkbook.Path & "\Projects\pdfs\" & strFileName & ".pdf"
    
  If Dir(strSavePath, vbDirectory) = "" Then
    MkDir Path:=strSavePath
  Else
    'Debug.Print "Save Path Existed"
  End If
  
  If wksExists(gstrFinalPSW) Then
    On Error Resume Next
      If Dir(strFilePath) <> vbNullString Then
        Kill strFilePath
      End If
      
      gwksFinalPSW.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=strFilePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
      If Err.Number = 0 Then
        MsgBox ("File " & strFileName & ".pdf saved to:" & vbLf & strSavePath)
      Else
        MsgBox ("Error saving """ & strFileName & ".pdf"". " & vbLf & Err.Description)
      End If
    On Error GoTo 0
  End If
  
End Sub

'***************************************************************************
'Procedure: wksFinalPSWtoXLSX
'Function: Exports sheet Final PSW to xls file
'Called by: checkAndExportPSW (initSheets)
'Calls: wksExists, makeValidFileName
'Arguments: strFileName as String
'Comments: Based on final section of print_click from module ExportAndPrint
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/01  Chris Kinion      Written
'***************************************************************************
Sub wksFinalPSWtoXLSX(Optional strFileName As String = "summary")
  Application.ScreenUpdating = False
  
  Dim strFilePath As String
  Dim strSavePath As String
  
  strFileName = makeValidFileName(strFileName)
  strSavePath = Application.ThisWorkbook.Path & "\Projects\ExcelFiles\"
  strFilePath = Application.ThisWorkbook.Path & "\Projects\ExcelFiles\" & strFileName & ".xlsx"
  
  If Dir(strSavePath, vbDirectory) = "" Then
    MkDir Path:=strSavePath
  Else
    'Debug.Print "Save Path Existed"
  End If
  
  If wksExists(gstrFinalPSW) Then
    On Error Resume Next
      If Dir(strFilePath) <> vbNullString Then
        SetAttr strFilePath, vbNormal
        Kill strFilePath
      End If
      
      gwksFinalPSW.Copy
      ActiveWorkbook.ActiveSheet.Unprotect
      ActiveWorkbook.ActiveSheet.Shapes("printButton").Delete
      ActiveWorkbook.SaveAs _
        fileName:=strFilePath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
      ActiveWindow.Close
      SetAttr strFilePath, vbReadOnly
        
      If Err.Number = 0 Then
        MsgBox ("File " & strFileName & ".xlsx saved to:" & vbLf & strSavePath)
      Else
        MsgBox ("Error saving """ & strFileName & ".xlsx"". " & vbLf & Err.Description)
      End If
    On Error GoTo 0
  End If
  Application.ScreenUpdating = True
End Sub

'***************************************************************************
'Procedure: makeValidFileName
'Purpose: returns a string without illegal characters
'Called by: wksFinalPSWtoPDF, wksFinalPSWtoXLSX
'Calls:
'Arguments:
'Comments: based on strLegalFileName from module ExportAndPrint
'Changes----------------------------------------------
' Date        Programmer        Change
' 2018/09/01  Chris Kinion      Written
'***************************************************************************
Function makeValidFileName(strFileNameIn As String) As String ' formatted by CK
  Dim i As Integer
  Const strIllegalCharacters As String = "\/|?*<>"":"
  makeValidFileName = strFileNameIn
  
  For i = 1 To Len(strIllegalCharacters)
    makeValidFileName = Replace(makeValidFileName, Mid(strIllegalCharacters, i, 1), "_")
  Next i
End Function

