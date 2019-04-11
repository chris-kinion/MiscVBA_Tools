Attribute VB_Name = "Event_MsgBox"
'***************************************************************************
'Module: Event_MsgBox
'Procedures: getFolderPath
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/08/2019  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: getFolderPath
'Purpose: Sets "startPath" as path to a selected folder, add "\" at end
'Comments: Uses File Dialog Folder Picker
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/08/2019  Chris Kinion      Created
'***************************************************************************
Sub getFolderPath(ByRef startPath As String, ByVal searchTitle As String)
  Dim intChoice As Integer
  Dim thisFileName As String
  With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = searchTitle
    .Show
    If .SelectedItems.Count = 0 Then
      startPath = ""
    Else
      startPath = .SelectedItems(1)
      If Right(startPath, 1) <> "\" Then startPath = startPath & "\"
    End If
  End With
End Sub
