Attribute VB_Name = "VBA_IO"
'***************************************************************************
'Module: VBA_IO
'Comments: Exports All VBA Modules. See instructions below
'          https://www.rondebruin.nl/win/s9/win002.htm
'Changes----------------------------------------------
' Date        Programmer        Change
' 03/06/2019  Chris Kinion      Created module, imported procedures
'***************************************************************************

'How does it work ?
'1: Open the file with the code from this page
'2: Open/Activate the file with the modules you want to export
'3: Run the ExportModules macro
'4: Note: If you look in the VBAProjectFiles folder you see the files now
'5: Open/Activate the workbook where you want to add the modules to
'6: Run the ImportModules macro (It delete all existing modules/userforms from this Workbook first)
'7: Done

'First, you need to set an reference to the VBA Extensibility library. The library contains the definitions of the objects that make up the VBProject.
'In the VBA editor, go the the Tools menu and choose References.
'In that dialog, scroll down to and check the entry for Microsoft Visual Basic For Applications Extensibility 5.3.
'In that dialog, scroll down to and check the entry for Microsoft Scripting Runtime.
'If you do not set this reference, you will receive a User-defined type not defined compiler error.

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub


Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.file
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function

