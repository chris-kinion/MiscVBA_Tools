Attribute VB_Name = "Manipulate_Shapes"
'***************************************************************************
'Module: Shapes
'Procedures:  Function CheckObjectExists>: Returns TRUE if finds an object of that name
            ' clearShapesFromCurrentSheet: Deletes all shapes from ActiveSheet
            ' listAllShapesFromCurrentSheet: List names of all shapes from ActiveSheet in Immediate Window
            ' checkForThisShape: Prints False in Immediate Window if finds shape name in worksheet
            ' Function DoesShapeExist: Returns Boolean if finds shape
            ' adjustShape: Modify selected shapes width and height
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Amalgamated to module
'***************************************************************************
Option Explicit

'***************************************************************************
'Procedure: CheckObjectExists
'Purpose: Returns TRUE if finds an object of that name
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Function CheckObjectExists(argName As String) As Boolean
 Dim obj As Object
 CheckObjectExists = False
 For Each obj In ActiveSheet.Shapes
   If UCase(obj.Name) = UCase(argName) Then CheckObjectExists = True: Exit Function
 Next obj
End Function

'***************************************************************************
'Procedure: clearShapesFromCurrentSheet
'Purpose: Deletes all shapes from ActiveSheet
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Sub clearShapesFromCurrentSheet()
 For Each possibleShape In ActiveSheet.Shapes
   possibleShape.Delete
 Next possibleShape
End Sub

'***************************************************************************
'Procedure: listAllShapesFromCurrentSheet
'Purpose: List names of all shapes from ActiveSheet in Immediate Window
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Sub listAllShapesFromCurrentSheet()
 For Each possibleShape In ActiveSheet.Shapes
   Debug.Print possibleShape.Name
 Next possibleShape
End Sub

'***************************************************************************
'Procedure: checkForThisShape
'Purpose: Prints False in Immediate Window if finds shape name in worksheet
'Comments: Loops through all shapes in search
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Sub checkForThisShape(mySearchVar As String, wksSheet As Worksheet)
  Dim possibleShape As Variant
  Dim foundMyShape As Boolean
  foundMyShape = True
  
  With wksSheet
    For Each possibleShape In .Shapes
      If Not StrComp(possibleShape.Name, mySearchVar, vbTextCompare) <> 0 Then
        foundMyShape = False
      End If
      Debug.Print possibleShape.Name & " " & foundMyShape
    Next possibleShape
  End With
End Sub

'***************************************************************************
'Procedure: DoesShapeExist
'Purpose: Returns Boolean if finds shape
'Comments: Uses error method of searching for shape
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Function DoesShapeExist(thisShapeName As String, wks As Worksheet) As Boolean
  ' Return true if shape exists, false if does not exist
  Dim myShape As sHape

  On Error Resume Next
  Set myShape = wks.Shapes(thisShapeName)
  On Error GoTo 0

  If myShape Is Nothing Then
     ' MsgBox "Box 1 does not exist on " & ActiveSheet.Name
     DoesShapeExist = False
     Exit Function
  Else
    ' MsgBox "This shape does exist on " & ActiveSheet.Name
    DoesShapeExist = True
  End If
End Function

'***************************************************************************
'Procedure: adjustShape
'Purpose: Modify selected shapes width and height
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Sub adjustShape(newWidth As Long, newHeight As Long)
  Selection.ShapeRange.Width = newWidth
  Selection.ShapeRange.Height = newHeight
End Sub
