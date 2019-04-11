Attribute VB_Name = "Manipulate_Types"
'***************************************************************************
'Module:
'Procedures: Function WhatIsTypeName: Return variable type in MsgBox
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: WhatIsTypeName
'Purpose: Return variable type in MsgBox
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 04/10/2019  Chris Kinion      Added to module
'***************************************************************************
Function WhatIsTypeName(myVar As Variant)
  MsgBox TypeName(myVar)
End Function
