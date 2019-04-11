Attribute VB_Name = "Math_Estimation"
'***************************************************************************
'Module:      Math_Estimation
'Procedures:  KitsFromVolume: Step function that inputs an amount needed and
            '   ratio (amount per kit) to return a whole number of kits required
            ' testKitsFromVolume: prints to Immediate Window results of a test
            '   to KitsFromVolume procedure
'Comments:
'Changes----------------------------------------------
' Date        Programmer        Change
' 02/17/2019  Chris Kinion      Created
'***************************************************************************
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: KitsFromVolume
'Purpose: Return minimum whole number of kits required
'Comments: Unless even, always rounds up
'Changes----------------------------------------------
' Date        Programmer        Change
' 01/10/2019  Chris Kinion      Created
'***************************************************************************
Function KitsFromVolume(ByVal volumeIn As Double, ByVal volumePerKit As Double) As Long
  Dim lngTempKits As Long
  Dim dblKits As Double
  Dim tempDecimal As Double
  Dim roundUp As Boolean
  
  dblKits = volumeIn / volumePerKit
  lngTempKits = CLng(dblKits)
  If dblKits - lngTempKits > 0 Then ' add a kit
    KitsFromVolume = lngTempKits + 1
  ElseIf dblKits - lngTempKits <= 0 Then ' correct number of kits
    KitsFromVolume = lngTempKits
  End If
End Function

'***************************************************************************
'Procedure: KitsFromVolume
'Purpose: Test procedure KitsFromVolume
'         Explicitly return minimum whole number of kits required to Immediate Window
'***************************************************************************
Sub testKitsFromVolume()
  Dim dblVolume As Double
  Dim KitRatio As Double
  KitRatio = 20
  dblVolume = 40.1
  Debug.Print KitsFromVolume(dblVolume, KitRatio)
End Sub

