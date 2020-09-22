Attribute VB_Name = "MRoutineBuilder"
Option Explicit

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Sub MakeOptionBold(ctlOption As Variant)
'=================================================================
' Routine Name: MakeOptionBold
' Description: makes the selected option button text bold and the
'              rest not bold. Pass it a control array (currently
'              this must be an array of option buttons ONLY). Returns nothing.
'              ex: MakeOptionBold MyOptionArray()
' Author: Kurt Tischer
' Date: 10/25/2002 8:48:15 PM
' Copyright © 2002 3rd Ear Productions
' Notes:
'Author:    Kurt J. Tischer
'Date:      05-13-96
'Modification History:
'=================================================================

On Error Resume Next

Dim ctlSingleOption As Object

    For Each ctlSingleOption In ctlOption
        With ctlSingleOption
            If .Value = True Then
                .FontBold = True
            Else
                .FontBold = False
            End If
        End With
    Next
    
    Set ctlSingleOption = Nothing
    
End Sub


Public Sub SelectAll(Optional Ctrl As Variant)
' Routine Name: SelectAll
' Description: Selects all text in a text box
' Author: Kurt Tischer
' Date: 10/25/2002 8:48:15 PM
' Copyright © 2002 3rd Ear Productions
' Notes:
' Modification History:
'=================================================================
' 06/21/96  KT  Modified from the original to not pass # of elements

' 09/12/96  KT  Modified to use for...each to prevent consecutive index
'               errors and to optimize for speed
'
On Error Resume Next

Dim sTemp As String
Dim objCtrl As Control

If IsMissing(Ctrl) Then _
    Set objCtrl = Screen.ActiveForm.ActiveControl

'if the control is a masked edit box then select all
'the text in the control including the mask.
sTemp = objCtrl.Text

With objCtrl
    .SelStart = 0
    .SelLength = Len(sTemp)
End With

Set objCtrl = Nothing

    
End Sub

'====================================================================
'this sub should be executed from the Immediate window
'in order to get this app added to the VBADDIN.INI file
'you must change the name in the 2nd argument to reflecty
'the correct name of your project
'====================================================================
Sub AddToINI()
    Dim ErrCode As Long
    ErrCode = WritePrivateProfileString("Add-Ins32", "vbRoutineBuilder.Connect", "0", "vbaddin.ini")
End Sub

