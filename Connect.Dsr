VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6975
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   9135
   _ExtentX        =   16113
   _ExtentY        =   12303
   _Version        =   393216
   Description     =   "Sets up and builds basic sub and fuction parameters"
   DisplayName     =   "Routine Builder"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmRoutineBuilder        As New mfrmRoutineBuilder
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1


Private Function AddLogging(strErrLogging As String) As String
'=================================================================
' Routine Name: AddLogging
' Description: Writes logging info
' Author: Kurt Tischer
' Copyright © 2002 3rd Ear Productions
' Notes:
'Author:    Kurt J. Tischer
'Date:      05-13-96
'Modification History:
'=================================================================

On Error GoTo AddLoggingErr

With mfrmRoutineBuilder
    'check for log file
    If .optLog(0).Value = True Then
        strErrLogging = strErrLogging & vbCrLf
        strErrLogging = strErrLogging & vbTab & "'add error to log file" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "Dim iFreeFile As Integer" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "iFreeFile = FreeFile" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "If Right$(App.Path, 1) = " & Chr$(34) & "\" & Chr$(34) & " Then" & vbCrLf
        strErrLogging = strErrLogging & vbTab & vbTab & "Open App.Path & App.Title & " & Chr$(34) & ".LOG" & Chr$(34)
        strErrLogging = strErrLogging & " For Append As #iFreeFile" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "Else" & vbCrLf
        strErrLogging = strErrLogging & vbTab & vbTab & "Open App.Path & " & Chr$(34) & "\" & Chr$(34) & " & App.Title"
        strErrLogging = strErrLogging & " For Append As #iFreeFile" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "End If" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "Print #iFreeFile, App.Title, Err.Number, Err.Description, Err.Source, Now" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "Close #iFreeFile" & vbCrLf & vbCrLf
    ElseIf .optLog(1).Value = True Then
        strErrLogging = strErrLogging & vbTab & "'add error to log file"
        strErrLogging = strErrLogging & vbTab & "Dim iFreeFile As Integer" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "iFreeFile = FreeFile" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "Open " & Chr$(34) & .txtLogFile.Text & Chr$(34)
        strErrLogging = strErrLogging & " For Append As #iFreeFile" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "Print #iFreeFile, App.Title, Err.Number, Err.Description, Err.Source, Now" & vbCrLf
        strErrLogging = strErrLogging & vbTab & "Close #iFreeFile" & vbCrLf & vbCrLf
    End If
End With

    AddLogging = strErrLogging
    
AddLoggingExit:
    On Error Resume Next
    
    Exit Function
    
AddLoggingErr:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo AddLoggingExit
    
End Function



Private Function BuildDisplayMessage(ByVal strDisplay As String)
    On Error Resume Next
    strDisplay = strDisplay & vbTab & "Dim sMsg As String" & vbCrLf
    strDisplay = strDisplay & vbTab & "'Add your error display message here, or use this simple MsgBox display..." & vbCrLf & vbCrLf
    strDisplay = strDisplay & vbTab & "With Err" & vbCrLf
    strDisplay = strDisplay & vbTab & vbTab & "sMsg = ""Error: "" & .Number & vbCrLf" & vbCrLf
    strDisplay = strDisplay & vbTab & vbTab & "sMsg = sMsg & ""Description: "" & .Description & vbCrLf" & vbCrLf
    strDisplay = strDisplay & vbTab & vbTab & "sMsg = sMsg & ""Code Location: "" & App.Title & "":: ObjectName_ProcName""  & vbCrLf" & vbCrLf
    strDisplay = strDisplay & vbTab & vbTab & "sMsg = sMsg & ""Source: "" & .Source & vbCrLf" & vbCrLf
    strDisplay = strDisplay & vbTab & vbTab & "sMsg = sMsg & ""Last DLL Error: "" & .LastDllError & vbCrLf" & vbCrLf
    strDisplay = strDisplay & vbTab & vbTab & "MsgBox sMsg, vbOKOnly + vbCritical, ""Error"" & vbcrlf" & vbCrLf
    strDisplay = strDisplay & vbTab & "End With" & vbCrLf
    
    BuildDisplayMessage = strDisplay
        
End Function

Private Function BuildRaise(ByVal strRaiser As String)
    On Error Resume Next
    strRaiser = strRaiser & vbTab & "Err.Raise Err.Number, Err.Source, Err.Description " & vbCrLf & vbCrLf
    BuildRaise = strRaiser
End Function

Public Sub BuildRoutine()
'=================================================================
' Routine Name: BuildRoutine
' Description: This is the part that actually writes stuff to the
'               active window pane
' Author: Kurt Tischer
' Copyright © 2002 3rd Ear Productions
' Notes:
'Author:    Kurt J. Tischer
'Date:      05-13-96
'Modification History:
'=================================================================
    Dim strRoutine As String
    Dim li As ListItem
    Dim iCtr As Integer
    Dim sErrMsg As String
    Dim sRoutineName As String
    Dim objMember As Member
    Dim objCodePane As CodePane
    Dim objCodeModule As CodeModule
    Dim prj As VBProject
    Dim cmp As VBComponent
    
    Dim lResult As VbMsgBoxResult
    
    On Error GoTo BuildRoutineErr
        
    iCtr = 0
    
    With mfrmRoutineBuilder
            
        'begin comment block
        strRoutine = strRoutine & "'============================================================" & vbCrLf & vbCrLf
        strRoutine = strRoutine & "' Routine Name: " & .txtRoutineName.Text & vbCrLf
        strRoutine = strRoutine & "' Description: " & .txtDescription.Text & vbCrLf
        strRoutine = strRoutine & "' Author: " & .txtAuthor.Text & vbCrLf
        strRoutine = strRoutine & "' Date: " & Now & vbCrLf
        strRoutine = strRoutine & "' " & .txtCopyright.Text & vbCrLf
        strRoutine = strRoutine & "' Notes: " & .txtNotes.Text & vbCrLf & vbCrLf
        strRoutine = strRoutine & "' Modification History: " & vbCrLf & vbCrLf
        strRoutine = strRoutine & "'============================================================" & vbCrLf & vbCrLf
        strRoutine = strRoutine & vbCrLf  'end comment point
         
        Select Case True  'Private/Public
            Case .optScope(0).Value
                strRoutine = strRoutine & "Private "
            Case .optScope(1).Value
                strRoutine = strRoutine & "Public "
            Case Else
            
        End Select
        
        Select Case True  'Sub/Function
            Case .optRoutineType(0).Value
                strRoutine = strRoutine & "Sub "
            Case .optRoutineType(1).Value
                strRoutine = strRoutine & "Function "
            Case Else
            
        End Select
        
        'open paren
        strRoutine = strRoutine & .txtRoutineName.Text & "("
        sRoutineName = .txtRoutineName.Text
        
        'add arguments as datatype
        For Each li In .ListView1.ListItems
            If li.SubItems(2) = "Optional" Then
                strRoutine = strRoutine & "Optional " & li.Text & " As " & li.SubItems(1)
                If li.SubItems(3) <> "" Then strRoutine = strRoutine & " = " & .txtDefaultValue.Text
                strRoutine = strRoutine & ", "
            Else
                strRoutine = strRoutine & li.Text & " As " & li.SubItems(1) & ", "
            End If
        Next
        If Right$(strRoutine, 2) = ", " Then _
            strRoutine = Left$(strRoutine, Len(strRoutine) - 2)
            
        strRoutine = strRoutine & ")"  'close paren
        
        'add return type if function
        If .optRoutineType(1).Value = True And Not .cboReturnType.Text = "" Then
                strRoutine = strRoutine & "As " & .cboReturnType.Text
        ElseIf .optRoutineType(1).Value = True Then
            If vbNo = MsgBox("Functions typically return values.  Do you want to leave this blank?", vbYesNo, "Verify Function") Then
                strRoutine = strRoutine & "As Variant"
            Else
                strRoutine = strRoutine & "As Object"
            End If
        End If
        strRoutine = strRoutine & vbCrLf  'end entry point
        
        'add error handling type
        Select Case True
            Case .optErr(0).Value
                strRoutine = strRoutine & vbTab & "On Error Resume Next"
            Case .optErr(1).Value
                strRoutine = strRoutine & vbTab & "On Error GoTo Handler"
            Case .optErr(2).Value
                strRoutine = strRoutine & vbTab & "On Error Goto " & .txtRoutineName.Text & "Err"
            Case Else
            
        End Select
        
        strRoutine = strRoutine & vbCrLf & vbCrLf
        strRoutine = strRoutine & vbTab & "'Begin Routine: The bulk of your routine  here..." & vbCrLf & vbCrLf
        strRoutine = strRoutine & vbTab & .txtRoutineText.Text & vbCrLf & vbCrLf
        strRoutine = strRoutine & vbTab & "'End Routine: The bulk of your routine here..." & vbCrLf & vbCrLf
        
        If .optRoutineType(1).Value = True Then _
            strRoutine = strRoutine & vbTab & "'Set Function Return Data/Value..." & vbCrLf & vbCrLf
            
        'add exit from sub/function
        Select Case True
            Case .optErr(1).Value, .optErr(2).Value
                strRoutine = strRoutine & .txtRoutineName.Text & "Exit:" & vbCrLf
                strRoutine = strRoutine & vbTab & "On Error Resume Next" & vbCrLf & vbCrLf
                If .optRoutineType(1).Value = True Then
                    strRoutine = strRoutine & vbTab & "Exit Function"
                Else
                    strRoutine = strRoutine & vbTab & "Exit Sub"
                End If
            Case Else
            
        End Select
        
        strRoutine = strRoutine & vbCrLf & vbCrLf
        
        'add error handler block
        Select Case True
            Case .optErr(1).Value
                strRoutine = strRoutine & "Handler:" & vbCrLf
                'Check for display/raise/log
                If .optErrMode(0).Value = True Then
                    strRoutine = BuildDisplayMessage(strRoutine)
                ElseIf .optErrMode(1).Value = True Then
                    strRoutine = BuildRaise(strRoutine)
                End If
                
                If .chkAddLogging.Value = vbChecked Then strRoutine = AddLogging(strRoutine)

                'add Goto Exit
                strRoutine = strRoutine & vbTab & "GoTo " & .txtRoutineName.Text & "Exit" & vbCrLf & vbCrLf
            Case .optErr(2).Value
                strRoutine = strRoutine & .txtRoutineName.Text & "Err:" & vbCrLf & vbCrLf
                'Check for display/raise/log
                If .optErrMode(0).Value = True Then
                    strRoutine = BuildDisplayMessage(strRoutine)
                ElseIf .optErrMode(1).Value = True Then
                    strRoutine = BuildRaise(strRoutine)
                End If
                
                If .chkAddLogging.Value = vbChecked Then strRoutine = AddLogging(strRoutine)
                'add Goto Exit
                strRoutine = strRoutine & "GoTo " & .txtRoutineName.Text & "Exit" & vbCrLf & vbCrLf
            Case Else
            
        End Select
                
        'end Sub/Function
        If .optRoutineType(0).Value = True Then
            strRoutine = strRoutine & "End Sub"
        ElseIf .optRoutineType(1).Value = True Then
            strRoutine = strRoutine & "End Function" & vbCrLf
        End If
    End With
        
    'check for available code panes
    If VBInstance.CodePanes.Count = 0 Then
        sErrMsg = "No Code Panes. You must be in an active code pane;" & vbCrLf
        sErrMsg = sErrMsg & "preferrably the 'General Declarations' section" & vbCrLf
        sErrMsg = sErrMsg & "of a Form, Module, Class Module, or other object."
        MsgBox sErrMsg, vbExclamation + vbOKOnly, "Routine Builder Error"
    Else
        'check for existing routine or member name
        'go through all the projects
        For Each prj In VBInstance.VBProjects
            'and all it's components
            For Each cmp In prj.VBComponents
                'and all the members of each code pane
                For Each objMember In cmp.CodeModule.Members
                    If objMember.Name = sRoutineName Then
                        'begin building message when same name found
                        'give us the routine name and the window caption of the
                        sErrMsg = "[" & sRoutineName & "] Already Exists " & vbCrLf
                        sErrMsg = sErrMsg & "in Project [" & prj.Name & "]" & vbCrLf
                        sErrMsg = sErrMsg & "in Module [" & cmp.Name & "]" & vbCrLf
                        sErrMsg = sErrMsg & "as: " & vbCrLf & vbCrLf
                        
                        Select Case objMember.Scope
                            Case vbext_Private
                                sErrMsg = sErrMsg & "Private "
                            Case vbext_Public
                                sErrMsg = sErrMsg & "Public "
                            Case vbext_Friend
                                sErrMsg = sErrMsg & "Friend "
                        End Select
                        Select Case objMember.Type
                            Case vbext_mt_Method
                                sErrMsg = sErrMsg & "Method (Method, Procedure, Function) "
                            Case vbext_mt_Event
                                sErrMsg = sErrMsg & "Event "
                            Case vbext_mt_Property
                                sErrMsg = sErrMsg & "Property "
                            Case vbext_mt_Const
                                sErrMsg = sErrMsg & "Constant "
                            Case vbext_mt_Variable
                                sErrMsg = sErrMsg & "Variable "
                        End Select
                        
                        sErrMsg = sErrMsg & vbCrLf & vbCrLf & "Continue?" & vbCrLf & vbCrLf
                        sErrMsg = sErrMsg & "YES to continue; NO to abort."
                        lResult = MsgBox(sErrMsg, vbExclamation + vbYesNo, "Routine Builder Error")
                        If lResult = vbNo Then GoTo BuildRoutineExit
                    End If
                Next
            Next
        Next
        VBInstance.ActiveCodePane.CodeModule.AddFromString (strRoutine)
    End If
    
BuildRoutineExit:
    On Error Resume Next
    Set objMember = Nothing
    Set objCodePane = Nothing
    Exit Sub
    
BuildRoutineErr:
    sErrMsg = "Error: " & Err.Number & vbCrLf
    sErrMsg = sErrMsg & Err.Description & vbCrLf & vbCrLf
    sErrMsg = sErrMsg & "In frmBuildRoutine::BuildRoutine"
    
    MsgBox sErrMsg, vbOKOnly + vbExclamation, App.Title
    GoTo BuildRoutineExit
End Sub

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmRoutineBuilder.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mfrmRoutineBuilder Is Nothing Then
        Set mfrmRoutineBuilder = New mfrmRoutineBuilder
    End If
    
    Set mfrmRoutineBuilder.VBInstance = VBInstance
    Set mfrmRoutineBuilder.Connect = Me
    FormDisplayed = True
    mfrmRoutineBuilder.Show
   
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("Routine Builder")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmRoutineBuilder
    Set mfrmRoutineBuilder = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

