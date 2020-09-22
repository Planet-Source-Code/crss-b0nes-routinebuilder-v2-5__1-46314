VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mfrmRoutineBuilder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vbRoutineBuilder v.2.5"
   ClientHeight    =   5820
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRoutineBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4965
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   8758
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Routine "
      TabPicture(0)   =   "frmRoutineBuilder.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraRoutine(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Arguments"
      TabPicture(1)   =   "frmRoutineBuilder.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRoutine(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Error Handling "
      TabPicture(2)   =   "frmRoutineBuilder.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraRoutine(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Routine Text"
      TabPicture(3)   =   "frmRoutineBuilder.frx":019E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtRoutineText"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtRoutineText 
         Appearance      =   0  'Flat
         Height          =   4095
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   47
         Top             =   600
         Width           =   5415
      End
      Begin VB.Frame fraRoutine 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Index           =   2
         Left            =   -74880
         TabIndex        =   39
         Top             =   480
         Width           =   5565
         Begin VB.OptionButton optErr 
            Caption         =   "On Error GoTo RoutineNameErr"
            Height          =   285
            Index           =   2
            Left            =   150
            TabIndex        =   20
            Top             =   960
            Width           =   5205
         End
         Begin VB.OptionButton optErr 
            Caption         =   "On Error GoTo Handler"
            Height          =   345
            Index           =   1
            Left            =   150
            TabIndex        =   19
            Top             =   570
            Width           =   5205
         End
         Begin VB.OptionButton optErr 
            Caption         =   "On Error Resume Next"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   18
            Top             =   240
            Width           =   5175
         End
         Begin VB.Frame Frame1 
            Caption         =   "Other Handling Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2685
            Left            =   120
            TabIndex        =   41
            Top             =   1500
            Width           =   5265
            Begin VB.CheckBox chkAddLogging 
               Caption         =   "Log Error"
               Height          =   315
               Left            =   3180
               TabIndex        =   46
               Top             =   270
               Width           =   1155
            End
            Begin VB.Frame fraLog 
               Caption         =   "Log File Options"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1785
               Left            =   120
               TabIndex        =   42
               Top             =   750
               Width           =   5025
               Begin VB.OptionButton optLog 
                  Caption         =   "User defined path\filename:"
                  Height          =   405
                  Index           =   1
                  Left            =   180
                  TabIndex        =   24
                  Top             =   660
                  Width           =   4605
               End
               Begin VB.OptionButton optLog 
                  Caption         =   "Use App.Path\App.Title (w *.LOG extension)"
                  Height          =   405
                  Index           =   0
                  Left            =   180
                  TabIndex        =   23
                  Top             =   270
                  Width           =   4605
               End
               Begin VB.TextBox txtLogFile 
                  Height          =   315
                  Left            =   150
                  MaxLength       =   255
                  TabIndex        =   25
                  Top             =   1320
                  Width           =   4695
               End
               Begin VB.Label Label3 
                  Caption         =   "Log File Path\Name:"
                  Height          =   225
                  Left            =   210
                  TabIndex        =   43
                  Top             =   1080
                  Width           =   1545
               End
            End
            Begin VB.OptionButton optErrMode 
               Caption         =   "Raise Error"
               Height          =   285
               Index           =   1
               Left            =   1710
               TabIndex        =   22
               Top             =   270
               Width           =   1395
            End
            Begin VB.OptionButton optErrMode 
               Caption         =   "Display Error"
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   270
               Width           =   1485
            End
         End
      End
      Begin VB.Frame fraRoutine 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Index           =   1
         Left            =   -74760
         TabIndex        =   36
         Top             =   480
         Width           =   5565
         Begin ComctlLib.ListView ListView1 
            Height          =   2145
            Left            =   120
            TabIndex        =   15
            Top             =   2070
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   3784
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Argument"
               Object.Width           =   1270
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Data Type"
               Object.Width           =   1270
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Optional"
               Object.Width           =   1270
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Default Value"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.TextBox txtDefaultValue 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1110
            Width           =   1755
         End
         Begin VB.CheckBox chkOptional 
            Caption         =   "Optional"
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   390
            Left            =   4410
            TabIndex        =   17
            Top             =   3120
            Width           =   990
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   390
            Left            =   4410
            TabIndex        =   16
            Top             =   2580
            Width           =   990
         End
         Begin VB.ComboBox cboDataType 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   14
            Top             =   1560
            Width           =   1755
         End
         Begin VB.TextBox txtArgument 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   11
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label2 
            Caption         =   "Default Value:"
            Height          =   225
            Index           =   2
            Left            =   180
            TabIndex        =   40
            Top             =   1170
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Data Type:"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   38
            Top             =   1620
            Width           =   1185
         End
         Begin VB.Label Label2 
            Caption         =   "Argument Name:"
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   37
            Top             =   270
            Width           =   1665
         End
      End
      Begin VB.Frame fraRoutine 
         Height          =   4335
         Index           =   0
         Left            =   180
         TabIndex        =   28
         Top             =   360
         Width           =   5565
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   5265
            Begin VB.OptionButton optScope 
               Caption         =   "Public"
               Height          =   225
               Index           =   1
               Left            =   3240
               TabIndex        =   2
               Top             =   210
               Width           =   1785
            End
            Begin VB.OptionButton optScope 
               Caption         =   "Private"
               Height          =   225
               Index           =   0
               Left            =   1140
               TabIndex        =   1
               Top             =   210
               Width           =   1785
            End
            Begin VB.Label Label1 
               Caption         =   "Scope:"
               Height          =   225
               Index           =   7
               Left            =   120
               TabIndex        =   45
               Top             =   210
               Width           =   1065
            End
         End
         Begin VB.TextBox txtDescription 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            MaxLength       =   80
            TabIndex        =   9
            Top             =   3060
            Width           =   3855
         End
         Begin VB.TextBox txtNotes 
            Appearance      =   0  'Flat
            Height          =   585
            Left            =   1500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   3540
            Width           =   3855
         End
         Begin VB.TextBox txtCopyright 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            MaxLength       =   80
            TabIndex        =   8
            Text            =   "Copyright © "
            Top             =   2580
            Width           =   3855
         End
         Begin VB.TextBox txtAuthor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "Author's Name"
            Top             =   2100
            Width           =   3855
         End
         Begin VB.OptionButton optRoutineType 
            Caption         =   "Function"
            Height          =   225
            Index           =   1
            Left            =   2520
            TabIndex        =   5
            Top             =   1260
            Width           =   1305
         End
         Begin VB.OptionButton optRoutineType 
            Caption         =   "Sub"
            Height          =   225
            Index           =   0
            Left            =   1500
            TabIndex        =   4
            Top             =   1260
            Width           =   765
         End
         Begin VB.ComboBox cboReturnType 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            TabIndex        =   6
            Top             =   1620
            Width           =   1935
         End
         Begin VB.TextBox txtRoutineName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            MaxLength       =   255
            TabIndex        =   3
            Text            =   "NewRoutine"
            Top             =   780
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Description:"
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   35
            Top             =   3090
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Notes:"
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   34
            Top             =   3540
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Copyright Info:"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   2610
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Author:"
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   2130
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Return Type:"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   1650
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Routine Type:"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   1230
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Routine Name:"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   810
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   390
      Left            =   5040
      TabIndex        =   27
      Top             =   5280
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3930
      TabIndex        =   26
      Top             =   5280
      Width           =   990
   End
End
Attribute VB_Name = "mfrmRoutineBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect


Private Sub chkAddLogging_Click()
    On Error Resume Next
    
    chkAddLogging.FontBold = Not chkAddLogging.FontBold
    
    DisableLogControls
    
End Sub

Private Sub chkOptional_Click()
' Routine Name: chkOptional_Click
' Description: If selected, corresponding argument will be optional
' Author: Kurt Tischer
' Date: 10/25/2002 8:51:23 PM
' Copyright © 2002 3rd Ear Productions
' Notes:

    On Error Resume Next

    If chkOptional.Value = 0 Then
        chkOptional.Font.Bold = False
        txtDefaultValue.Enabled = False
        txtDefaultValue.BackColor = fraRoutine(0).BackColor
    ElseIf chkOptional.Value = 1 Then
        chkOptional.Font.Bold = True
        txtDefaultValue.Enabled = True
        txtDefaultValue.BackColor = txtRoutineName.BackColor
    End If



End Sub
Private Sub cboReturnType_KeyPress(KeyAscii As Integer)
' Routine Name: cboReturnType_KeyPress
' Description: Kills Space Key
' Author: Kurt Tischer
' Date: 10/25/2002 8:50:03 PM
' Copyright © 2002 3rd Ear Productions
' Notes:

    On Error Resume Next

    If KeyAscii = 32 Then KeyAscii = 0



End Sub
Private Sub cboDataType_KeyPress(KeyAscii As Integer)
' Routine Name: cbo_DataType_KeyPress
' Description: Kills Space Key
' Author: Kurt Tischer
' Date: 10/25/2002 8:48:15 PM
' Copyright © 2002 3rd Ear Productions
' Notes:

    On Error Resume Next

    If KeyAscii = 32 Then KeyAscii = 0

End Sub
Private Function DisableErrModeControls(Index) As Boolean
' Routine Name: DisableErrModeControls
' Description: Enables/disables error mode controls based]
'               based on index value passed here
' Author: Kurt Tischer
' Date: 10/25/2002 8:48:15 PM
' Copyright © 2002 3rd Ear Productions
' Notes:

    Dim iCtr As Integer
    
    On Error Resume Next
    
    If Index = 0 Then
        For iCtr = 0 To optErrMode.UBound
            optErrMode(iCtr).Enabled = False
        Next
        
        Frame1.Enabled = False
        chkAddLogging.Enabled = True
    Else
        For iCtr = 0 To optErrMode.UBound
            optErrMode(iCtr).Enabled = True
        Next
        
        Frame1.Enabled = True
        chkAddLogging.Enabled = True
    End If
    
End Function


Private Sub DisableLogControls()
' Routine Name: DisableLogControls
' Description: Enables/disables log controls based
'               based on index value passed here
' Author: Kurt Tischer
' Date: 10/25/2002 8:48:15 PM
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    If chkAddLogging.Value = vbUnchecked Then
        fraLog.Enabled = False
        optLog(0).Enabled = False
        optLog(1).Enabled = False
    Else
        fraLog.Enabled = True
        optLog(0).Enabled = True
        optLog(1).Enabled = True
    End If
    
End Sub









Private Sub cmdAdd_Click()
' Routine Name: cmdAdd_Click
' Description: Adds argument parameters to list box
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:

    Dim li As ListItem
    Dim sArgName As String
    Dim sErrMsg As String
    
    On Error GoTo cmdAddClickErr
    
    sArgName = txtArgument.Text
    
    With ListView1
        If chkOptional.Value = 0 Then
            Set li = .ListItems.Add(1, sArgName, sArgName)
            li.SubItems(1) = cboDataType.Text
            li.SubItems(2) = "Required"
        ElseIf chkOptional.Value = 1 Then
            Set li = .ListItems.Add(, sArgName, sArgName)
            li.SubItems(1) = cboDataType.Text
            li.SubItems(2) = "Optional"
            li.SubItems(3) = txtDefaultValue.Text
        End If
    End With
    
    
cmdAddClickExit:
    On Error Resume Next
    Set li = Nothing
    cmdRemove.Enabled = IIf(ListView1.ListItems.Count > 0, True, False)
    Exit Sub
    
cmdAddClickErr:
    sErrMsg = "Error: " & Err.Number & vbCrLf
    sErrMsg = sErrMsg & Err.Description & vbCrLf & vbCrLf
    sErrMsg = sErrMsg & "In frmBuildRoutine::BuildRoutine"
    
    MsgBox sErrMsg, vbOKOnly + vbExclamation, App.Title
    GoTo cmdAddClickExit
    
End Sub
Private Sub cmdCancel_Click()
' Routine Name: cmdCancel_Click
' Description: Hide this puppy
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:

    On Error Resume Next
    
    Connect.Hide

End Sub

Private Sub cmdOK_Click()
' Routine Name: cmdOK_Click
' Description: Takin' care of business
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:

    On Error Resume Next
    
    Connect.BuildRoutine
            
    Connect.Hide
    
    Connect.Show
End Sub


Private Sub cmdRemove_Click()
' Routine Name: cmdRemove_Click
' Description: Removes selected argument parameters from list box
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:

    On Error Resume Next
       
    With ListView1
        .ListItems.Remove (.SelectedItem.Index)
        .Refresh
        cmdRemove.Enabled = IIf(.ListItems.Count > 0, True, False)
        
    End With
    
End Sub

Private Sub Form_Load()
' Routine Name: Form_Load
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002 8:48:15 PM
' Copyright © 2002 3rd Ear Productions
' Notes:
    Dim i As Integer
    
    On Error GoTo frmLoadErr
    
    'fill combo with common data types
    cboReturnType.AddItem "Array"
    cboReturnType.AddItem "Boolean"
    cboReturnType.AddItem "Byte"
    cboReturnType.AddItem "Currency"
    cboReturnType.AddItem "Date"
    cboReturnType.AddItem "Double"
    cboReturnType.AddItem "Integer"
    cboReturnType.AddItem "Long"
    cboReturnType.AddItem "Object"
    cboReturnType.AddItem "Single"
    cboReturnType.AddItem "String"
    cboReturnType.AddItem "Variant"
    
    'and again for function return data types
    For i = 0 To cboReturnType.ListCount - 1
        cboDataType.AddItem cboReturnType.List(i)
    Next
    cboDataType.Enabled = False
    
    'set copyright year to now
    txtCopyright.Text = txtCopyright.Text & Year(Now) & " "
        
    'initialize form settings
    SSTab1.Tab = 0
    
    optScope(0).Value = True
    optRoutineType(0).Value = True
    optErr(0).Value = True
    optErrMode(0).Value = True
    optLog(0).Value = True
    fraLog.Enabled = False
    optLog(0).Enabled = False
    optLog(1).Enabled = False
    chkAddLogging.Enabled = False
    chkOptional.Font.Bold = False
    chkOptional.Enabled = False
    txtDefaultValue.Enabled = False
    txtDefaultValue.BackColor = fraRoutine(0).BackColor
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    
frmLoadExit:
    On Error Resume Next
    Exit Sub
    
frmLoadErr:
    Err.Raise Err.Number, App.Title & "::mfrmRoutineBuilder_Load::" & Err.Source, Err.Description
    GoTo frmLoadExit

End Sub

Private Sub optErr_Click(Index As Integer)
' Routine Name: optErr_Click
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:

    On Error Resume Next
    
    DisableErrModeControls Index
    
    MakeOptionBold optErr
    
End Sub

Private Sub optErrMode_Click(Index As Integer)
' Routine Name: optErrMode_Click
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    MakeOptionBold optErrMode
    
    'DisableLogControls
    
End Sub

Private Sub optLog_Click(Index As Integer)
' Routine Name: optLog_Click
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    MakeOptionBold optLog
    
    If Index = optLog.UBound Then
        txtLogFile.Enabled = True
        txtLogFile.BackColor = txtRoutineName.BackColor
    Else
        txtLogFile.Enabled = False
        txtLogFile.BackColor = fraLog.BackColor
    End If
    
End Sub


Private Sub optRoutineType_Click(Index As Integer)
' Routine Name: optRoutineType_Click
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    MakeOptionBold optRoutineType
    
    If Index = 0 Then
        cboReturnType.Enabled = False
    Else
        cboReturnType.Enabled = True
    End If
    
End Sub





Private Sub optScope_Click(Index As Integer)
' Routine Name: optScope_Click
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    MakeOptionBold optScope
    
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
' Routine Name: SSTab1_Click
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    Dim i As Integer
    
    On Error Resume Next
    
    For i = 0 To fraRoutine.UBound
        If i = SSTab1.Tab Then
            fraRoutine(i).Enabled = True
        Else
            fraRoutine(i).Enabled = False
        End If
    Next
    
    
End Sub

Private Sub txtArgument_Change()
' Routine Name: txtArgument_Change
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    Dim i As VbVarType
    
    On Error Resume Next
    
    With cboDataType
        Select Case txtArgument.Text
            Case "a"
                .ListIndex = 0
            Case "b"
                .ListIndex = 1
            Case "bt"
                .ListIndex = 2
            Case "c"
                .ListIndex = 3
            Case "dt"
                .ListIndex = 4
            Case "d"
                .ListIndex = 5
            Case "i"
                .ListIndex = 6
            Case "l"
                .ListIndex = 7
            Case "obj"
                .ListIndex = 8
            Case "sng", "sgl"
                .ListIndex = 9
            Case "s"
                .ListIndex = 10
            Case "v"
                .ListIndex = 11
            Case Else
            
        End Select
        .Text = .List(.ListIndex)
    End With
    If Len(txtArgument.Text) > 0 Then
        cmdAdd.Enabled = True
        chkOptional.Enabled = True
        cboDataType.Enabled = True
    Else
        cmdAdd.Enabled = False
        chkOptional.Value = vbUnchecked
        chkOptional.Enabled = False
        cboDataType.Enabled = False
        cboDataType.Text = ""
    End If
    
End Sub

Private Sub txtArgument_KeyPress(KeyAscii As Integer)
' Routine Name: txtArgument_KeyPress
' Description: Kill Space Key
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    If KeyAscii = 32 Then KeyAscii = 0
    
End Sub


Private Sub txtAuthor_GotFocus()
' Routine Name: txtAuthor_GotFocus
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    SelectAll
End Sub


Private Sub txtCopyright_GotFocus()
' Routine Name: txtCopyright_GotFocus
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    SelectAll
End Sub


Private Sub txtDescription_GotFocus()
' Routine Name: txtDescription_GotFocus
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    SelectAll
End Sub


Private Sub txtNotes_GotFocus()
' Routine Name: txtNotes_GotFocus
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    SelectAll
End Sub


Private Sub txtRoutineName_GotFocus()
' Routine Name: txtRoutineName_GotFocus
' Description:
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    SelectAll
End Sub


Private Sub txtRoutineName_KeyPress(KeyAscii As Integer)
' Routine Name: txtRoutineName_KeyPress
' Description: Kill Space Key
' Author: Kurt Tischer
' Date: 10/25/2002
' Copyright © 2002 3rd Ear Productions
' Notes:
    On Error Resume Next
    
    If KeyAscii = 32 Then KeyAscii = 0
    
End Sub


