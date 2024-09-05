VERSION 5.00
Begin VB.Form frmUI 
   BackColor       =   &H00400000&
   Caption         =   "Drill & Tap"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   11880
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleMode       =   0  'User
   ScaleWidth      =   11880
   Begin VB.Frame frameToolCount 
      BackColor       =   &H00400000&
      Caption         =   "Tool Count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3975
      Left            =   8880
      TabIndex        =   16
      Top             =   2880
      Width           =   2655
      Begin VB.CommandButton cmdResetCount 
         Caption         =   "Reset Tap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   1440
         TabIndex        =   21
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdResetCount 
         Caption         =   "Reset Drill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblDrillTapCount 
         BackColor       =   &H00400000&
         Caption         =   "Cycles:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDrillTapCount 
         BackColor       =   &H00400000&
         Caption         =   "Drill :     0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblDrillTapCount 
         BackColor       =   &H00400000&
         Caption         =   "Tap :     0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   2175
      End
   End
   Begin VB.Frame frameDRO 
      BackColor       =   &H00400000&
      Caption         =   "DRO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2415
      Left            =   8880
      TabIndex        =   13
      Top             =   240
      Width           =   2655
      Begin VB.Label lblDRO 
         BackColor       =   &H00400000&
         Caption         =   "Y:     0.000 in"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblDRO 
         BackColor       =   &H00400000&
         Caption         =   "X:     0.000 in"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame frameSetup 
      BackColor       =   &H00400000&
      Caption         =   "  SETUP  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6615
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   8295
      Begin VB.CheckBox chkDryRun 
         BackColor       =   &H00400000&
         Caption         =   "Dry Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   360
         TabIndex        =   33
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Frame frmSystemStatus 
         BackColor       =   &H00400000&
         Caption         =   "System Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2775
         Left            =   3000
         TabIndex        =   31
         Top             =   360
         Width           =   5055
         Begin VB.Label lblSystemStatus 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            Caption         =   "System Booting"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   735
            Left            =   480
            TabIndex        =   32
            Top             =   600
            Width           =   4095
         End
      End
      Begin VB.Frame frmDrillTapDepth 
         BackColor       =   &H00400000&
         Caption         =   " Operation Limits "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   3135
         Left            =   4560
         TabIndex        =   24
         Top             =   3240
         Width           =   3495
         Begin VB.TextBox txtDeep 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   1200
            TabIndex        =   27
            Text            =   "80"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtDeep 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   1200
            TabIndex        =   25
            Text            =   "2.000"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblDeep 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   2640
            TabIndex        =   30
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label lblDeep 
            BackStyle       =   0  'Transparent
            Caption         =   "in"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   2640
            TabIndex        =   29
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label lblDeep 
            BackStyle       =   0  'Transparent
            Caption         =   "Tap to Depth:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblDeep 
            BackStyle       =   0  'Transparent
            Caption         =   "Drill to Depth:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdGoToDrillTap 
         Caption         =   "Go to Tap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   23
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CommandButton cmdGoToDrillTap 
         Caption         =   "Go to Drill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Timer tmrFSM 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   120
         Top             =   360
      End
      Begin VB.Frame frameSpeeds 
         BackColor       =   &H00400000&
         Caption         =   "  System Speeds  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2295
         Left            =   240
         TabIndex        =   6
         Top             =   4080
         Width           =   3975
         Begin VB.TextBox txtSpeeds 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   2280
            TabIndex        =   12
            Text            =   "3.000"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox txtSpeeds 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   2280
            TabIndex        =   11
            Text            =   "1.000"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtSpeeds 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   2280
            TabIndex        =   10
            Text            =   "0.250"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblSpeeds 
            BackStyle       =   0  'Transparent
            Caption         =   "Drill Speed:                 ips"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label lblSpeeds 
            BackStyle       =   0  'Transparent
            Caption         =   "Tap Speed:                 ips"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   1080
            Width           =   3735
         End
         Begin VB.Label lblSpeeds 
            BackStyle       =   0  'Transparent
            Caption         =   "Jog Speed:                 ips"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   1560
            Width           =   3735
         End
      End
      Begin VB.OptionButton optDrillTapAuto 
         BackColor       =   &H00400000&
         Caption         =   "Auto Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   2415
      End
      Begin VB.OptionButton optDrillTapAuto 
         BackColor       =   &H00400000&
         Caption         =   "Run One Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton optDrillTapAuto 
         BackColor       =   &H00400000&
         Caption         =   "Run Tap Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
      End
      Begin VB.OptionButton optDrillTapAuto 
         BackColor       =   &H00400000&
         Caption         =   "Run Drill Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   7080
      Width           =   11175
   End
   Begin VB.Menu topbarJoy 
      Caption         =   "Joystick"
   End
   Begin VB.Menu topbarSet0 
      Caption         =   "Home Machine"
   End
   Begin VB.Menu topbarSetDrill 
      Caption         =   "Set Drill Position"
   End
   Begin VB.Menu topbarSetTap 
      Caption         =   "Set Tap Position"
   End
   Begin VB.Menu topbarMaint 
      Caption         =   "Maintenance"
   End
End
Attribute VB_Name = "frmUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGoToDrillTap_Click(Index As Integer)

    'Update velocities from form
    myDrillTap.setVelDefaults frmUI.txtSpeeds(0), frmUI.txtSpeeds(1), frmUI.txtSpeeds(2)
    
    'Clear position to allow for repeat movements to location
    myDrillTap.clrPos
    
    'Request move destination from FSM
    Dim myDestination As moveTo
    
    If Index = 0 Then myDestination = moveToDrillNow
    If Index = 1 Then myDestination = moveToTapNow

    myFSM.setMoveDest (myDestination)

End Sub

Private Sub cmdResetCount_Click(Index As Integer)
    
    Dim myReset As uiCounter
    
    If Index = 0 Then myReset = countDrill
    If Index = 1 Then myReset = countTap
    
    myUI.resetCounter (myReset)
    
End Sub

Private Sub Form_Load()

    'Initialize UI
    myUI.initUI
    
    myDrillTap.initDrillTap
    
    
    'Load maintenance form
    Load frmMaintenance
    frmMaintenance.Hide
    
    'Clear maintenance Form Open flag
    maintenanceOpen = False
    
    
    
    'Initialize FSM
    myFSM.initializeFSM
    
    'Start UI form timer
    frmUI.tmrFSM.Enabled = True
    
    'Turn on key preview for joystick operation
    Me.KeyPreview = True


End Sub

Private Sub Form_LostFocus()

myUI.uiKeyUp = False
myUI.uiKeyRight = False
myUI.uiKeyDown = False
myUI.uiKeyLeft = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If myDrillTap.isJoyOn Then
        Select Case KeyCode
            Case vbKeyUp, vbKeyNumpad8
                myUI.uiKeyUp = True
            Case vbKeyRight, vbKeyNumpad6
                myUI.uiKeyRight = True
            Case vbKeyDown, vbKeyNumpad2
                myUI.uiKeyDown = True
            Case vbKeyLeft, vbKeyNumpad4
                myUI.uiKeyLeft = True
            Case Else
        End Select
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If myDrillTap.isJoyOn Then
        Select Case KeyCode
            Case vbKeyUp, vbKeyNumpad8
                myUI.uiKeyUp = False
            Case vbKeyRight, vbKeyNumpad6
                myUI.uiKeyRight = False
            Case vbKeyDown, vbKeyNumpad2
                myUI.uiKeyDown = False
            Case vbKeyLeft, vbKeyNumpad4
                myUI.uiKeyLeft = False
            Case Else
        End Select
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim msgBoxResponse As VbMsgBoxResult
    
    msgBoxResponse = MsgBox("Would you like to close the program?", vbOKCancel, "Close Drill & Tap?")
    
    If msgBoxResponse = vbCancel Then Cancel = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmConsole.cmdConnect(1).Enabled = True
    frmConsole.cmdConnect(2).Enabled = True

    Unload frmMaintenance
    
    'Stop UI form timer and re-enable console only timer
    frmUI.tmrFSM.Enabled = False
    
    frmConsole.timer6kRead.Enabled = True
    frmConsole.Show

End Sub

Private Sub cmdGO_Click()

    'Update velocities from Form
    myDrillTap.setVelDefaults frmUI.txtSpeeds(0), frmUI.txtSpeeds(1), frmUI.txtSpeeds(2)

    myUI.go

End Sub

Private Sub tmrFSM_Timer()

    'Update Console
    myCns.update

    'Run FSM
    myFSM.runFSM

End Sub

Private Sub topbarJoy_Click()

    myFSM.setJoystick boolTrue

End Sub

Private Sub topbarMaint_Click()

    maintenanceOpen = True
    frmMaintenance.Show
    Me.Show

End Sub

Private Sub topbarSet0_Click()

End Sub

Private Sub topbarSetDrill()

    Dim currentPos As myCoordinate
    Dim lastZero As myCoordinate

    Dim userReturn As VbMsgBoxResult
    
    currentPos.X = myDrillTap.getCoords(dX)
    currentPos.Y = myDrillTap.getCoords(dY)
    lastZero.X = myDrillTap.getLastZero(dX)
    lastZero.Y = myDrillTap.getLastZero(dY)
    
    myMsg = "Current position differs from zero by:" & vbCrLf & "X: " & currentPos.X & vbCrLf & "Y: " & currentPos.Y & vbCrLf & vbCrLf
    myMsg = myMsg & "And from the prior zero by:" & vbCrLf & "X: " & lastZero.X & vbCrLf & "Y: " & lastZero.Y & vbCrLf & vbCrLf
    myMsg = myMsg & "Would you like to reset the Zero position?"
    
    userReturn = MsgBox(myMsg, vbYesNo, "Set Zero?")
    
    If userReturn = vbYes Then
    
        'Disable Joystick
        myDrillTap.joyState False
        myFSM.setJoystick 0
    
        'Set Home
        myDrillTap.setHome
    End If


End Sub

Private Sub topbarSetTap()



End Sub

Private Sub txtDeep_Change(Index As Integer)

    'Ensure textbox entry is valid numeric value
    Dim txtbx As uiTxtbxs
    Select Case Index
        Case 0
            txtbx = txtDrillTo
        Case 1
            txtbx = txtTapTo
    End Select
    
    myUI.validate txtbx, txtDeep(Index).Text
    
    'Ensure entry is not out of bounds
    Select Case Index
        Case 0
            If CDbl(txtDeep(Index).Text) > 3 Then txtDeep(Index).Text = "3.00"
        Case 1
    End Select

End Sub

Private Sub txtDeep_LostFocus(Index As Integer)
    
    'Format the text in the textbox
    Dim tmpTxt As String
    tmpTxt = txtDeep(Index).Text
    
    If Index = 0 Then txtDeep(Index).Text = myUI.reformat(tmpTxt) Else txtDeep(Index).Text = Format(txtDeep(Index), "###0.0")

End Sub

Private Sub txtSpeeds_Change(Index As Integer)
    
    'Ensure textbox entry is valid numeric value
    Dim txtbx As uiTxtbxs
    Select Case Index
        Case 0
            txtbx = txtDrillSpeed
        Case 1
            txtbx = txtTapSpeed
        Case 2
            txtbx = txtJogSpeed
    End Select
    
    myUI.validate txtbx, txtSpeeds(Index).Text
    
    'Ensure that entries are not less than zero
    If txtSpeeds(Index).Text < 0 Then txtSpeeds(Index).Text = "0"
            
End Sub

Private Sub txtSpeeds_LostFocus(Index As Integer)
    
    'Format the text in the textbox
    Dim tmpTxt As String
    tmpTxt = txtSpeeds(Index).Text
    
    txtSpeeds(Index).Text = myUI.reformat(tmpTxt)

End Sub
