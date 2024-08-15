VERSION 5.00
Begin VB.Form frmUI 
   BackColor       =   &H00400000&
   Caption         =   "Drill & Tap"
   ClientHeight    =   10515
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11910
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
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
      Top             =   3000
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
            Size            =   15
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
            Size            =   15
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
            Size            =   15
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
      Top             =   360
      Width           =   2655
      Begin VB.Label lblDRO 
         BackColor       =   &H00400000&
         Caption         =   "Y:     0.000 in"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
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
            Size            =   15
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
      Top             =   360
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
            Text            =   "1.625"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtDeep 
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
            Index           =   3
            Left            =   2760
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
            Left            =   2760
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
         Left            =   120
         Top             =   360
      End
      Begin VB.Frame frameSpeeds 
         BackColor       =   &H00400000&
         Caption         =   "  System Speeds  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
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
            Text            =   "30"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txtSpeeds 
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
            Text            =   "15"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtSpeeds 
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
            Text            =   "0.625"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblSpeeds 
            BackStyle       =   0  'Transparent
            Caption         =   "Drill Speed:"
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
            Width           =   2055
         End
         Begin VB.Label lblSpeeds 
            BackStyle       =   0  'Transparent
            Caption         =   "Tap Speed:"
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
            Width           =   2055
         End
         Begin VB.Label lblSpeeds 
            BackStyle       =   0  'Transparent
            Caption         =   "Jog Speed:"
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
            Width           =   2535
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
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   7200
      Width           =   11175
   End
   Begin VB.Menu topbarJoy 
      Caption         =   "Joystick"
   End
   Begin VB.Menu topbarSet0 
      Caption         =   "Set Home"
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
Private Sub Form_Load()


Me.KeyPreview = True


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyUp
        myUI.uiKeyUp = True
    Case vbKeyRight
        myUI.uiKeyRight = True
    Case vbKeyDown
        myUI.uiKeyDown = True
    Case vbKeyLeft
        myUI.uiKeyLeft = True
    Case Else
End Select

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyUp
        myUI.uiKeyUp = False
    Case vbKeyRight
        myUI.uiKeyRight = False
    Case vbKeyDown
        myUI.uiKeyDown = False
    Case vbKeyLeft
        myUI.uiKeyLeft = False
    Case Else
End Select

End Sub



Private Sub cmdGO_Click()

    myUI.go

End Sub

Private Sub tmrFSM_Timer()

    myFSM.runFSM

End Sub

Private Sub topbarJoy_Click()

    myFSM.setJoystick

End Sub

Private Sub topbarMaint_Click()

    myFSM.setMaintenance
    frmMaintenance.Show

End Sub

Private Sub topbarSet0_Click()

    Dim currentPos As myCoordinate
    Dim lastZero As myCoordinate

    Dim userReturn As VbMsgBoxResult
    
    currentPos = myDrillTap.getPos
    lastZero = myDrillTap.getLastZero
    
    myMsg = "Current position differs from zero by:" & vbCrLf & "X: " & currentPos.X & vbCrLf & "Y: " & currentPos.Y & vbCrLf & vbCrLf
    myMsg = myMsg & "And from the prior zero by:" & vbCrLf & "X: " & lastZero.X & vbCrLf & "Y: " & lastZero.Y & vbCrLf & vbCrLf
    myMsg = myMsg & "Would you like to reset the Zero position?"
    
    userReturn = MsgBox(myMsg, vbYesNo, "Set Zero?")
    
    If userReturn = vbOK Then myDrillTap.setHome

End Sub
