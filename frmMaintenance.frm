VERSION 5.00
Begin VB.Form frmMaintenance 
   BackColor       =   &H00400000&
   Caption         =   "Maintenance"
   ClientHeight    =   9630
   ClientLeft      =   12105
   ClientTop       =   705
   ClientWidth     =   5880
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   5880
   Begin VB.Frame frameMaintPos 
      BackColor       =   &H00400000&
      Caption         =   "Manual Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   7920
      Width           =   5415
      Begin VB.CommandButton cmdMaintGoTo 
         Caption         =   "Go To"
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
         Left            =   3480
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtMaintPos 
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
         Left            =   1080
         TabIndex        =   3
         Text            =   "0.000"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtMaintPos 
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
         Left            =   1080
         TabIndex        =   2
         Text            =   "0.000"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblMaintGo 
         BackColor       =   &H00400000&
         Caption         =   "Y:                      in"
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
         TabIndex        =   5
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblMaintGo 
         BackColor       =   &H00400000&
         Caption         =   "X:                      in"
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
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame frameMaintIO 
      BackColor       =   &H00400000&
      Caption         =   "I/O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   7815
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   "16 - Input:    Clamp Off Button ===================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   7320
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   "15 - Input:    ESTOP Button ======================"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6840
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   "14 - Input:    Clamp On Button ===================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6360
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   "13 - Input:    Y-Travel Limit ======================"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5880
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   "12 - Input:    Joystick Y- ========================"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5400
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   "11 - Input:    Joystick Y+ ========================"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4920
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   "10 - Input:    Joystick X- ========================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4440
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 9 - Input:    Joystick X+ ========================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3960
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 8 - Output: N/A ==============================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3600
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 7 - Output: N/A ==============================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3240
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 6 - Output: Coolant Motor ======================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2760
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 5 - Output: Tap Motor =========================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2280
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 4 - Output: Drill Motor =========================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 3 - Output: Clamp Pump ========================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 2 - Output: Clamp Solenoid ======================"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   5055
      End
      Begin VB.CommandButton cmdMaintIO 
         Caption         =   " 1 - Output: N/A ==============================="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Menu topbarEnIO 
      Caption         =   "Enable I/O"
   End
   Begin VB.Menu topbarOpenConsole 
      Caption         =   "Show Console"
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMaintIO_Click(Index As Integer)

    myDrillTap.maintActive (Index + 1), mToggle

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    maintenanceOpen = False
    
    'Ensure active maintenance mode is disabled
    myDrillTap.maintActive 0, mActiveOff
    
    'If closed from the window, hide form instead of unloading
    If UnloadMode = vbFormControlMenu Then
        'Cancel overall unload and hide form
        Cancel = 1
        Me.Hide
        frmUI.Show
    End If
    
End Sub


Private Sub topbarEnIO_Click()

    'Enable active maintenance mode
    myDrillTap.maintActive 0, mActiveOn
    Me.Show
    

End Sub

Private Sub topbarOpenConsole_Click()
    
    frmConsole.Show
    
End Sub

Private Sub txtMaintPos_Change(Index As Integer)
    
    'Ensure textbox entry is valid numeric value
    Dim txtbx As uiTxtbxs
    Select Case Index
        Case 0
            txtbx = txtMaintX
        Case 1
            txtbx = txtMaintY
    End Select
    
    myUI.validate txtbx, txtMaintPos(Index).Text
    
End Sub

Private Sub txtMaintPos_LostFocus(Index As Integer)
    
    'Format the text in the textbox
    Dim tmpTxt As String
    tmpTxt = txtMaintPos(Index).Text
    
    txtMaintPos(Index).Text = myUI.reformat(tmpTxt)

End Sub
