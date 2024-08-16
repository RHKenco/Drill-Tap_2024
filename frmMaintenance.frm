VERSION 5.00
Begin VB.Form frmMaintenance 
   BackColor       =   &H00400000&
   Caption         =   "Maintenance"
   ClientHeight    =   10740
   ClientLeft      =   17895
   ClientTop       =   3165
   ClientWidth     =   5880
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   5880
   Begin VB.Frame frmMaintPos 
      BackColor       =   &H00400000&
      Caption         =   "Manual Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   7320
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
         Height          =   1335
         Left            =   3600
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtMaintPos 
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
         TabIndex        =   3
         Text            =   "0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtMaintPos 
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
         TabIndex        =   2
         Text            =   "0"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblMaintGo 
         BackColor       =   &H00400000&
         Caption         =   "Y:                      in"
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
         Left            =   480
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblMaintGo 
         BackColor       =   &H00400000&
         Caption         =   "X:                      in"
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
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.Frame frmMaintIO 
      BackColor       =   &H00400000&
      Caption         =   "I/O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
