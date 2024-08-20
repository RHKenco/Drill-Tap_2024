VERSION 5.00
Begin VB.Form frmConsole 
   Caption         =   "Unishank Drill & Tap 2024 - Console"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConsole 
      BackColor       =   &H00800000&
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
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Text            =   "  >  "
      Top             =   5040
      Width           =   10575
   End
   Begin VB.TextBox txtConsole 
      BackColor       =   &H00800000&
      ForeColor       =   &H8000000E&
      Height          =   4215
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   720
      Width           =   10575
   End
   Begin VB.Timer timer6kRead 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10320
      Top             =   120
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Open Drill/Tap"
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
      Index           =   2
      Left            =   8040
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Disconnect"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lbl6kConsole 
      Caption         =   "Drill - Tap C6k2 Console:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Const ipPath = "C:\Drill-Tap_2024\savedData\ipAddr.txt"



Private Sub cmdConnect_Click(Index As Integer)

    Select Case Index
        Case 0 'Connect Button
        
        
            'Run batch file to establish ethernet connection to the 6k
            Shell ("C:\6K.BAT")
            
            'Set up IP Address Default
            Dim ipAddr As String
            Dim tempIP As String
            ipAddr = "192.168.1.30"
    
            'Store IP Address in file
            If Dir(ipPath) <> "" Then
                Open ipPath For Input As #1
                Input #1, ipAddr
                Close #1
            End If
    
            'Prompt user for IP address
            tempIP = InputBox("Enter target IP Address:", "Port Setting", ipAddr)
            If Len(ipAddr) = 0 Then Exit Sub
            
            'If the entered IP address differs from the loaded one, save the new address
            If ipAddr <> tempIP Then
                Open ipPath For Output As #1
                Write #1, tempIP
                Close #1
                ipAddr = tempIP
            End If
    
            'Attempt Connection
            Set c6k = CreateObject("COM6SRVR.NET")
            If c6k.Connect(ipAddr) > 0 Then
                myCns.writeQueue "TREV:"    'send TREV command
                timer6kRead.Enabled = True  'enable response polling
                connected = True            'set connected flag to true
            Else
                timer6kRead.Enabled = False 'disable response polling (default)
                connected = False           'set connected flag to false
                MsgBox "Connection attempt failed...", 0, "Status"
            End If
            
            
            'Read/Write to the 6k periodically
            timer6kRead.Enabled = True
            
            
            cmdConnect(0).Visible = False
            cmdConnect(1).Visible = True
            cmdConnect(2).Visible = True
            
            cmdConnect(0).Enabled = False
            cmdConnect(1).Enabled = True
            cmdConnect(2).Enabled = True
            
            
        Case 1 'Disconnect Button
            
            'Disable 6k read/write
            timer6kRead.Enabled = False
            
            'Disconnect from 6k
            Set c6k = Nothing
            
            
            cmdConnect(0).Visible = True
            cmdConnect(1).Visible = False
            cmdConnect(2).Visible = False
            
            cmdConnect(0).Enabled = True
            cmdConnect(1).Enabled = False
            cmdConnect(2).Enabled = False
            
        Case 2 'Launch Button
        
            Me.Hide
            frmUI.Show
            
            
    End Select

End Sub

Private Sub Form_Load()

    'Connection is disabled by default
    connected = False

    'Initialize console class
    myCns.initConsole
    

    'Initialize Form Controls
    cmdConnect(0).Visible = True
    cmdConnect(1).Visible = False
    cmdConnect(2).Visible = False
    
    cmdConnect(0).Enabled = True
    cmdConnect(1).Enabled = False
    cmdConnect(2).Enabled = False

    Me.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Ensure all forms are closed
    Dim frm As Form
    For Each frm In Forms
        If frm.Name <> formNameLeftOpen Then
            Unload frm
            Set frm = Nothing
        End If
    Next

    'Disconnect from 6k
    If connected Then cmdConnect_Click (1)

End Sub

Private Sub timer6kRead_Timer()

    myCns.update

End Sub

Private Sub txtConsole_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    'If any entry is in console textbox instead of command line textbox, ignore
    If Index = 0 Then Exit Sub
    
    Static pointer As Integer

    Select Case KeyCode
        Case vbKeyUp
            pointer = pointer + 1
            If pointer > 9 Then pointer = 9
            e.Handled = True
        Case vbKeyDown
            If pointer <> 0 Then
                pointer = pointer - 1
                If pointer < 1 Then pointer = 1
            End If
            e.Handled = True
        Case vbKeyReturn 'Enter is pressed
            myCns.commandLineEnter
            pointer = 0
        Case Else
    End Select

End Sub
