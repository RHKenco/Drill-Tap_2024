VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "6k Drill & Tap Program 2024"
   ClientHeight    =   2055
   ClientLeft      =   1260
   ClientTop       =   2130
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenUI 
      Caption         =   "Drill && Tap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   7
      Top             =   210
      Width           =   1830
   End
   Begin VB.OptionButton Option2 
      Caption         =   "RS232"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   180
      Width           =   1080
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ethernet"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   465
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Upload File"
      Height          =   420
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Download OS"
      Height          =   420
      Left            =   4200
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   225
      Top             =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1395
      TabIndex        =   2
      Top             =   210
      Width           =   1830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download File"
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   780
      Width           =   10500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo cmd1err

    If (Not connected) Then Exit Sub    'exit if not connected
    
    Timer1.Enabled = False          'disable response polling to avoid simultaneous read/write
    If (c6k.SendFile("") > 0) Then  'download program files - empty string means to prompt for filename
        c6k.Write ("TDIR" & Chr$(13))         'send TDIR command
    End If
    Timer1.Enabled = True           'enable response polling
    Text1.SetFocus
    Exit Sub
    
cmd1err:
    Disconnect
End Sub


Private Sub cmdConnect_Click()

On Error GoTo cmd2err
    Dim fh As Long      'file handle
Dim i


Shell ("C:\6K.BAT")
'("C:\VB PROGRAMS\6k Drill_Tap\6K.BAT")

For i = 1 To 100000
Next i

        fh = FreeFile   'get first avaiable handle
    
    'disconnect if already connected
    If connected Then
        Timer1.Enabled = False      'disable response polling
        Set c6k = Nothing           'disconnect and free up the c6k object
        connected = False           'set connection flag to false
    Else
    
    If Option1.Value Then   'ethernet
        ' use this code for Ethernet
        Dim ipaddr$
            ipaddr = "192.168.1.30"
    
        'attempt to open file where ip address is stored if file exists
        If Len(Dir$("ipaddr.dat")) Then
            Open "ipaddr.dat" For Input As #fh
                Line Input #fh, ipaddr
            Close #fh
        End If
    
        'prompt for ip address using default
        ipaddr = InputBox("Enter target IP Address.", "Port Setting", ipaddr)
        If Len(ipaddr) = 0 Then Exit Sub
    
        'save user specified ipaddr
        Open "ipaddr.dat" For Output As #fh
            Print #fh, ipaddr
        Close #fh
    
        'now attempt connection
        Set c6k = CreateObject("COM6SRVR.NET")
        If c6k.Connect(ipaddr) > 0 Then
            c6k.Write "TREV" & vbCr    'send TREV command
            Timer1.Enabled = True       'enable response polling
            connected = True            'set connected flag to true
        Else
            Timer1.Enabled = False      'disable response polling (default)
            connected = False           'set connected flag to false
            MsgBox "Connection attempt failed...", 0, "Status"
        End If
    
    Else    'RS232
        ' use this code for RS232
        Dim commport$
            commport = "2"
    
        'attempt to open file where ip address is stored if file exists
        If Len(Dir$("commport.dat")) Then
            Open "commport.dat" For Input As #fh
                Line Input #fh, commport
            Close #fh
        End If
    
        'prompt for com port number using default
        commport = InputBox("Enter PC COMPORT number.", "Port Setting", commport)
        If Len(commport) = 0 Then Exit Sub
    
        'save user specified ipaddr
        Open "commport.dat" For Output As #fh
            Print #fh, commport
        Close #fh
        
        
        Set c6k = CreateObject("COM6SRVR.RS232")
        If c6k.Connect(CInt(commport)) > 0 Then
            c6k.Write "TREV" & vbCr     'send TREV command
            Timer1.Enabled = True       'enable response polling
            connected = True            'set connected flag to true
        Else
            Timer1.Enabled = False      'disable response polling (default)
            connected = False           'set connected flag to false
            MsgBox "Connection attempt failed...", 0, "Status"
        End If
    End If
    
    End If
    
    If connected Then
        cmdConnect.Caption = "Disconnect"
        Option1.Enabled = False
        Option2.Enabled = False
    Else
        cmdConnect.Caption = "Connect"
        Option1.Enabled = True
        Option2.Enabled = True
        Set c6k = Nothing           'release the comm server
    End If
    Text1.SetFocus
    Exit Sub
    
cmd2err:
    Disconnect
End Sub

Private Sub Command3_Click()
On Error GoTo cmd3err

    If (connected And Option2.Value) Then
        Timer1.Enabled = False      'disable response polling
        c6k.SendOS ("")             'download the Operating System - prompt for OS file
        Timer1.Enabled = True       'enable response polling
        c6k.Write ("TREV" & Chr$(13))
        Text1.SetFocus
    Else
        MsgBox "Operating System download is only supported via RS232.", 0, "OS Download Unavailable"
    End If
    
    Exit Sub
    
cmd3err:
    Disconnect
End Sub

Private Sub Command4_Click()
On Error GoTo cmd4err

    If (Not connected) Then Exit Sub    'exit if not connected
    
    Timer1.Enabled = False      'disable response polling to avoid simultaneous read/write
    c6k.GetFile ("")            'upload program files - empty string means to prompt for filename
    Timer1.Enabled = True       'enable response polling
    Text1.SetFocus
    Exit Sub
    
cmd4err:
    Disconnect
End Sub


Private Sub Command5_Click()
 Shell ("C:\6K.BAT")

End Sub

Private Sub cmdOpenUI_Click()
    If (Timer1.Enabled And Option1.Value) Then
        Me.Hide
        frmUI.Show
       
    Else
        MsgBox "An ethernet connection is needed for Fast Staus display.", 0, "Display Unavailable"
    End If
End Sub

Private Sub Form_GotFocus()
    If connected Then Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    connected = False   'connection disabled by default
End Sub


Private Sub Form_LostFocus()
    Timer1.Enabled = False
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Text1.Width = Me.Width - 345
        Text1.Height = Me.Height - 1275
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'make sure to disconnect on unload
    If connected Then
        Set c6k = Nothing
    End If
End Sub


Private Sub Text1_Change()
    'the text box has a finite buffer so
    'make sure it doesn't overflow
    
    If Len(Text1.Text) > 16000 Then
        Text1.Text = Right$(Text1.Text, 500)    'buffer just the last 500 characters
    End If
    
End Sub

Private Sub Text1_DblClick()
    Text1.Text = ""     'clear the terminal display
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'this routine processes the terminal's key presses
On Error GoTo text1keypress_error

    'exit if not connected
    If Not connected Then
        KeyAscii = 0
    End If
    
    Dim temp%
    Static buffer$      'local command buffer
    
    'perform action based on value of key being pressed
    Select Case KeyAscii
        'backspace
        Case 8
            If Len(buffer) > 0 Then buffer = Left$(buffer, Len(buffer) - 1) 'erase one char from buffer
            
        
        'CR or colon - 6000 command delimeter
        Case 13, Asc(":")
            If Format$(buffer, ">") = "CLS" Then      'internal clear screen command
                Text1.Text = ""
                KeyAscii = 0
            Else
                buffer = buffer & Chr$(13)      'append the CR
                Timer1.Enabled = False          'disable response polling to avoid simultaneous read/write
                temp = c6k.Write(buffer)        'send commands to 6k
                Timer1.Enabled = True           'enable response polling
            End If
            buffer = ""                         'empty the command local buffer
        
        
        'anything else just add to the buffer
        Case Else
            buffer = buffer & Chr$(KeyAscii)    'append char to the local command buffer
            
    End Select
    Exit Sub
    
text1keypress_error:
    Disconnect
End Sub


Private Sub Timer1_Timer()
On Error GoTo timer1err

    'this timer routine polls for response from the controller
    Dim temp$
    temp = c6k.Read()                           'get response
    If Len(temp) Then Text1.SelText = temp      'if not empty then display in the text box
    Exit Sub
    
timer1err:
    Disconnect
End Sub


