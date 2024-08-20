VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class6kConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- c6k Communication/Console Class -----
'|
'|
'|
'|====================================================================================================================|

Option Explicit


'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===

'--------------------------- Class Enumerations & Constants -----------------------------------------------------------

Private readBuffer As String
Private writeBuffer As String
Private writeDisplayString As String

Const c6kEntry = "c6k   > "
Const usrEntry = "user  > "
Const bnkEntry = "        "

Const maxConsoleChars = 5000

'Command History Variables
Private cmdPoint As Integer
Private cmdH1 As String
Private cmdH2 As String
Private cmdH3 As String
Private cmdH4 As String
Private cmdH5 As String
Private cmdH6 As String
Private cmdH7 As String
Private cmdH8 As String
Private cmdH9 As String


'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=








'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------

Public Sub initConsole()

    'Initialize command history
    cmdPoint = 9
    cmdH1 = ""
    cmdH2 = ""
    cmdH3 = ""
    cmdH4 = ""
    cmdH5 = ""
    cmdH6 = ""
    cmdH7 = ""
    cmdH8 = ""
    cmdH9 = ""

End Sub



Public Sub writeQueue(writeStr As String)

    If writeStr <> "" Then writeBuffer = writeBuffer & writeStr

End Sub


Private Sub write6kNow(writeStr As String)

    'Add command to the write buffer, to be entered into the console
    If wirteStr <> "" Then
        writeBuffer = writeBuffer & writeStr
    End If
    
    writeToConsole writeBuffer
    
    'Write the command to the 6k
    c6k.Write (writeBuffer)
    
    writeBuffer = ""

End Sub

Private Sub writeToConsole(writeStr As String)

    If writeStr <> "" Then
    
        writeDisplayString = writeDisplayString & usrEntry & writeStr & Chr$(13)
    
    End If

End Sub


Private Sub read6kNow()

    Dim readString As String
    
    'Read from the 6k
    readString = c6k.Read()
    
    'Append the read 6k data to the end of the console buffer
    readBuffer = readBuffer + readString

End Sub

Public Function getHistory(Index As Integer)

    Dim rtnString As String
    Dim i As Integer
    
    rtnString = ""
    
    If i < 1 Then i = 1
    If i > 9 Then i = 9
    i = Index + cmdPoint
    
    'Ensure that the value does not cycle
    If i > 8 Then i = i - 8
    Do:
        Select Case i
            Case 1
                rtnString = cmdH1
            Case 2
                rtnString = cmdH2
            Case 3
                rtnString = cmdH3
            Case 4
                rtnString = cmdH4
            Case 5
                rtnString = cmdH5
            Case 6
                rtnString = cmdH6
            Case 7
                rtnString = cmdH7
            Case 8
                rtnString = cmdH8
            Case 9
                rtnString = cmdH9
        End Select
        
        'If there are empty history values, return oldest value. If no values, exit
        If rtnString = "" Then i = i - 1
        If i = 0 Then Exit Do
        
    Loop Until rtnString <> ""
    
    getHistory = rtnString

End Function

Public Sub setHistory(newCmd As String)

    'Find correct location and store command
    Select Case cmdPoint
        Case 1
             cmdH1 = newCmd
        Case 2
             cmdH2 = newCmd
        Case 3
            cmdH3 = newCmd
        Case 4
            cmdH4 = newCmd
        Case 5
            cmdH5 = newCmd
        Case 6
            cmdH6 = newCmd
        Case 7
            cmdH7 = newCmd
        Case 8
            cmdH8 = newCmd
        Case 9
            cmdH9 = newCmd
    End Select
    
    'Move cmdPointer up along storage array
    cmdPoint = cmdPoint - 1
    If cmdPoint = 0 Then cmdPoint = 9

End Sub




'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=













'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===-- Advanced Functions - Operate Console, Etc. ---===---===---===---===---===---===---===---===

Public Sub update()
    
    write6kNow (writeBuffer)
    
    read6kNow
    
    updConsole

End Sub

Private Sub updConsole()

    Dim nextChar As Char
    Dim i As Integer
    
    'Until there is a carriage return in the output, the limit is reached, or the end of the buffer is reached, Output
    For i = 0 To Min(Len(readBuffer), 300)
        'Select the next character in the read buffer
        nextChar = Left(readBuffer, 1)
    
        'Print that character to the console
        
        'If the character is a carriage return
        If nextChar = Chr$(13) Then
            
            'move to the next line
            frmConsole.txtConsole(0).Text = txtConsole(0).Text & Chr$(13)
            
            'If a command has been written to the 6k, write it to the console when possible
            If writeDisplayString <> "" Then
                frmConsole.txtConsole(0).Text = frmConsole.txtConsole(0).Text & writeDisplayString
                writeDisplayString = ""
            End If
            
            'Add space to front of line
            frmConsole.txtConsole(0).Text = frmConsole.txtConsole(0).Text & c6kEntry
            
            'Exit loop each time a line is completed
            Exit For
            
        'Otherwise, add it to the console
        Else
        
            'Add character to textbox
            frmConsole.txtConsole(0).SelStart = Len(tfrmConsole.xtConsole(0).Text)
            frmConsole.txtConsole(0).Text = frmConsole.txtConsole(0).Text & nextChar
            frmConsole.txtConsole(0).SelStart = Len(frmConsole.txtConsole(0).Text)
            
        End If
        
        'Remove the character from the read buffer
        readBuffer = Right(readBuffer, (Len(readBuffer) - 1))
    Next i
    
    'Truncate the console to ensure it does not overflow
    If Len(frmConsole.txtConsole(0).Text) Then frmConsole.txtConsole(0).Text = Right(frmConsole.txtConsole(0).Text, maxConsoleChars)

End Sub

Public Sub commandLineEnter()
    
    Dim cmdInput As String
    
    'Pull the text from the console entry window
    cmdInput = frmConsole.txtConsole(1).Text
    
    'Remove the user entry prefix from the string, then remove leading or trailing spaces
    cmdInput = Right(cmdInput, (Len(cmdInput) - Len(usrEntry) + 1))
    cmdInput = Trim(cmdInput)
    
    'Ensure that a command terminator is the last character sent
    If Left(cmdInput, 1) <> ":" Then cmdInput = cmdInput & ":"
    

    'Write input to the queue
    writeQueue (cmdInput)
    
    'Clear the textbox for new entry
    frmConsole.txtConsole(1).Text = usrEntry

    'Store the command in the history array
    setHistory (cmdInput)

End Sub


'---===---===---===---===-- Advanced Functions - Operate Console, Etc. ---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
















