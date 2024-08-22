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
Const bnkEntry = "          "
Const cmdPrompt = " > "

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


    frmConsole.txtConsole(1).Text = cmdPrompt

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

Private Sub write6k()

    writeToConsole writeBuffer
    
    'Write the commands to the 6k
    c6k.Write (writeBuffer)
    
    writeBuffer = ""

End Sub

Public Sub write6kNow(writeStr As String)

    'Add command to the write buffer, to be entered into the console
    writeBuffer = writeStr & writeBuffer

    'Write the command to the 6k
    c6k.Write (writeStr)

End Sub

Private Sub writeToConsole(writeStr As String)

    If writeStr <> "" Then
    
        writeDisplayString = writeDisplayString & usrEntry & writeStr & vbCrLf
    
    End If

End Sub


Private Sub read6k()

    Dim readString As String
    
    'Read from the 6k
    readString = c6k.read()
    
    'Append the read 6k data to the end of the console buffer
    If readString <> "" Then
        readBuffer = readBuffer & readString
    End If
    
End Sub

Public Function getHistory(Index As Integer)

    Dim rtnString As String
    Dim i As Integer
    
    rtnString = ""
    
    If Index < 1 Then Index = 1
    If Index > 9 Then Index = 9
    i = Index + cmdPoint
    
    'Ensure that the value does not cycle
    If i > 9 Then i = i - 9
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
        If i = 0 Then i = 9
        If i = cmdPoint Then Exit Do
        
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
    
    'If the 6k is not overloaded with commands, add commands
    'If c6k.CommandCount < 5 Then
    write6k
    
    read6k
    
    updConsole
    
    
    frmConsole.lblCmdCount.Caption = c6k.CommandCount
    frmConsole.Refresh

End Sub

Public Sub CLEAR()

    readBuffer = ""
    writeBuffer = ""
    writeDisplayString = ""

End Sub

Private Sub updConsole()

    Dim nextChar As String
    Dim i As Integer
    Dim Limit As Integer
    Dim tempBuffer As String
    Dim addBuffer As String
    
    Static newLine As Boolean
    
    tempBuffer = readBuffer
    
    If Len(tempBuffer) > 300 Then Limit = 300 Else Limit = Len(tempBuffer)
    
    If (Limit = 0) Then
        If writeDisplayString <> "" Then
            GoTo DisplayWrite
        Else
            Exit Sub
        End If
    End If
    
    'Until there is a carriage return in the output, the limit is reached, or the end of the buffer is reached, Output
    For i = 1 To Limit
        'Select the next character in the read buffer
        nextChar = Left(tempBuffer, 1)
    
        'Print that character to the console
        
        'If the character is a carriage return
        If (nextChar = Chr$(13)) Or (nextChar = Chr$(10)) Then
            
            'Remove the character from the read buffer
            tempBuffer = Right(tempBuffer, (Len(tempBuffer) - 1))
            
            
            If addBuffer <> "" Then
                'Add generated line to the console
                frmConsole.txtConsole(0).SelStart = Len(frmConsole.txtConsole(0).Text)
                frmConsole.txtConsole(0).Text = frmConsole.txtConsole(0).Text & addBuffer
                addBuffer = ""
                frmConsole.txtConsole(0).SelStart = Len(frmConsole.txtConsole(0).Text)
            End If
            
            If newLine Then
                GoTo DisplayNextI
            End If
            
DisplayWrite:
            
            'move to the next line
            frmConsole.txtConsole(0).Text = frmConsole.txtConsole(0).Text & vbCrLf
            newLine = True
            
            'If a command has been written to the 6k, write it to the console
            If writeDisplayString <> "" Then
                'When writing stored commands to the console, add extra space before, then add two spaces after and show 6k entry indicator
                frmConsole.txtConsole(0).Text = vbCrLf & frmConsole.txtConsole(0).Text & writeDisplayString & vbCrLf & vbCrLf & c6kEntry
                writeDisplayString = ""
                newLine = False
            End If
            
            
            'After a line is printed, exit loop
            Exit For
            
        'Otherwise, add it to the console
        Else
        
            'Add space to front of line if entry is starting on a clear line
            If newLine Then frmConsole.txtConsole(0).Text = frmConsole.txtConsole(0).Text & bnkEntry
            If nextChar <> "" Then newLine = False
        
            'Add character to the addbuffer
            addBuffer = addBuffer & nextChar
            'Remove the character from the read buffer
            tempBuffer = Right(tempBuffer, (Len(tempBuffer) - 1))
            
        End If

DisplayNextI:

    Next i
    
    'Set read buffer to the truncated tempBuffer
    readBuffer = tempBuffer
    
    'Truncate the console to ensure it does not overflow
    If Len(frmConsole.txtConsole(0).Text) Then frmConsole.txtConsole(0).Text = Right(frmConsole.txtConsole(0).Text, maxConsoleChars)

    'Ensure cursor is at end of window
    frmConsole.txtConsole(0).SelStart = Len(frmConsole.txtConsole(0).Text)

End Sub

Public Sub commandLineEnter()
    
    Dim cmdInput As String
    
    'Pull the text from the console entry window
    cmdInput = frmConsole.txtConsole(1).Text
    
    'Remove the user entry prefix from the string, then remove leading or trailing spaces
    cmdInput = Right(cmdInput, (Len(cmdInput) - (Len(cmdPrompt) - 1)))
    cmdInput = Trim(cmdInput)
    
    'Ensure that a command terminator is the last character sent
    If Left(cmdInput, 1) <> ":" Then cmdInput = cmdInput & ":"
    

    'Write input to the queue
    write6kNow (cmdInput)
    
    'Clear the textbox for new entry
    frmConsole.txtConsole(1).Text = cmdPrompt
    frmConsole.txtConsole(1).SelStart = Len(frmConsole.txtConsole(1).Text)

    'Store the command in the history array
    setHistory (cmdInput)

End Sub


'---===---===---===---===-- Advanced Functions - Operate Console, Etc. ---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
















