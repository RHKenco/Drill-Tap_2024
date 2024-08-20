VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class6kXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- c6k 2-Axis Control Class -----
'|
'|
'|
'|====================================================================================================================|

Option Explicit


'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===

'--------------------------- Class Enumerations & Constants -----------------------------------------------------------

    Const myDriveScale = 5000   'Steps per Revolution - Check Dip Switches on Drives
    Const myScaleX = 50000  'Steps per Inch - Ball screw & 1"/Rev + 10:1 transmission -> DriveScale * 10
    Const myScaleY = 50000

    Enum myMotors
        mNull
        m1
        m2
    End Enum

    'enumerations for the input to the moveTo command
    Enum moveType
        movNone
        movLine
        movX
        movXY
        movY
        movYX
        movY0XY
    End Enum
      
'--------------------------- Class Variables --------------------------------------------------------------------------


    'Write-only Variables
    
    
    'Internal Variables
        'Variable defining the move complete acknowledgement virtual input pin. Used to stop windows code until mov end
        Private movComplete As Integer
        'Variable that indicates that a move has been initiated,but not completed
        Private movStarted As Boolean
        
        
        Private motorX As Integer
        Private motorY As Integer
        
        Private zeroOffset As myCoordinate
        
        Private YlimitInput As Integer      'Input Pin for Y-Travel Limit Switch
        Private ESTOPpin As Integer         'Input Pin for ESTOP
        

'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=














'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------
    Public Sub init6k(myMotorX As myMotors, myMotorY As myMotors, movAcknInpt As Integer, ESTOP_Input_Pin As Integer, Y_Limit_Input_Pin As Integer)
        
        'Initialize 6k:
        
        'Enable FastStatus
            c6k.fsenabled = True
        
        'Enable input-selected programs
            c6k.Write "INSELP1,50:"
        
        'Set up axes
            'Set axes to behave as steppers
            c6k.Write "AXSDEF00:"
            
            'Set Scaling Factors to correct motors
            Dim strDriveScale As String
            Dim strResScale As String
            
            strDriveScale = Str(myDriveScale) & "," & Str(myDriveScale)
            
            If myMotorX = m1 Then
                strResScale = Str(myScaleX) & "," & Str(myScaleY)
            Else
                strResScale = Str(myScaleY) & "," & Str(myScaleX)
            End If
                
            c6k.Write "DRES" & strDriveScale & ":SCLD" & strResScale & ":SCLV" & strResScale & ":SCLA" & strResScale & ":"
            
            'Enable Scaling
            c6k.Write "SCALE1:"
            
        'Set up IO
            
            YlimitInput = Y_Limit_Input_Pin
            
            'Set up Y-Axis negative limit switch
            Dim strLimitSwitch As String
            
            If myMotorY = m1 Then
                strLimitSwitch = Str(YlimitInput) & "-1S:"
            Else
                strLimitSwitch = Str(YlimitInput) & "-2S:"
            End If
            
            c6k.Write "1INFNC" & strLimitSwitch
            
            'Enable EOT Limit for negative Y-Axis
            Dim strEOT As String
            
            If myMotorY = m1 Then strEOT = "0,1" Else strEOT = "1,0"
            
            c6k.Write "LH" & strEOT & ":LHAD100,100:"
            
            
            'Set up estop button
            ESTOPpin = ESTOP_Input_Pin
            
            c6k.Write "1INFNC" & Str(ESTOPpin) & "-C:"
            
            
        'Initialize the 6k's motion parameters
        reset6k
        
        
        'Set initialized internal variables
        motorX = myMotorX
        motorY = myMotorY
        
        movComplete = movAcknInpt
        
        
    End Sub
    
    'reset6k - resets 6k state to initialized state.
    Public Sub reset6k()
    
        'Initialize 6k motion parameters
        c6k.Write "DRIVE1,1:"
        
        'Ensure Continuous Command Execution is Disabled
        c6k.Write "COMEXC0:"
        
        'Set default velocity to 0, Set initial values for Acceleration, Tie Decell to accel
        c6k.Write "V0,0:A10,10:AD0,0:"
        
        'Set jog motion parameters
        c6k.Write "JOGA10,5:JOGAD300,150:JOGVH6,3:JOGVL30,15:"
    
    End Sub
    
    'Function specifically for operating the hydraulic clamp on the Drill-Tap
    Public Sub initClamp(mySol As Integer, myPump As Integer, myOnInput As Integer, myOffInput As Integer)
    
        c6k.Write "DEL CLAMPOFF:DEF CLAMPOFF:"
            setOut mySol, True
        c6k.Write "END:"
        
        c6k.Write "DEL CLAMPON:DEF CLAMPON:"
        
            setOut mySol, True
            setOut myPump, True
            c6k.Write "T1.9:"
            setOut mySol, False
            c6k.Write "T.5:"
            setOut myPump, False
    
        c6k.Write "END:"
        
        c6k.Write "1INFNC" & Str(myOffInput) & "P1:"
        c6k.Write "1INFNC" & Str(myOnInput) & "P2:"
    
    End Sub
    
    
    'Estop - disable all outputs
    Public Function chkESTOP() As Boolean
    
        'Return false by default
        chkESTOP = False
    
        'If ESTOP pin is not raised, exit
        If Not chkIn(ESTOPpin) Then Exit Function
        
        'Imediate stop command, imediate shut down all outputs, disable motors
        'c6k.Write "!AD100,100:!S:!OUTALL1,32,0:DRIVE00:"
        chkESTOP = True
    
    End Function
    

'---------------------------- Internal Functions ----------------------------------------------------------------------
    
    
    'genCmd - Function to generate commands for the 6k to send values to a specific motor
    '   - Inputs:
    '       - motor - input of which motor to send the command to: Will be motorX or motorY
    '       - cmdStr - A string containing the desired command. Ex: "V"
    '       - valStr - A string containing the value to be sent to the 6k. Ex: "0.125"
    '   - Returns:
    '       - rtnStr - A string containing the completed command. Ex: "V,0.125:"
    Private Function genCmd(motor As Integer, cmdStr As String, valStr As String) As String
        
        'Create temporary string
        Dim rtnStr As String
        
        rtnStr = cmdStr
        
        'Place cmdStr according to motor
        Select Case motor
            Case 1
                rtnStr = rtnStr & valStr
            Case 2
                rtnStr = rtnStr & "," & valStr
        End Select
        
        rtnStr = rtnStr & ":"
        
        genCmd = rtnStr
    
    End Function
    
    
    'genCmdBin - Function to generate the binary argument for commands such as "GO" and "MOV = b"
    '   - Inputs:
    '       - motor - input of which motor to send the command to: Will be motorX or motorY
    '       - optX - A boolean to determine if values need to be X
    '       - opt10 - A boolean to determine if the X values need to replace 1 or 0
    '   - Returns:
    '       - rtnStr - A string containing the completed command. Ex: "V,0.125:"
    Private Function genCmdBin(motor As Integer, inv As Boolean, optX As Boolean, opt10 As Boolean) As String
    
        'Create Temporary string
        Dim rtnStr As String
        
        'Generate initial binary based on motor
        Select Case motor
            Case 1
                If inv Then rtnStr = "01" Else rtnStr = "10"
            Case 2
                If inv Then rtnStr = "10" Else rtnStr = "01"
        End Select
        
        'Adjust binary with "Don't Care" (X) as necessary
        If optX Then
            If opt10 Then
                rtnStr = Replace(rtnStr, "1", "X")
            Else
                rtnStr = Replace(rtnStr, "0", "X")
            End If
        End If
                
        genCmdBin = rtnStr
    
    
    End Function


'---------------------------- Generic Functions -----------------------------------------------------------------------
        
    
    'setZero - Function to set the zero location of the 6k
    Public Sub setZero(offsetX As Double, offsetY As Double)
    
        zeroOffset.X = offsetX
        zeroOffset.Y = offsetY
    
        'Sets zero to the offset location
        c6k.Write "PSET" + Str(offsetX) + "," + Str(offsetY) + ":"
    
    End Sub
    
    'setZero - Function to set the zero location of the 6k to the current location
    Public Sub setZeroHere()
    
        'Sets zero to the offset location
        c6k.Write "PSET0,0:"
    
    End Sub
    
    
    'getPosX - Function to get the position of the X-axis
    Public Function getPosX() As Double
        
        getPosX = fsinfo.MotorPos(motorX) - zeroOffset.X
        
    End Function
    
    'getPosY - Function to get the position of the Y-axis
    Public Function getPosY() As Double
        
        getPosY = fsinfo.MotorPos(motorY) - zeroOffset.Y
        
    End Function
    
    
    'setVelX - Function to set the velocity of the X-Axis
    Public Sub setVelX(vel As Double)
        
        'Create string for command
        Dim vel_Str As String
        
        'Generate velocity command
        vel_Str = genCmd(motorX, "V", Str(vel))
                
        'Send command to 6k
        c6k.Write vel_Str
    
    End Sub


    'setVelY - Function to set the velocity of the Y-Axis
    Public Sub setVelY(vel As Double)
        
        'Create string for command
        Dim vel_Str As String
        
        'Generate velocity command
        vel_Str = genCmd(motorY, "V", Str(vel))
                
        'Send command to 6k
        c6k.Write vel_Str
    
    End Sub
    
    Public Sub setVel(vel As Double)
    
        'Set velocity of both axes
        c6k.Write "V" & Str(vel) & "," & Str(vel) & ":"
    
    End Sub
    
    'update - Function to update fast-status from 6k
    Public Sub update()
    
        Dim temp() As Byte
        temp = c6k.FastStatus
        Call CopyMemory(fsinfo, temp(0), 280)

    End Sub
    
    Public Sub jogMode(myState As Boolean)
    
        'If jog is being enabled
        If myState Then
            'Set inputs to move motors
            c6k.Write "INFNC2-2J:INFNC1-2K:INFNC4-1J:INFNC3-1K:JOG11:"
        
        'If jog is being disabled
        Else
            c6k.Write "JOG00:INFNC2-A:INFNC1-A:INFNC4-A:INFNC3-A:"
        End If
    
    End Sub



'---------------------------- I/O -------------------------------------------------------------------
 
 
    'chkIn - Function to check the status of an input.
    Public Function chkIn(myInput As Integer) As Boolean
    
        'Convert the input from an integer identifier to a binary long
        Dim inBin As Long
        inBin = 2 ^ (myInput - 1)
        
        'Compare the binary input identifier to the c6k Input register. Return the result
        chkIn = fsinfo.ProgIn(1) And inBin
            
    End Function
    
    'chkOut - Function to check the status of an output.
    Public Function chkOut(myOutput As Integer) As Boolean
    
        'Convert the input from an integer identifier to a binary long
        Dim outBin As Long
        outBin = 2 ^ (myOutput - 1)
        
        'Compare the binary input identifier to the c6k Input register. Return the result
        chkOut = fsinfo.ProgOut(1) And inBin
    
    End Function
    
    
    'setIn - Function to enable or disable a virtual input.
    Public Sub setIn(myInput As Integer, myAction As Integer)
    
        Dim inputStr As String
        
        inputStr = "INEN." & Str(myInput)
        
        Select Case myAction
            Case 0
                inputStr = inputStr & "0"
            Case 1
                inputStr = inputStr & "1"
            Case 2
                inputStr = inputStr & "E"
        End Select
        
        inputStr = inputStr & ":"
        
        c6k.Write inputStr
    
    End Sub
    
    Public Sub virtualPress(myInput As Integer, myTime As Double)
    
        setIn myInput, 1
        
        c6k.Write "T" & Str(myTime) & ":"
        
        setIn myInput, 2
    
    End Sub
    
    
    'setOut - Sub to set the status of an output pin.
    Public Sub setOut(myOutput As Integer, myStatus As Boolean)
    
        Dim outStr As String
        
        outStr = "10OUT." & Str(myOutput)
        
        If myStatus Then outStr = outStr & "-1:" Else outStr = outStr & "-0:"
        
        c6k.Write outStr
    
    End Sub
    
    Private Sub startMov()
    
        movStarted = True
        
    End Sub
    
    Private Sub completeMov()
        
        setIn movComplete, 1
        movStarted = False
        
    End Sub
    
    Public Function checkMov() As Integer
    
        'If movStarted, then return 1
        If movStarted Then
            checkMov = 1
        'If not movStarted, return based on status ov movComplete
        Else
            'If movComplete raised, then return state 2 - move complete. Else return 0 - no move in progress
            If chkIn(movComplete) Then checkMov = 2 Else checkMov = 0
        End If

    End Function
    
    Public Function ackMov() As Boolean
        
        'By default, return false
        ackMov = False
        
        'If there is not currently a move in progress, exit function
        If Not chkIn(movComplete) Then
            Exit Function
        'If there is a currently queued move,
        Else
            setIn movComplete, 0
            ackMov = True
        End If
    
    End Function


'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=










'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===- Advanced Functions - Move to Position, etc. ---===---===---===---===---===---===---===---===

    Private Sub moveX(newPos As Double)
    
        'Create temporary location var
        Dim myNewPos As Double
        
        'Ensure that the motion does not exceed acceptable bounds
        myNewPos = newPos - zeroOffset.X
        If (myNewPos < 0) Then myNewPos = 0
    
        'Create temporary String
        Dim xStr As String
        Dim goX As String
        Dim waitX As String
                
        'Generate Position Command
        xStr = genCmd(motorX, "D", Str(myNewPos))
        
        'Generate Go Command
        goX = "GO" & genCmdBin(motorX, False, False, False) & ":"
        
        'Generate Wait Command
        waitX = "WAIT(MOV = b" & genCmdBin(motorX, True, True, True) & "):"
        
        'Send to 6k
        c6k.Write xStr & goX & waitX
    
    End Sub
    
    Private Sub moveY(newPos As Double)

        'Create temporary location var
        Dim myNewPos As Double
        
        'Ensure that the motion does not exceed acceptable bounds
        myNewPos = newPos - zeroOffset.Y
        If (myNewPos < 0) Then myNewPos = 0
        
        'Create temporary String
        Dim yStr As String
        Dim goY As String
        Dim waitY As String
                
        'Generate Position Command
        yStr = genCmd(motorX, "D", Str(myNewPos))
        
        'Generate Go Command
        goY = "GO" & genCmdBin(motorY, False, False, False) & ":"
        
        'Generate Wait Command
        waitY = "WAIT(MOV = b" & genCmdBin(motorY, True, True, True) & "):"
        
        'Send to 6k
        c6k.Write yStr & goY & waitY
    
    End Sub

    Private Sub moveXatVel(newPosX As Double, vel As Double)
    
        setVelX (vel)
        
        moveX (newPosX)
    
    End Sub
    
    Private Sub moveYatVel(newPosY As Double, vel As Double)
    
        setVelY (vel)
        
        moveY (newPosX)
    
    End Sub
    

    
    'moveLine - Moves multiple axes to a specified coordinate in a straight line.
    Private Sub moveLine(newPosX As Double, newPosY As Double)

        'Create temporary location var
        Dim myNewPosX As Double
        Dim myNewPosY As Double
        
        'Ensure that the motion does not exceed acceptable bounds
        myNewPosX = newPosX - zeroOffset.X
        If (myNewPosX < 0) Then myNewPosX = 0
        myNewPosY = newPosY - zeroOffset.Y
        If (myNewPosY < 0) Then myNewPosY = 0
        
        c6k.Write "D" & Str(myNewPosX) & "," & Str(myNewPosY) & ":GO11:WAIT(MOV = b00):"
        
    End Sub
    
    'moveLineAtVel - Moves multiple axes to a specified coordinate in a straight line.
    Private Sub moveLineAtVel(newPosX As Double, newPosY As Double, vel As Double)
    
        setVel (vel)
        
        moveLine newPosX, newPosY
        
    End Sub
    
    Public Sub moveTo(newPosX As Double, newPosY As Double, moveOption As moveType)
        
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMov Then Exit Sub
        
        'Raise flags to start motion
        startMov
            
            Select Case moveOption
                Case movNone
                    'Do Nothing
                Case movLine  'Move line
                    moveLine newPosX, newPosY
                Case movX
                    moveX newPosX
                Case movXY  'Move X first
                    moveX newPosX
                    moveY newPosY
                Case movY
                    moveY newPosY
                Case movYX  'Move Y first
                    moveY newPosY
                    moveX newPosX
                Case movY0XY  'Back-Over-Front
                    moveY -1000
                    moveX newPosX
                    moveY newPosY
            End Select
        
        'Clear moving flag and set completion flag
        completeMov
        
    End Sub
    
    'movePeck - Moves front then back, at separate velocities
    Public Sub movePeck(peckDeep As Double, peckBack As Double, peckVel As Double, returnVel As Double)
    
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMov Then Exit Sub
        
        'Raise flags to start motion
        startMov
        
            'Set advance velocity
            setVelY peckVel
            
            'Advance axis
            moveY peckDeep
            
            'Set retract velocity
            setVelY peckVel
            
            'Retract axis
            moveY (peckDeep - peckBack)
        
        'Clear moving flag and set completion flag
        completeMov
        
    End Sub
    
    Public Sub moveDelay(delayTime As Double)
    
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMov Then Exit Sub
        
        'Raise flags to start motion
        startMov
        
            'Delay the 6k by the indicated time
            c6k.Write "T" & Str(delayTime) & ":"
        
        'Clear moving flag and set completion flag
        completeMov
    
    End Sub

'---===---===---===---===- Advanced Functions - Move to Position, etc. ---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
