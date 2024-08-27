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
        mBoth
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
            
        'Erase predefined programs to ensure new programs have the correct identifiers
            myCns.write6k "ERASE:"
            
            'Ensure drives are disabled on startup, to allow scaling to be set properly
            myCns.write6k "@DRIVE0:"
        
        'Enable Absolute Mode
            myCns.write6k "@MA1:"
        
        'Set up axes
            'Set axes to behave as steppers
            myCns.write6k "@AXSDEF0:@DRFEN0:"
            
            'Set Scaling Factors to correct motors
            Dim strDriveScale As String
            Dim strResScale As String
            
            strDriveScale = Str(myDriveScale) & "," & Str(myDriveScale)
            
            If myMotorX = m1 Then
                strResScale = Str(myScaleX) & "," & Str(myScaleY)
            Else
                strResScale = Str(myScaleY) & "," & Str(myScaleX)
            End If
                
            myCns.write6k "DRES" & strDriveScale & ":SCLD" & strResScale & ":SCLV" & strResScale & ":SCLA" & strResScale & ":"
            
            'Enable Scaling
            myCns.write6k "SCALE1:@MC0:"
            
        'Set up IO
            
            YlimitInput = Y_Limit_Input_Pin
            
            'Set up Y-Axis negative limit switch
            'Dim strLimitSwitch As String
            
            'If myMotorY = m1 Then
                'strLimitSwitch = Str(YlimitInput) & "-1S:"
            'Else
                'strLimitSwitch = Str(YlimitInput) & "-2S:"
            'End If
            
            'myCns.write6k "1INFNC" & strLimitSwitch
            
            'Enable EOT Limit for negative Y-Axis
            'Dim strEOT As String
            
            'If myMotorY = m1 Then strEOT = "0,1" Else strEOT = "1,0"
            
            'myCns.write6k "@LH0:LH" & strEOT & ":LHAD100,100:LHLVL1:"
            
            
            'Set up estop button
            'ESTOPpin = ESTOP_Input_Pin
            
            'myCns.write6k "1INLVL" & Str(ESTOPpin) & "-1:1INFNC" & Str(ESTOPpin) & "-D:"
            
            
        'Initialize the 6k's motion parameters
        reset6k
        
        
        'Set initialized internal variables
        motorX = myMotorX
        motorY = myMotorY
        
        movComplete = movAcknInpt
        
        'Simulate move to delay UI initialization until the 6k initialization is complete
        movStarted = True
        setAckPin True
        
        
        
    End Sub
    
    'reset6k - resets 6k state to initialized state.
    Public Sub reset6k()
    
        'Initialize 6k motion parameters
        myCns.write6k "DRIVE11:"
        
        'Ensure Continuous Command Execution is Disabled
        myCns.write6k "COMEXC1:"
        
        'Set default velocity to 0, Set initial values for Acceleration, Tie Decell to accel
        myCns.write6k "V0,0:A10,10:AD0,0:"
        
        'Ensure acknowledge pin is set low for future motion
        setAckPin False
        
        'Set jog motion parameters
        'myCns.write6k "JOGA10,5:JOGAD300,150:JOGVH6,3:JOGVL30,15:"
    
    End Sub
    
    'Function specifically for operating the hydraulic clamp on the Drill-Tap
    'Public Sub initClamp(mySol As Integer, myPump As Integer, myOnInput As Integer, myOffInput As Integer)
    
        'myCns.write6k "DEL COFF:DEF COFF:"
            'setOut mySol, True
        'myCns.write6k "END:"
        
        'myCns.write6k "DEL CON:DEF CON:"
        
            'setOut mySol, True
            'setOut myPump, True
            'myCns.write6k "T1.9:"
            'setOut mySol, False
            'myCns.write6k "T.5:"
            'setOut myPump, False
    
        'myCns.write6k "END:"
        
        'myCns.write6k "1INFNC" & Str(myOffInput) & "-P1:"
        'myCns.write6k "1INFNC" & Str(myOnInput) & "-P2:"
    
    'End Sub
    
    
    'chkESTOP - Checks ESTOP pin and stops all motion if it is raised
    '   - INPUTS:
    '       - Active - Boolean that allows bypass of controller ESTOP shutdown. Used for maintenance purposes
    '   - OUTPUTS:
    '       - Boolean containing ESTOP State - True -> ESTOP Active, False -> ESTOP not Active
    '   - ACTIONS:
    '       - If Estop pin is set; stops all motion, disables all outputs, disables drives
    Public Function chkESTOP(Active As Boolean) As Boolean
    
        'Return false by default
        chkESTOP = False
    
        'If ESTOP pin is not raised, exit
        If Not chkIn(ESTOPpin) Then Exit Function
        
        If Active Then
            'Imediate stop command, imediate shut down all outputs, disable motors
            myCns.write6k "!AD100,100:!S:!OUTALL1,32,0:WAIT(MOV=b00):@DRIVE0:"
            
            'Clear move started flag
            movStarted = False
            
        End If
        
        chkESTOP = True
    
    End Function
    

'---------------------------- Internal Functions ----------------------------------------------------------------------
    
    
    Private Function coGlobal(motor As Integer, localPos As Double) As Double
    
        Dim globalOff As Double
    
        Select Case motor
            Case motorX
                globalOff = zeroOffset.X
            Case motorY
                globalOff = zeroOffset.Y
        End Select
        
        coGlobal = localPos + globalOff
                
    
    End Function
    
    Private Function coLocal(motor As Integer, globalPos As Double) As Double
    
        Dim localOff As Double
    
        Select Case motor
            Case motorX
                localOff = zeroOffset.X
            Case motorY
                localOff = zeroOffset.Y
        End Select
        
        coLocal = globalPos - localOff
    
    End Function
    
    Private Function genCmd(myCmd As String, motor As Integer, argument As String) As String

        Dim rtnCmd As String

        Select Case myCmd
            Case "GO"
                
                Dim posStr As String
                Dim goStr As String
                Dim wtStr
                
                Select Case motor
                    Case m1
                        posStr = argument
                        goStr = "10"
                        wtStr = "0X"
                    Case m2
                        posStr = "," & argument
                        goStr = "01"
                        wtStr = "X0"
                    Case mBoth
                        rtnCmd = "D" & argument & ":GO11:WAIT(MOV = b00):"
                        Exit Function
                End Select
                
                rtnCmd = "D" & posStr & ":" & myCmd & goStr & ":WAIT(MOV = b" & wtStr & "):"
                
            Case Else
        End Select
            
    End Function


'---------------------------- Generic Functions -----------------------------------------------------------------------
        
    
    'setZero - Function to set the zero location of the 6k
    Public Sub setZero(offsetX As Double, offsetY As Double)
    
        Dim cmdStr As String
        
        zeroOffset.X = offsetX
        zeroOffset.Y = offsetY
    
        'Sets zero to the offset location
        cmdStr = "PSET"
        
        If motorX = 1 Then
            cmdStr = cmdStr & Str(offsetX) & "," & Str(offsetY)
        Else
            cmdStr = cmdStr & Str(offsetY) & "," & Str(offsetX)
        End If
        
        cmdStr = cmdStr & ":"
        
        myCns.write6k cmdStr
    
    End Sub
    
    'setZero - Function to set the zero location of the 6k to the current location
    Public Sub setZeroHere()
    
        'Sets zero to the offset location
        myCns.write6k "PSET0,0:"
    
    End Sub
    
    
    'getPosX - Function to get the position of the X-axis
    Public Function getPosX() As Double
        
        Dim globalPos As Double
        Dim localPos As Double
        globalPos = (c6k.MotorPos(motorX) / myScaleX)
        
        localPos = globalPos - zeroOffset.X
        
        getPosX = localPos
        
    End Function
    
    'getPosY - Function to get the position of the Y-axis
    Public Function getPosY() As Double
        
        Dim globalPos As Double
        Dim localPos As Double
        globalPos = (c6k.MotorPos(motorY) / myScaleY)
        
        localPos = globalPos + zeroOffset.Y
        
        getPosY = localPos
        
    End Function
    
    
    'setVelX - Function to set the velocity of the X-Axis
    Public Sub setVelX(vel As Double)
        
        'Create string for command
        Dim vel_Str As String
        
        'Generate velocity command
        vel_Str = genCmd(motorX, "V", Str(vel))
                
        'Send command to 6k
        myCns.write6k vel_Str
    
    End Sub


    'setVelY - Function to set the velocity of the Y-Axis
    Public Sub setVelY(vel As Double)
        
        'Create string for command
        Dim vel_Str As String
        
        'Generate velocity command
        vel_Str = genCmd(motorY, "V", Str(vel))
                
        'Send command to 6k
        myCns.write6k vel_Str
    
    End Sub
    
    Public Sub setVel(vel As Double)
    
        'Set velocity of both axes
        myCns.write6k "V" & Str(vel) & "," & Str(vel) & ":"
    
    End Sub
    
    'update - Function to update fast-status from 6k
    Public Sub update()
    
        Dim temp() As Byte
        temp = c6k.FastStatus
        Call CopyMemory(c6k, temp(0), 280)

    End Sub
    
    Public Sub jogMode(myState As Boolean)
    
        'If jog is being enabled
        If myState Then
            'Set inputs to move motors
            myCns.write6k "INFNC2-2J:INFNC1-2K:INFNC4-1J:INFNC3-1K:JOG11:"
        
        'If jog is being disabled
        Else
            myCns.write6k "JOG00:INFNC2-A:INFNC1-A:INFNC4-A:INFNC3-A:"
        End If
    
    End Sub
    
    Public Sub setJogIn(myPin As Integer, myState As Boolean)
        
        'Verify that the pin needs to be changed - XOR returns false if both are the same -> Not xor
        Dim inputState As Long
        Dim pinBinary As Long
        
        Static myIn(4) As Boolean
        
        'Resolve the input state into a boolan for the desired pin
        'inputState = c6k.Inputs(0)
        'pinBinary = (2 ^ (myPin - 1))
        
        'inputState = CBool(inputState And pinBinary)
        
        inputState = myIn(myPin)
        
        'Compare the state of the pin to the desired state
        If (inputState Xor myState) Then
            
            
            Dim cmdStr As String
            
            'Initialize command - No number before INEN to indicate virtual input is in onboard I/O (Not expansion)
            cmdStr = "INEN." & myPin & "-"
            
            'Set state of pin to On or Off per myState input
            If myState Then cmdStr = cmdStr & "1:" Else cmdStr = cmdStr & "0:"
        
            'Execute command
            myCns.write6kNow cmdStr
            myIn(myPin) = myState
            
        End If
    
    End Sub



'---------------------------- I/O -------------------------------------------------------------------
 
 
    'chkIn - Function to check the status of an input.
    Public Function chkIn(myInput As Integer) As Boolean
    
        'Convert the input from an integer identifier to a binary long
        Dim inBin As Long
        Dim in6k As Long
        inBin = 2 ^ (myInput - 1)
        in6k = c6k.Inputs(1)
        
        'Compare the binary input identifier to the c6k Input register. Return the result
        chkIn = in6k And inBin
            
    End Function
    
    Public Function chkAckPin() As Boolean
    
        Dim in6k As Long
        in6k = c6k.Inputs(0)
    
        chkAckPin = in6k And (2 ^ (movComplete - 1))
    
    End Function
    
    Private Sub setAckPin(State As Boolean)
    
        Dim cmdStr As String
        cmdStr = "INEN." & Str(movComplete)
        
        If State Then cmdStr = cmdStr & "-1:" Else cmdStr = cmdStr & "-0:"
    
        myCns.write6k cmdStr
    
    End Sub
    
    'chkOut - Function to check the status of an output.
    Public Function chkOut(myOutput As Integer) As Boolean
    
        'Convert the input from an integer identifier to a binary long
        Dim outBin As Long
        outBin = 2 ^ (myOutput - 1)
        
        'Compare the binary input identifier to the c6k Input register. Return the result
        chkOut = c6k.Outputs(1) And outBin
    
    End Function
    
    
    'setIn - Function to enable or disable a virtual input.
    Public Sub setIn(myInput As Integer, myAction As Integer)
    
        Dim inputStr As String
        
        'Initialize the command & select pin
        inputStr = "!1INEN." & Str(myInput) & "-"
        
        'Set Output state
        Select Case myAction
            Case 0
                inputStr = inputStr & "0"
            Case 1
                inputStr = inputStr & "1"
            Case 2
                inputStr = inputStr & "E"
        End Select
        
        'Properly terminate the command
        inputStr = inputStr & ":"
        
        'Send the command to the 6k
        myCns.write6k inputStr
    
    End Sub
    
    Public Sub virtualPress(myInput As Integer, myTime As Double)
    
        setIn myInput, 1
        
        myCns.write6k "T" & Str(myTime) & ":"
        
        setIn myInput, 2
    
    End Sub
    
    
    'setOut - Sub to set the status of an output pin.
    Public Sub setOut(myOutput As Integer, myStatus As Boolean)
    
        Dim outStr As String
        
        outStr = "1OUT." & Str(myOutput)
        
        If myStatus Then outStr = outStr & "-1:" Else outStr = outStr & "-0:"
        
        myCns.write6k outStr
    
    End Sub
    
    
    
    ' checkMov - Checks the current status of the commands given
    '   - Access: Local - multiple functions use this one to determine if they can proceed to deliver commands to the 6k
    '   - Returns:
    '       - 0 - If there is no move in progress
    '       - 1 - If there is currently a move in progress that has not been completed
    '       - 2 - If there is a move that has been completed and needs to be acknowledged
    Private Function checkMov() As Integer
    
        'If movStarted, then return 1
        If movStarted Then
            checkMov = 1
        'If not movStarted, return based on status ov movComplete
        Else
            'If movComplete raised, then return state 2 - move complete. Else return 0 - no move in progress
            If chkAckPin Then checkMov = 2 Else checkMov = 0
        End If

    End Function
    
    ' ackMov - Acknowledges the completion of a move
    '   - Access: Public - This function is called by outside functions to determine if they may proceed. This code is non-blocking to allow looping checks
    '   - Returns:
    '       - True  - if the acknowledge pin has been identified AND reset
    '       - False - if:
    '                   - The acknowledge pin has not been raised
    '                   - The acknowledge pin has been raised and the reset command sent, but the pin has not been cleared
    Public Function ackMov() As Boolean
        
        'By default, return false
        ackMov = False
        
        'Prevent repeated acknowlege commands to the 6k by raising a static boolean when the command is sent to lower the pin
        Static cmdSent As Boolean
        
        'If there is not currently a move pending acknowledgement, exit function
        If Not chkAckPin Then
            Exit Function
        'If there is a currently completed move,
        Else
            'If the command has not been sent, send acknowledge command
            If Not cmdSent Then
            
                'Send command to clear pin
                setAckPin False
                'Set boolean to prevent repeat sends
                cmdSent = True
            
            'Once the command has been sent
            Else
                
                'Do not confirm acknowledgement until the pin has been cleared
                If Not chkAckPin Then
                    
                    'Return TRUE
                    ackMov = True
                    'Reset boolean for next loop
                    cmdSent = False
                    'Clear move started flag
                    movStarted = False
                    
                End If
            End If
        End If
    
    End Function


'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=










'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===- Advanced Functions - Move to Position, etc. ---===---===---===---===---===---===---===---===

    'Generic single-axis move command for any input axis.
    Private Sub move(motor As Integer, localPos As Double)
    
        Dim globalPos As Double
        Dim cmdStr As String
        
        'Convert local coordinates to global coordinates
        globalPos = coGlobal(motor, localPos)
        
        cmdStr = genCmd("GO", motor, CStr(globalPos))
        
        myCns.write6k cmdStr
    
    End Sub
    
    
    'moveLine - Moves multiple axes to a specified coordinate in a straight line.
    Private Sub moveLine(newPosX As Double, newPosY As Double)

        'Create temporary location var
        Dim globalPosX As Double
        Dim globalPosY As Double
        Dim argStr As String
        Dim cmdStr As String
        
        globalPosX = coGlobal(motorX, newPosX)
        globalPosY = coGlobal(motorY, newPosY)
        
        argStr = Format(globalPosX, "###0.000") & "," & Format(globalPosY, "###0.000")
        
        cmdStr = genCmd("GO", mBoth, argStr)
        
        myCns.write6k cmdStr
        
    End Sub
    
    Public Function moveTo(newPosX As Double, newPosY As Double, moveOption As moveType)
        
        moveTo = False
        
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMov Then Exit Function
            
        'Set move started flag
        movStarted = True
            
        Select Case moveOption
            Case movNone
                'Do Nothing
            Case movLine  'Move line
                moveLine newPosX, newPosY
            Case movX
                move motorX, newPosX
            Case movXY  'Move X first
                move motorX, newPosX
                move motorY, newPosY
            Case movY
                move motorY, newPosY
            Case movYX  'Move Y first
                move motorY, newPosY
                move motorX, newPosX
            Case movY0XY  'Back-Over-Front
                moveY 0
                move motorX, newPosX
                move motorY, newPosY
        End Select
    
        'set completion flag
        setAckPin True
        
        moveTo = True
        
    End Function
    
    'movePeck - Moves front then back, at separate velocities
    Public Function movePeck(peckDeep As Double, peckBack As Double, peckVel As Double, returnVel As Double)
    
        movePeck = False
    
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMov Then Exit Function
        
        'Set move started flag
        movStarted = True
        
        'Set advance velocity
        setVelY peckVel
        
        'Advance axis
        move motorY, peckDeep
        
        'Set retract velocity
        setVelY peckVel
        
        'Retract axis
        move motorY, (peckDeep - peckBack)
        
        'set completion flag
        setAckPin True
        
        movePeck = True
        
    End Function
    
    Public Function moveDelay(delayTime As Double)
    
        moveDelay = False
    
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMov Then Exit Function
        
        'Set move started flag
        movStarted = True
        
        'Delay the 6k by the indicated time
        myCns.write6k "T" & Str(delayTime) & ":"
        
        'set completion flag
        setAckPin True
    
        moveDelay = True
    
    End Function

'---===---===---===---===- Advanced Functions - Move to Position, etc. ---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
