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

    Const myDeadzoneX = -0.025   'Offsets for X and Y from 0,0 triggers to ensure a command to global (0,0) does not trigger limits
    Const myDeadzoneY = 0.025
    
    Const x_Max = 7.1
    Const y_Min = -2.9
    
    Const encCal = 225.6        'Calibration constant for the encoder, in pulses per inch. This ignores quatrature (i.e. a 56.4ppi encoder would be 225.6)
    Const c6kCycleMax = 50      'Maximum number of cycles in which the 6k tries to reach the encoder position

    '6k Variables:
    'VAR1 - Stores Encoder Position
    'VAR2 - Stores Desired Position
    'VAR3 - Stores Number of Failed Cycles
    'VAR4 - Stores Number of Failed Moves (i.e. max cycles reached)

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
        movHomeX
        movHomeY
    End Enum
      
'--------------------------- Class Variables --------------------------------------------------------------------------


    'Write-only Variables
    
    
    'Internal Variables
        'Variable defining the move complete acknowledgement virtual input pin. Used to stop windows code until mov end
        Private movComplete As Integer
        'Variable that indicates that a move has been initiated,but not completed
        Private moveInProgress As Boolean
        
        Private motorX As Integer
        Private motorY As Integer
        
        Private zeroOffset As myCoordinate
        
        Private YHomeIn As Integer      'Input Pin for Y-Travel Limit Switch
        Private XHomeIn As Integer
        Private ESTOPpin As Integer         'Input Pin for ESTOP
        

'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=














'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------
    Public Sub init6k(myMotorX As myMotors, myMotorY As myMotors, movAcknInpt As Integer, ESTOP_Input_Pin As Integer, X_Home_Pin As Integer, Y_Home_Pin As Integer)
        
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
            
            'Set initialized internal variables
            motorX = myMotorX
            motorY = myMotorY
            
            'Set axes to behave as steppers
            myCns.write6k "@AXSDEF0:@DRFEN0:"
            
            'Set Scaling Factors to correct motors
            Dim strDriveScale As String
            Dim strResScale As String
            
            strDriveScale = Str(myDriveScale) & "," & Str(myDriveScale)
            
            If motorX = m1 Then
                strResScale = Str(myScaleX) & "," & Str(myScaleY)
            Else
                strResScale = Str(myScaleY) & "," & Str(myScaleX)
            End If
                
            myCns.write6k "DRES" & strDriveScale & ":SCLD" & strResScale & ":SCLV" & strResScale & ":SCLA" & strResScale & ":"
            
            'Enable Scaling
            myCns.write6k "SCALE1:@MC0:"
            
        'Set up Home Limits & initialize EOT Limits
            
            movComplete = movAcknInpt
            
            YHomeIn = Y_Home_Pin
            XHomeIn = X_Home_Pin
    
            myCns.write6k "@HOMV1:@HOMBAC1:@HOMVF0.5:"
            
        'Set up estop button
            ESTOPpin = ESTOP_Input_Pin
            
            myCns.write6k "1INFNC" & Str(ESTOPpin) & "-D:"
            
            
        'Initialize the 6k's motion parameters
        reset6k
        
        'Simulate move to delay UI initialization until the 6k initialization is complete
        moveInProgress = True
        setAckPin True
        
    End Sub
    
    'reset6k - resets 6k state to initialized state.
    Public Sub reset6k()
    
        'Initialize 6k motion parameters
        myCns.write6k "DRIVE11:"
        
        'Ensure Continuous Command Execution is Disabled
        myCns.write6k "COMEXC1:"
        
        'Set default velocity to 0, Set initial values for Acceleration, Tie Decell to accel
        myCns.write6k "@D0:@V0:@A2:@AD0:"
        
        'Ensure acknowledge pin is set low for future motion
        setAckPin False
        
        moveInProgress = False
        
        'Set jog motion parameters
        'myCns.write6k "JOGA10,5:JOGAD300,150:JOGVH6,3:JOGVL30,15:"
    
    End Sub
    
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
        
        If Active Then callESTOPnow
        
        chkESTOP = True
    
    End Function
    
    Public Sub callESTOPnow()
    
        'Imediate stop command, imediate shut down all outputs, disable motors
        myCns.write6k "!1OUTALL1,8,0:!AD100,100:!S:WAIT(MOV=b00):@DRIVE0:"
        
        'Set move in progress flag to cause move commands to abort
        moveInProgress = True
        
    End Sub
    

'---------------------------- Internal Functions ----------------------------------------------------------------------
    
    
    Private Function coGlobal(motor As Integer, localPos As Double) As Double
    
        Dim globalOff As Double
        Dim globalPos As Double
    
        Select Case motor
            Case motorX
                globalOff = zeroOffset.X
            Case motorY
                globalOff = zeroOffset.Y
        End Select
        
        globalPos = localPos + globalOff
                
        Select Case motor
            Case motorX
                If globalPos < 0 Then globalPos = 0
                If globalPos > x_Max Then globalPos = x_Max
            Case motorY
                If globalPos > 0 Then globalPos = 0
                If globalPos < y_Min Then globalPos = y_Min
        End Select
        
        coGlobal = globalPos
    
    End Function
    
    Private Function coLocal(motor As Integer, globalPos As Double) As Double
    
        Dim localOff As Double
        Dim localPos As Double
    
        Select Case motor
            Case motorX
                localOff = zeroOffset.X
            Case motorY
                localOff = zeroOffset.Y
        End Select
        
        localPos = globalPos - localOff
        
        coLocal = localPos
    
    End Function
    
    Private Function genCmd(myCmd As String, motor As Integer, argument As String, Optional isEncoder As Boolean = False) As String

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
                
                If Not isEncoder Then
                    'Simply tell the controller to go to the position and wait for command execution to resume after the position is reached
                    rtnCmd = "D" & posStr & ":GO" & goStr & ":WAIT(MOV = b" & wtStr & "):"
                Else
                    'Define the desired position and tell the controller to go
                    rtnCmd = "D" & posStr & ":GO" & goStr & ":"
                    'Repeat the following check
                    rtnCmd = rtnCmd & "REPEAT:"
                        'Wait until motion has stopped
                        rtnCmd = rtnCmd & "WAIT(MOV = b" & wtStr & "):"
                        'Store the encoder position (corrected to reference inches) and the desired position as variables for comparison
                        rtnCmd = rtnCmd & "VAR1=" & CStr(motor) & "PE:VAR1=VAR1/" & CStr(encCal) & ":VAR2=" & CStr(motor) & "D:"
                        'Compare the encoder and commanded positions. If they differ
                        rtnCmd = rtnCmd & "IF(VAR1<>VAR2):"
                            'correct the stepper position in the controller
                            rtnCmd = rtnCmd & "PSET(VAR1):"
                            'If the 6k has not run through this cycle more times than the cycle max
                            rtnCmd = rtnCmd & "IF(VAR3 < " & CStr(c6kCycleMax) & "):"
                                'Repeat the Go command and increment the cycle counter
                                rtnCmd = rtnCmd & "GO" & goStr & ":VAR3 = VAR3+1:"
                            rtnCmd = rtnCmd & "ELSE:"
                                'Exit the repeat cycle and increment the failed cycle counter
                                rtnCmd = rtnCmd & "VAR2=VAR1:VAR4 = VAR4+1:VAR3=0:"
                            rtnCmd = rtnCmd & "NIF:"
                        rtnCmd = rtnCmd & "ELSE:"
                            rtnCmd = rtnCmd & "VAR3=0"
                        rtnCmd = rtnCmd & "NIF:"
                    'Repeat the loop until the motor reaches the desired position or the cycle counter has been exceeded
                    rtnCmd = rtnCmd & "UNTIL(VAR1=VAR2):"
                End If
                
            Case "V", "D", "A"
            
                Dim axisStr As String
                
                Select Case motor
                    Case m1
                        axisStr = argument
                    Case m2
                        axisStr = "," & argument
                    Case mBoth
                        axisStr = argument & "," & argument
                End Select
                
                rtnCmd = myCmd & axisStr & ":"
                
            Case Else
        End Select
            
        genCmd = rtnCmd
            
    End Function


'---------------------------- Generic Functions -----------------------------------------------------------------------
        
    
    'setZero - Function to set the zero location of the 6k
    Public Sub setZero(offsetX As Double, offsetY As Double)
        
        zeroOffset.X = offsetX
        zeroOffset.Y = offsetY
    
    End Sub
    
    
    'getPos - Function to get the position of the specified axis
    Public Function getCo(motor As Integer, Optional isEncoder As Boolean = False, Optional coGlobal As Boolean = False) As Double
        
        Dim globalPos As Double
        Dim localPos As Double
        Dim returnPos As Double
        
        Select Case motor
            Case motorX
                globalPos = (c6k.MotorPos(motorX) / myScaleX)
            Case motorY
                If isEncoder Then
                    globalPos = (c6k.EncoderPos(motorY) / encCal)
                Else
                    globalPos = (c6k.MotorPos(motorY) / myScaleY)
                End If
        End Select
        
        If Not coGlobal Then returnPos = coLocal(motor, globalPos) Else returnPos = globalPos
        
        getCo = returnPos
        
    End Function
    
    'setVel - Function to set the velocity of the specified axis
    Public Sub setVel(motor As Integer, vel As Double)
        'Create string for command
        Dim vel_Str As String
        
        'Generate velocity command
        vel_Str = genCmd("V", motor, Str(vel))
                
        'Send command to 6k
        myCns.write6k vel_Str
    
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
    
    Private Sub setAckPin(state As Boolean)
    
        Dim cmdStr As String
        cmdStr = "WAIT(MOV=b00):INEN." & Str(movComplete)
        
        If state Then cmdStr = cmdStr & "-1:" Else cmdStr = cmdStr & "-0:"
    
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


'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=










'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===- Advanced Functions - Move to Position, etc. ---===---===---===---===---===---===---===---===
    
    Public Sub setClamp(Pump_Pin As Integer, Sol_Pin As Integer, clampSet As Boolean)
    
        'If the request is to release the clamp
        If Not clampSet Then
        
            'Open the clamp solenoid valve
            setOut CStr(Sol_Pin), True
        
        'Otherwise, close the clamp
        Else
            
            'Start by opening the clamp solenoid
            setOut CStr(Sol_Pin), True
            'Wait for 0.3 seconds
            myCns.write6k "T.3:"
            'Start the pump motor
            setOut CStr(Pump_Pin), True
            'Wait 1.9 Seconds
            myCns.write6k "T1:"
            'Close Solenoid
            setOut CStr(Sol_Pin), False
            'Wait half a second
            myCns.write6k "T.5:"
            'Then stop the pump
            setOut CStr(Pump_Pin), False
        
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
    
    
    
    
    ' checkMove - Simplifies the acknowledge move response down to a simple True-False for continued motion
    '   - Returns: False if the program may proceed to a new move command; True if a move is in progress.
    Public Function checkMove() As Boolean
    
        Dim response As Integer
        
        response = ackMov
        
        Select Case response
            Case 0, 3
                checkMove = False
            Case Else
                checkMove = True
        End Select
    
    End Function
    
    ' ackMov - Acknowledges the completion of a move
    '   - Returns:
    '       - 0 - If there is no move pending
    '       - 1 - If there is a move in progress, but the acknowledge pin has not been raised
    '       - 2 - If the acknowledge pin has been identified and the command to reset it has been issued, but not yet executed
    '       - 3 - If the acknowledge pin has been set and cleared such that an additional move may occur
    Public Function ackMov() As Integer
    
        'Prevent repeated acknowlege commands to the 6k by raising a static boolean when the command is sent to lower the pin
        Static CmdSent As Boolean
        
        If Not moveInProgress Then
            ackMov = 0
        Else
            ackMov = 1
            
            If chkAckPin And Not CmdSent Then
                
                'Send command to reset acnkowledge pin
                setAckPin False
                CmdSent = True
                
                ackMov = 2
                
            ElseIf Not chkAckPin And CmdSent Then
                
                'Clear variables and return state
                CmdSent = False
                moveInProgress = False
                
                ackMov = 3
            
            End If
            
        End If
    
    End Function
    
    'Generic single-axis move command for any input axis.
    Private Sub move(motor As Integer, localPos As Double, Optional isEncoder As Boolean = False)
    
        'Create temporary location var
        Dim globalPos As Double
        Dim cmdStr As String
        
        'Ensure that the motion does not exceed acceptable bounds
        globalPos = coGlobal(motor, localPos)
        
        'Generate the initial GO command
        cmdStr = genCmd("GO", motor, Format(globalPos, "###0.000"), isEncoder)
        
        'Send to 6k
        myCns.write6k cmdStr
    
    End Sub
   
    'moveLine - Moves multiple axes to a specified coordinate in a straight line.
    Private Sub moveLine(newPosX As Double, newPosY As Double)

        Dim cmdStr As String
        Dim newPos As myCoordinate
        
        newPos.X = coGlobal(motorX, newPosX)
        newPos.Y = coGlobal(motorY, newPosY)
        
        cmdStr = genCmd("GO", mBoth, (CStr(newPos.X) & "," & CStr(newPos.Y)))
        
        myCns.write6k cmdStr
        
    End Sub
    
    Public Sub moveStop()
        
        'Stop 6k
        myCns.write6k "!S:"
        
        'Clear all movement variables
        moveInProgress = False
        setAckPin False
        
    End Sub
    
    Private Sub moveHome(motor As Integer, Direction As Boolean)
    
        'Ensure limits are disabled such that 6k can use the limits as home limits
        myCns.write6k "@LH0:LIMLVL0,0,0,0,0,0:"
    
        myCns.write6k "1INLVL." & CStr(XHomeIn) & "-1:1INLVL." & CStr(YHomeIn) & "-1:"
    
        Dim cmdStr As String
    
        'Set limit switches to appropriate axis
        Dim strLimX As String
        Dim strLimY As String
        
        If motorX = m1 Then
            strLimX = Str(XHomeIn) & "-1T:"
            strLimY = Str(YHomeIn) & "-2T:"
        Else
            strLimX = Str(XHomeIn) & "-2T:"
            strLimY = Str(YHomeIn) & "-1T:"
        End If
        
        'Set input function for those inputs as home-limit inputs
        myCns.write6k "1INFNC" & strLimX & "1INFNC" & strLimY
        
        Dim deadzone1 As Double
        Dim deadzone2 As Double
        If motorX = m1 Then
            deadzone1 = myDeadzoneX
            deadzone2 = myDeadzoneY
        Else
            deadzone2 = myDeadzoneX
            deadzone1 = myDeadzoneY
        End If
        
        Select Case motor
            Case 1
                'Determine home direction - True is Positive, False is Negative
                If Direction Then
                    'Start with Home motion paramters:
                    'Home final direction should be the same as the initial direction
                    cmdStr = cmdStr & "HOMDF0X:"
                    'Home should end on the edge closest to the initial direction (therefore opposite edge - home negative, positive edge)
                    cmdStr = cmdStr & "HOMEDG1X:"
                    'Begin Home Motion
                    cmdStr = cmdStr & "HOM0X:"
                Else
                    cmdStr = cmdStr & "HOMDF1X:"
                    cmdStr = cmdStr & "HOMEDG0X:"
                    cmdStr = cmdStr & "HOM1X:"
                End If
                
                cmdStr = cmdStr & "WAIT(MOV = b00):PSET" & CStr(deadzone1) & ":PESET," & CStr(deadzone2 * encCal) & ":"
            Case 2
                If Direction Then
                    cmdStr = cmdStr & "HOMDFX0:"
                    cmdStr = cmdStr & "HOMEDGX1:"
                    cmdStr = cmdStr & "HOMX0:"
                Else
                    cmdStr = cmdStr & "HOMDFX1:"
                    cmdStr = cmdStr & "HOMEDGX0:"
                    cmdStr = cmdStr & "HOMX1:"
                End If
                
                cmdStr = cmdStr & "WAIT(MOV = b00):PSET," & CStr(deadzone2) & ":PESET," & CStr(deadzone2 * encCal) & ":"
        End Select
    
        'Initiate Home Motion
        myCns.write6k cmdStr
    
        Dim strLimAxes As String
        
        'Convert the Home limts to EOT limits
        If motorX = m1 Then
            strLimX = Str(XHomeIn) & "-1S:"
            strLimY = Str(YHomeIn) & "-2R:"
            strLimAxes = "LH1,2:"
        Else
            strLimX = Str(XHomeIn) & "-2S:"
            strLimY = Str(YHomeIn) & "-1R:"
            strLimAxes = "LH2,1:LIMLVL1,1,1,1,1,1:"
        End If
        
        myCns.write6k "1INLVL." & CStr(XHomeIn) & "-0:1INLVL." & CStr(YHomeIn) & "-0:"
        
        'Set input function for those inputs as home-limit inputs
        myCns.write6k "1INFNC" & strLimX & "1INFNC" & strLimY & strLimAxes
        
    End Sub
    
    Public Function moveTo(newPosX As Double, newPosY As Double, moveOption As moveType, Optional isEncoder As Boolean = False) As Boolean
        
        moveTo = False
        
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMove Then Exit Function
            
        'Set move started flag
        moveInProgress = True
            
        Select Case moveOption
            Case movNone
                'Do Nothing
            Case movLine  'Move line
                moveLine newPosX, newPosY
            Case movX
                move motorX, newPosX
            Case movXY  'Move X first
                move motorX, newPosX
                move motorY, newPosY, isEncoder
            Case movY
                move motorY, newPosY, isEncoder
            Case movYX  'Move Y first
                move motorY, newPosY, isEncoder
                move motorX, newPosX
            Case movY0XY  'Back-Over-Front
                move motorY, 10, isEncoder
                move motorX, newPosX
                move motorY, newPosY, isEncoder
            Case movHomeX
                moveHome motorX, False
            Case movHomeY
                moveHome motorY, True
        End Select
    
        'set completion flag
        setAckPin True
        
        moveTo = True
        
    End Function
    
    'movePeck - Moves front then back, at separate velocities
    Public Function movePeck(peckDeep As Double, peckBack As Double, peckVel As Double, returnVel As Double) As Boolean
    
        movePeck = False
    
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMove Then Exit Function
        
        'Set move started flag
        moveInProgress = True
        
        'Set advance velocity
        setVel motorY, peckVel
        
        'Advance axis
        move motorY, peckDeep
        
        'Set retract velocity
        setVel motorY, returnVel
        
        'Retract axis
        move motorY, (peckDeep + peckBack)
        
        'set completion flag
        setAckPin True
        
        movePeck = True
        
    End Function
    
    Public Function moveDelay(delayTime As Double) As Boolean
    
        moveDelay = False
    
        'If there is either a currently pending movement or a completed-not-acknowledged movement, abort
        If checkMove Then Exit Function
        
        'Set move started flag
        moveInProgress = True
        
        'Delay the 6k by the indicated time
        myCns.write6k "T" & Str(delayTime) & ":"
        
        setAckPin True
        
        moveDelay = True
    
    End Function

'---===---===---===---===- Advanced Functions - Move to Position, etc. ---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
