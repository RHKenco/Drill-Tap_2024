VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDrillTap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- Drill/Tap Class -----
'|
'|
'|
'|====================================================================================================================|

Option Explicit


'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===

'--------------------------- Class Enumerations & Constants -----------------------------------------------------------

    'Onboard Inputs - Virtual inputs for Jog Mode
    Const In_Joy_X1 = 1
    Const In_Joy_X2 = 2
    Const In_Joy_Y1 = 3
    Const In_Joy_Y2 = 4
    
    Const In_MoveAck = 5
    
    'Expansion Outputs - Pins 1-8
    Const Out_Clamp_Sol = 2
    Const Out_Clamp_Pump = 3
    Const Out_Drill = 4
    Const Out_Tap = 5
    Const Out_Cool = 6
    
    'Expansion Inputs - Pins 9-16
    Const In_Clamp_Lim = 11
    Const In_HomX = 12
    Const In_HomY = 13
    Const In_Clamp_On = 14
    Const In_ESTOP = 15
    Const In_Clamp_Off = 16
    
    'Path to stored data pertaining to Offsets
    Const dataSavePath = "C:\Drill-Tap_2024\savedData\offsets.txt"
    
    Const Axis_X = 2
    Const Axis_Y = 1

    Public Enum dtCo
        dNULL
        dX
        dY
    End Enum
    
    Public Enum myTool
        tDrill
        tTap
    End Enum

    Public Enum joyDirection
        joyNeutral
        joyUp
        joyUpRight
        joyRight
        joyDownRight
        joyDown
        joyDownLeft
        joyLeft
        joyUpLeft
    End Enum
     
    'Integer and Enum to store position of the machine
    Public Enum atPosition
        atIdle
        atUnk
        atDrill
        atTap
    End Enum
    
    Public Enum maintOption
        mActiveOn
        mActiveOff
        mEnable
        mDisable
        mToggle
    End Enum
    
    Private myPosition As atPosition
    Private myDestination As atPosition
    Private ESTOP_SET As Boolean
     
'--------------------------- Class Variables --------------------------------------------------------------------------

    'Write-only Variables
    Private V_Jog As Double
    Private V_Drill As Double
    Private V_Tap As Double
    
    'Boolean flag for joystick state - prevent certain actions while joystick is on
    Private joyOn As Boolean
    Private zeroSet As Boolean
    
    
    'Object to store the last DRO reading when a set-zero occurred.
    Private PosLastZero As myCoordinate
    
    'Objects to store the offsets for the drill & tap locations.
    Private stdOffset As Double         'Offset from which the origin point will be determined relative to the surface contact
    Private zeroOffset As myCoordinate  'Location of drill relative to hard stops of machine
    Private goTap As myCoordinate       'Location of Tap relative to Drill
    Private goDrill As myCoordinate
    
    Private maintIO(17) As Boolean
    
    Private myXY As New Class6kXY
    

'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=








'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------
    Public Sub initDrillTap()
        
        'Initialize class variables
        zeroSet = False
        joyOn = False
        ESTOP_SET = False
        
            
        'Create 6k movement control object
        myXY.init6k Axis_X, Axis_Y, In_MoveAck, In_ESTOP, In_HomX, In_HomY    'X-axis is motor 2, Y-axis is motor 1, acknowledge input is pin 32
        
        'Pull saved offsets from data file; will set defaults if file does not exist
        loadOffsets
        myXY.setZero zeroOffset.X, zeroOffset.Y
        
        'Pull velocity defaults from the UI
        setVelDefaults frmUI.txtSpeeds(0), frmUI.txtSpeeds(1), frmUI.txtSpeeds(2)
            
            
    End Sub
    
    Private Function loadData() As Boolean
        
        If Dir(dataSavePath) <> "" Then
            
            'Load previous data
            Open dataSavePath For Input As #1
            Input #1, zeroOffset.X, zeroOffset.Y, goTap.X, goTap.Y
            Close #1
            
            loadData = True
        Else
            loadData = False
        End If
    
    End Function
    
    Private Sub loadOffsets()

        If Not loadData Then
            
            stdOffset = 0.375
            zeroOffset.X = 0.875
            zeroOffset.Y = -0.125
            goTap.X = 5.625
            goTap.Y = -1.375
            
        End If
        
        goDrill.X = 0
        goDrill.Y = 0
        
    End Sub
    
    Private Function saveData() As Boolean
    
        'Save to file
        Open dataSavePath For Output As #1
        Write #1, zeroOffset.X, zeroOffset.Y, goTap.X, goTap.Y
        Close #1
    
    End Function
    
    Public Function setHome()
    
        'Used to set ONLY the global coordinate offsets to calibrate the local coordinate system.
    
        'Get current position, and store it as the last zero location
        PosLastZero.X = getCoords(dX)
        PosLastZero.Y = getCoords(dY)
        
        'Invert the coordinates so the last zero is stored relative to the current zero (not the other way around)
        PosLastZero.X = -PosLastZero.X
        PosLastZero.Y = -PosLastZero.Y
        
        'Get the current position in global coordinates to update the zeroOffset
        zeroOffset.X = myXY.getCo(Axis_X, True)
        zeroOffset.Y = myXY.getCo(Axis_Y, True)
        
        'Set zero using the zeroOffset
        myXY.setZero zeroOffset.X, zeroOffset.Y
        
        'Save the new offsets
        saveData
        
        'Indicate that the zero is properly set
        zeroSet = True
        
    End Function
    
    Public Sub setToolOffset()
    
        'Used to set ONLY the offset between the drill and the tap.
    
        'Update the offsets to the current local position
        goTap.X = getCoords(dX)
        goTap.Y = getCoords(dY)
        
        'Save the new offsets
        saveData
        
    End Sub

'---------------------------- Read-Only Set & Returns -----------------------------------------------------------------
        
        
    Public Sub update()
        myXY.update
    End Sub
    
        
        
    ''' getPos - Gets the current position of both axes
    Public Function getCoords(coordinate As dtCo) As Double
        
        Select Case coordinate
            Case cX
                getCoords = myXY.getCo(Axis_X)
            Case cY
                getCoords = myXY.getCo(Axis_Y)
            Case Else
        End Select
        
    End Function
    
    Public Function getPos() As Integer
        
        getPos = myPosition
    
    End Function
    
    Private Sub setPos(position As atPosition)
    
        Select Case position
            Case atDrill
                myPosition = atDrill
            Case atTap
                myPosition = atTap
            Case Else
                myPosition = atUnk
        End Select
    
    End Sub
    
    Public Sub clrPos()
    
        myPosition = atUnk
    
    End Sub
    
    ''' getLastZero - Gets the last zero-set for the indicated motor. Used to determine setzero variance
    Public Function getLastZero(coordinate As dtCo) As Double
    
        Select Case coordinate
            Case cX
                getLastZero = PosLastZero.X
            Case cY
                getLastZero = PosLastZero.Y
            Case Else
        End Select
    
    End Function
    
    ''' setVelDefaults - Sets the internal velocity variables for drill/tap movement
    Public Sub setVelDefaults(drillVel As String, tapVel As String, jogVel As String)
    
        V_Jog = CDbl(jogVel)
        V_Drill = CDbl(drillVel)
        V_Tap = CDbl(tapVel)
    
    End Sub
    
    Public Function isHomeSet() As Boolean
        isHomeSet = zeroSet
    End Function
    
    Public Function isJoyOn() As Boolean
        isJoyOn = joyOn
    End Function

    'checkESTOP - Checks ESTOP pin and stops all motion if it is raised
    '   - INPUTS:
    '       - Active - Boolean that allows bypass of controller ESTOP shutdown. Used for maintenance purposes
    '   - OUTPUTS:
    '       - Boolean containing ESTOP State - True -> ESTOP Active, False -> ESTOP not Active
    '   - ACTIONS:
    '       - If ESTOP is set; sets class flag, ignores pin until cleared, sets FSM Estop state
    Public Function checkESTOP(Active As Boolean, Optional IsFSMcall As Boolean) As Boolean
                       
        'If estop flag set, pin data irrelevant. If it is not set, check estop pin
        If Not ESTOP_SET Then ESTOP_SET = myXY.chkESTOP(Active)
        
        If ESTOP_SET Then If Not IsFSMcall Then myFSM.setESTOP
        
        'Return value of ESTOP SET
        checkESTOP = ESTOP_SET
        
    End Function
    
    Public Function ESTOPset() As Boolean
        ESTOPset = ESTOP_SET
    End Function
    
    Public Function resetESTOP()
        myXY.reset6k
        ESTOP_SET = False
    End Function


    Public Sub maintPassive()
    
        Dim i As Integer
        Dim check As Boolean
        
        For i = 0 To 15
            'For the first 8 pins, check as outputs. For the next 8 pins, check as inputs
            If i < 9 Then check = myXY.chkOut(i + 1) Else check = myXY.chkIn(i + 1)
            
            'If Active Maintenance mode and manually activated I/O, color red
            If ((check And frmMaintenance.frameMaintIO.Enabled) And maintIO(i + 1)) Then
                frmMaintenance.cmdMaintIO(i).BackColor = &H9191EA
            
            'If I/O is enabled, color green
            ElseIf check Then
                frmMaintenance.cmdMaintIO(i).BackColor = &H91EA9A
            
            'If I/O is disabled, color Blue
            Else
                frmMaintenance.cmdMaintIO(i).BackColor = &HEAC091
            End If
            
            
        Next i
    
    End Sub
    
    Public Sub maintActive(inputPin As Integer, myAction As maintOption)
        
        Dim enable As Boolean
        
        'Determine what action should be taken
        Select Case myAction
        
            Case mActiveOff
                inputPin = 0
                enable = False
            Case mActiveOn
                inputPin = 0
                enable = True
            Case mEnable
                enable = True
                maintIO(inputPin) = True
            Case mDisable
                enable = False
                maintIO(inputPin) = True
            Case mToggle
                enable = Not maintIO(inputPin)
                maintIO(inputPin) = enable
            Case Else
        End Select
        
        
        'If the input is 0, this is a call to enable or disable active maintenance mode
        If inputPin = 0 Then
        
            Dim i As Integer
            
            'Clear all I/O Booleans
            For i = 0 To 16
                maintIO(i) = False
            Next i
            
            'Enable or disable the maintenance frame according to the enable input
            If enable Then
                frmMaintenance.frameMaintIO.Enabled = True
                myFSM.setMaintenance True
                myUI.updStatus uiMaintEnabled, ""
            Else
                frmMaintenance.frameMaintIO.Enabled = False
                myFSM.setMaintenance False
                myUI.updStatus uiStandby, ""
            End If
                
            
        'Any other input, assuming the mode is active, should be set and cleared manually
        Else
            
            'Verify that maintenance mode is enabled
            If frmMaintenance.frameMaintIO.Enabled Then
                
                'The boolean array should be updated accordingly
                maintIO(inputPin) = enable
                
                
                'If 1-8, output pins. If 9-16, input pins.
                If inputPin < 8 Then
                    myXY.setOut inputPin, enable
                Else
                    'Input pins are switched between virtual on and enable, not on/off (true false)
                    If enable Then
                        myXY.setIn inputPin, 1
                    Else
                        myXY.setIn inputPin, 2
                    End If
                End If
            End If
        End If
    
    End Sub


'---------------------------- Output Enable/Disable -------------------------------------------------------------------
 
    Private Sub setDrillTap(drillState As Boolean, tapState As Boolean, coolantState As Boolean)
    
        myXY.setOut Out_Drill, drillState
        
        myXY.setOut Out_Tap, tapState
        
        myXY.setOut Out_Cool, coolantState
        
    
    End Sub
 
    Public Function setClamp(clampState As Boolean) As Boolean
    
        Static CmdSent As Boolean
        Static sendCounter As Integer
        Dim commandMaximum As Integer
        commandMaximum = 40    'Cycles at 50ms = 20 hz therefore 40 cycles = 2 second
        
        'Increment counter each loop after the command is sent; resend if the command does not get executed after 2 seconds
        If CmdSent Then
            sendCounter = sendCounter + 1
            If sendCounter > commandMaximum Then
                sendCounter = 0
                CmdSent = False
            End If
        End If
        
        'If the clamp is requested to close but is already closed, (and a command is not already in progress), open clamp
        If (clampState And checkClamp) And Not CmdSent Then clampState = False
            
        'If the clamp is requested to open
        If Not clampState Then
            'If it is not already open
            If checkClamp Then
                'Send the command if it has not already been given
                If Not CmdSent Then
                    myXY.setClamp Out_Clamp_Pump, Out_Clamp_Sol, False
                    CmdSent = True
                    sendCounter = 0
                End If
            'If the clamp is open, exit function and return true
            Else
                CmdSent = False
                setClamp = True
            End If
        'If the clamp is requested to be closed
        Else
            'If the clamp is open
            If Not checkClamp Then
                'Send the close command if it has not already been given
                If Not CmdSent Then
                    myXY.setClamp Out_Clamp_Pump, Out_Clamp_Sol, True
                    CmdSent = True
                    sendCounter = 0
                End If
            'If the clamp is closed, exit function and return true
            Else
                CmdSent = False
                setClamp = True
            End If
        End If
        
    End Function
    
    Public Function checkClamp() As Boolean
    
        'Check clamp pin. Invert due to pin being normally closed
        checkClamp = Not myXY.chkIn(In_Clamp_Lim)
    
    End Function
    
    'checkInputs - Checks the state of all user inputs each FSM loop, and acts accordingly.
    Public Function checkInputs(passActive As Boolean) As Boolean
    
        Static inputState As Integer
    
        'Return false by default
        checkInputs = False
        
        'Check clamp - if the clamp is not set, see if a pass is in progress
        If Not checkClamp Then
        
            'If pass is in progress, then set estop
            If passActive Then
                myFSM.setESTOP
                Exit Function
            End If
            
        End If
        
        'If the pass is not in progress and the clamp-on button is pressed, engage the clamp
        If (myXY.chkIn(In_Clamp_On) Or (inputState = 1)) And Not passActive Then
            
            'Verify that clamp-off is not in progress
            If inputState <> 2 Then
                'Set InputState variable to indicate that a clamp-on operation is in progress
                If inputState = 0 Then inputState = 1
                
                'Call the set-Clamp function. This will return true when the clamp is confirmed to have been set, clearing the input state
                If setClamp(True) Then inputState = 0
            
            End If
        'If the pass is not in progress and the clamp-on button is pressed, engage the clamp
        ElseIf (myXY.chkIn(In_Clamp_Off) Or (inputState = 2)) And Not passActive Then
            
            'Verify that clamp-off is not in progress
            If inputState <> 1 Then
                'Set InputState variable to indicate that a clamp-on operation is in progress
                If inputState = 0 Then inputState = 2
                
                'Call the set-Clamp function. This will return true when the clamp is confirmed to have been Released, clearing the input state
                If setClamp(False) Then inputState = 0
            
            End If
        End If
        
        checkInputs = True
    
    End Function
    
    
'---------------------------- Joystick --------------------------------------------------------------------------------

    Public Sub joyState(myState As Boolean)
    
        If (myState And (Not joyOn)) Then
        
            joyEnable
        
        ElseIf ((Not myState) And joyOn) Then
        
            joyDisable
        
        Else
            
            'Current state and desired state are the same
            Exit Sub
        
        End If
    
        joyOn = myState
    
    End Sub
    
    Private Sub joyEnable()
    
        'Enable jog mode
        myXY.jogMode True
    
    End Sub
    
    Private Sub joyDisable()

        'Ensure all virtual inputs are disabled
        joy joyNeutral
        
        'Disable jog mode
        myXY.jogMode False

    End Sub
    
    Public Sub joy(Direction As joyDirection)
    
        If Not joyOn Then Exit Sub
    
        Select Case Direction
            Case joyNeutral  'No Movement
                myXY.setJogIn In_Joy_X1, False
                myXY.setJogIn In_Joy_X2, False
                myXY.setJogIn In_Joy_Y1, False
                myXY.setJogIn In_Joy_Y2, False
            Case joyUp  'Move Up
                myXY.setJogIn In_Joy_X1, False
                myXY.setJogIn In_Joy_X2, False
                myXY.setJogIn In_Joy_Y1, True
                myXY.setJogIn In_Joy_Y2, False
            Case joyUpRight  'Move Up Right
                myXY.setJogIn In_Joy_X1, True
                myXY.setJogIn In_Joy_X2, False
                myXY.setJogIn In_Joy_Y1, True
                myXY.setJogIn In_Joy_Y2, False
            Case joyRight  'Move Right
                myXY.setJogIn In_Joy_X1, True
                myXY.setJogIn In_Joy_X2, False
                myXY.setJogIn In_Joy_Y1, False
                myXY.setJogIn In_Joy_Y2, False
            Case joyDownRight  'Move Down Right
                myXY.setJogIn In_Joy_X1, True
                myXY.setJogIn In_Joy_X2, False
                myXY.setJogIn In_Joy_Y1, False
                myXY.setJogIn In_Joy_Y2, True
            Case joyDown  'Move Down
                myXY.setJogIn In_Joy_X1, False
                myXY.setJogIn In_Joy_X2, False
                myXY.setJogIn In_Joy_Y1, False
                myXY.setJogIn In_Joy_Y2, True
            Case joyDownLeft  'Move Down Left
                myXY.setJogIn In_Joy_X1, False
                myXY.setJogIn In_Joy_X2, True
                myXY.setJogIn In_Joy_Y1, False
                myXY.setJogIn In_Joy_Y2, True
            Case joyLeft  'Move Left
                myXY.setJogIn In_Joy_X1, False
                myXY.setJogIn In_Joy_X2, True
                myXY.setJogIn In_Joy_Y1, False
                myXY.setJogIn In_Joy_Y2, False
            Case joyUpLeft  'Move Up Left
                myXY.setJogIn In_Joy_X1, False
                myXY.setJogIn In_Joy_X2, True
                myXY.setJogIn In_Joy_Y1, True
                myXY.setJogIn In_Joy_Y2, False
            Case Else
        End Select

    End Sub


'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=













'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---=== Advanced Functions - Drill & Tap cycles; Move to Position ---===---===---===---===---===---===---==


' waitForMove - Returns the myXY check move result
Public Function waitForMove() As Boolean

    waitForMove = True
    
    waitForMove = myXY.checkMove

End Function

Public Sub stopNow()

    'Disable Drill, Tap, and Coolant Pump
    setDrillTap False, False, False

    'Stop motion
    myXY.moveStop
    
    'Set location variables to unknown position
    myDestination = atIdle
    myPosition = atUnk

End Sub

Public Function Home(motor As String) As Boolean

    Home = False
    Static CmdSent As Boolean

    'If there is currently a move in progress
    If myXY.checkMove Then Exit Function
        
    'Home the appropriate motor; exit function if the move does not start
    Select Case motor
        Case "X"
            If Not myXY.moveTo(0, 0, movHomeX) Then Exit Function
        Case "Y"
            If Not myXY.moveTo(0, 0, movHomeY) Then Exit Function
    End Select
    
    'If homing motion has been started, return true
    Home = True
    
End Function

Public Function moveTo(destination As atPosition) As Boolean

    moveTo = False

    If joyOn Then Exit Function
    
    If myPosition = destination Then Exit Function
    
    'If there is currently a move in progress, do not proceed
    If myXY.checkMove Then Exit Function

    'When the move has been completed, update the current position and set return true
    If myDestination = destination Then
        setPos destination
        myDestination = atIdle
        moveTo = True
        Exit Function
    End If

    'Set velocity for motion
    myXY.setVel mBoth, V_Jog
    'Move to the drill position
    Select Case destination
        Case atDrill
            myXY.moveTo goDrill.X, goDrill.Y, movY0XY
        Case atTap
            myXY.moveTo goTap.X, goTap.Y, movY0XY
        Case atUnk
            myXY.moveTo CDbl(frmMaintenance.txtMaintPos(0)), CDbl(frmMaintenance.txtMaintPos(1)), movY0XY
    End Select
    
    myDestination = destination

End Function

Public Function runTool(tool As myTool, totalDeep As Double, currentCycle As Integer, totalCycles As Integer, dryRun As Boolean) As Integer
    
    runTool = currentCycle
    
    If joyOn Then Exit Function

    'If there is currently a move in progress, do not proceed. Else acknowledge prior move
    If myXY.checkMove Then Exit Function

    Select Case currentCycle
        'On the prep cycle, initialize the drill and coolant if it is not a dry run. Delay 1 second
        Case 0
        
            'Power drill & Coolant if it is not a dry run
            If Not dryRun Then
                'Enable the appropriate tool and increment the use counter
                Select Case tool
                    Case tDrill
                        setDrillTap True, False, True
                        myUI.increment (countDrill)
                    Case tTap
                        setDrillTap False, True, True
                        myUI.increment (countTap)
                End Select
                
            End If
            
            runTool = 1
        
        'Second to last cycle, go to the tool's default position
        Case (totalCycles - 1)
        
            'Move to the default position for the selected tool
            Select Case tool
                Case tDrill
                    If myXY.moveTo(0, goDrill.Y, movY) Then runTool = currentCycle + 1
                Case tTap
                    If myXY.moveTo(0, goTap.Y, movY) Then runTool = currentCycle + 1
            End Select
        
        'If the last cycle has been reached, disable tool; delay for 1 second then exit
        Case totalCycles
            
            Static delayBool As Boolean
            
            If Not dryRun Then
                setDrillTap False, False, False
            End If
            
            'Delay for 1 second before proceeding
            If Not delayBool Then
                If myXY.moveDelay(1) Then delayBool = True
                Exit Function
            End If
            
            delayBool = False
            runTool = currentCycle + 1
            
        'All Other Cycles
        Case Else
                
            'First cycle, delay before running
            If currentCycle = 1 Then
                If Not myXY.moveDelay(1) Then Exit Function
                runTool = 2
                Exit Function
            End If
                
            'Compute the drill depth and send drill command. Cycles run from cycle 2 to total - 2 as cycle 1 is delay and final 2 cycles return to start
            Dim cycleDeep As Double
            cycleDeep = -((((currentCycle - 1) / (totalCycles - 3))) * totalDeep + stdOffset)
            
            If tool = tTap Then cycleDeep = cycleDeep + goTap.Y
            
            Dim moveResult As Boolean
            
            Select Case tool
                Case tDrill
                    moveResult = myXY.movePeck(cycleDeep, 0.125, V_Drill, V_Jog)
                Case tTap
                    moveResult = myXY.movePeck(cycleDeep, 10, V_Tap, V_Tap)
            End Select
            
            
            'If the move is sucessfully initiated, increment cycle counter
            If moveResult Then runTool = currentCycle + 1
        
        
    End Select

End Function

'---===---===---=== Advanced Functions - Drill & Tap cycles; Move to Position ---===---===---===---===---===---===---==
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=


