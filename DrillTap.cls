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
    Const In_Clamp_Lim = 12
    Const In_TravelLim = 13
    Const In_Clamp_On = 14
    Const In_ESTOP = 15
    Const In_Clamp_Off = 16
    

    
    
    Const Axis_X = 2
    Const Axis_Y = 1

    Public Enum dtCo
        dNULL
        dX
        dY
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
        atMove
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
    Private V_Jog As String
    Private V_Drill As String
    Private V_Tap As String
    
    'Boolean flag for joystick state - prevent certain actions while joystick is on
    Private joyOn As Boolean
    Private zeroSet As Boolean
    
    
    'Object to store the last DRO reading when a set-zero occurred.
    Private PosLastZero As myCoordinate
    
    'Objects to store the offsets for the drill & tap locations.
    Private zeroOffset As myCoordinate  'Location of drill relative to hard stops of machine
    Private goTap As myCoordinate  'Location of Tap relative to Drill
    Private goDrill As myCoordinate
    
    Private maintIO(17) As Boolean
    
    Private myXY As New Class6kXY
    

'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=








'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------
    Public Sub initDrillTap(stdOffset As Double, myToolOffsetX As Double, myToolOffsetY As Double, myZeroOffsetX As Double, myZeroOffsetY As Double)
        
        'Initialize class variables
        zeroSet = False
        joyOn = False
        ESTOP_SET = False
            
        'Create 6k movement control object
        myXY.init6k Axis_X, Axis_Y, In_MoveAck, In_ESTOP, In_TravelLim    'X-axis is motor 2, Y-axis is motor 1, acknowledge input is pin 32
        
        'Store calibration constants
        zeroOffset.X = myZeroOffsetX
        zeroOffset.Y = myZeroOffsetY
        
        goDrill.X = 0
        goDrill.Y = 0
        goTap.X = myToolOffsetX
        goTap.Y = myToolOffsetY
        
        'Initialize clamp functions
        'myXY.initClamp Out_Clamp_Sol, Out_Clamp_Pump, In_Clamp_On, In_Clamp_Off
            
    End Sub
    
    Public Function setHome()
    
        'Get current position, and store it as the last zero location
        PosLastZero.X = getCoords(cX)
        PosLastZero.Y = getCoords(cY)
        
        'Invert the coordinates so the last zero is stored relative to the current zero (not the other way around)
        PosLastZero.X = -PosLastZero.X
        PosLastZero.Y = -PosLastZero.Y
        
        'Set zero using the zeroOffset
        myXY.setZero zeroOffset.X, zeroOffset.Y
        
        'Indicate that the zero is properly set
        zeroSet = True
        
    End Function

'---------------------------- Read-Only Set & Returns -----------------------------------------------------------------
        
        
    Public Sub update()
        myXY.update
    End Sub
    
        
        
    ''' getPos - Gets the current position of both axes
    Public Function getCoords(coordinate As dtCo) As Double
        
        Select Case coordinate
            Case cX
                getCoords = myXY.getPosX()
            Case cY
                getCoords = myXY.getPosY()
            Case Else
        End Select
        
    End Function
    
    Public Function getPos() As Integer
        
        getPos = myPosition
    End Function
    
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
    Public Sub setVelDefaults(JogVel As String, drillVel As String, TapVel As String)
    
        V_Jog = JogVel
        V_Drill = drillVel
        V_Tap = TapVel
    
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
    Public Function checkESTOP(Active As Boolean) As Boolean
                       
        'If estop flag set, pin data irrelevant. If it is not set, check estop pin
        If Not ESTOP_SET Then ESTOP_SET = myXY.chkESTOP(Active)
        
        If ESTOP_SET Then myFSM.setESTOP
        
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
 
    Public Sub setDrillTap(drillState As Boolean, tapState As Boolean, coolantState As Boolean)
    
        myXY.setOut Out_Drill, drillState
        
        myXY.setOut Out_Tap, tapState
        
        myXY.setOut Out_Cool, coolantState
        
    
    End Sub
 
    Public Sub setClamp(clampState As Boolean)
    
        If clampState Then myXY.virtualPress In_Clamp_On, 0.25 Else myXY.virtualPress In_Clamp_Off, 0.25
        
    End Sub
    
    Public Function checkClamp() As Boolean
    
        checkClamp = myXY.chkIn(In_Clamp_Lim)
    
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
    
    Public Sub joy(direction As joyDirection)
    
        If Not joyOn Then Exit Sub
    
        Select Case direction
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

Public Function waitForMove() As Boolean

    waitForMove = False
    
    Dim checkForMove As Integer
    
    checkForMove = myXY.checkMov
    
    'If there is no move currently in progress, return true
    If checkForMove = 0 Then
        GoTo NoMoveInProgress
    'If there is a move in progress
    Else
        'If the move is completed, acknowledge the completion
        If checkForMove = 2 Then
            'Do not continue until the acnowledgement is confirmed
            If Not myXY.ackMov Then Exit Function
        'If the move is ongoing, or the input errors, exit function
        Else
            Exit Function
        End If
    End If
    
    Select Case myDestination
        Case atDrill
            myPosition = atDrill
        Case atTap
            myPosition = atTap
        Case Else
            myPosition = atUnk
    End Select
    
NoMoveInProgress:
    
    waitForMove = True

End Function

Public Function goToPosition(positionX As Double, positionY As Double) As Boolean

    goToPosition = False
    
    If joyOn Then Exit Function
    
    'If there is currently a move in progress, do not proceed
    If myXY.checkMov Then Exit Function

    'Set velocity for motion
    myXY.setVel CDbl(V_Jog)
    'Move to the drill position
    myXY.moveTo positionX, positionY, movYX


    myDestination = atUnk
    
    goToPosition = True

End Function

Public Function goToDrill() As Boolean

    goToDrill = False

    If joyOn Then Exit Function
    
    If myPosition = atDrill Then Exit Function
    
    'If there is currently a move in progress, do not proceed
    If myXY.checkMov Then Exit Function

    'Set velocity for motion
    myXY.setVel CDbl(V_Jog)
    'Move to the drill position
    myXY.moveTo goDrill.X, goDrill.Y, movY0XY


    myDestination = atDrill
    
    goToDrill = True

End Function

Public Function goToTap() As Boolean
    
    goToTap = False
    
    If joyOn Then Exit Function
    
    If myPosition = atTap Then Exit Function
    
    'If there is currently a move in progress, do not proceed.
    If myXY.checkMov Then Exit Function

    'Set velocity for motion
    myXY.setVel CDbl(V_Jog)
    'Move to the Tap position
    myXY.moveTo goTap.X, goTap.Y, movY0XY
    
    
    myDestination = atTap
    goToTap = True

End Function


Public Function runDrill(drillDeep As Double, currentCycle As Integer, totalCycles As Integer, dryRun As Boolean) As Integer
    
    runDrill = currentCycle
    
    If joyOn Then Exit Function
    
    'On the prep cycle, initialize the drill and coolant if it is not a dry run. Delay 1 second
    If currentCycle = 0 Then
        
        'Power drill & Coolant if it is not a dry run
        If Not dryRun Then
            myXY.setOut Out_Drill, True
            myXY.setOut Out_Cool, True
            
            'Increment drill cycle counter
            myUI.increment (countDrill)
        End If
        
        myXY.moveDelay 1
        

    
    End If
    
    'If there is currently a move in progress, do not proceed. Else acknowledge prior move
    If myXY.checkMov Then
        Exit Function
    Else
        If Not myXY.ackMov Then Exit Function
    End If

    'If the last cycle has been reached, disable drill and exit
    If currentCycle = totalCycles Then
        If Not dryRun Then
            myXY.setOut Out_Drill, False
            myXY.setOut Out_Cool, False
        End If
        
        runDrill = currentCycle + 1
        Exit Function
        
    End If

    'Compute the drill depth and send drill command
    Dim cycleDeep As Double
    cycleDeep = (currentCycle / (totalCycles - 1)) * drillDeep
    
    'Set velocity for motion
    myXY.setVel CDbl(V_Drill)
    myXY.movePeck cycleDeep, 0.5, V_Drill, V_Jog
    
    
    'If the move is sucessfully initiated, increment cycle counter
    runDrill = currentCycle + 1

End Function

Public Function runTap(tapDeep As Double, currentCycle As Integer, totalCycles As Integer, dryRun As Boolean) As Integer

    If joyOn Then Exit Function
    
    'On the prep cycle, initialize the drill and coolant if it is not a dry run. Delay 1 second
    If currentCycle = 0 Then
        
        If Not dryRun Then
            myXY.setOut Out_Tap, True
            myXY.setOut Out_Cool, True
        End If
        
        myXY.moveDelay 1
        
        'Increment tap cycle counter
        myUI.increment (countTap)
    
    End If
    
    'If there is currently a move in progress, do not proceed. Else acknowledge prior move
    If myXY.checkMov Then
        Exit Function
    Else
        If Not myXY.ackMov Then Exit Function
    End If

    'If the last cycle has been reached, disable Tap and exit
    If currentCycle = totalCycles Then
        If Not dryRun Then
            myXY.setOut Out_Tap, False
            myXY.setOut Out_Cool, False
        End If
        
        runTap = currentCycle + 1
        Exit Function
        
    End If

    'Compute the drill depth and send drill command
    Dim cycleDeep As Double
    cycleDeep = (currentCycle / (totalCycles - 1)) * drillDeep
    
    'Set velocity for motion
    myXY.setVel CDbl(V_Tap)
    myXY.movePeck cycleDeep, 0.5, V_Tap, V_Tap
    
    
    'If the move is sucessfully initiated, increment cycle counter
    runTap = currentCycle + 1
    

End Function







'---===---===---=== Advanced Functions - Drill & Tap cycles; Move to Position ---===---===---===---===---===---===---==
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=


