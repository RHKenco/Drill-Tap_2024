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

    Const Out_Clamp_Sol = 2
    Const Out_Clamp_Pump = 3
    Const Out_Drill = 4
    Const Out_Tap = 5
    Const Out_Cool = 6
    
    Const In_Joy_X1 = 9
    Const In_Joy_X2 = 10
    Const In_Joy_Y1 = 11
    Const In_Joy_Y2 = 12
    Const In_TravelLim = 13
    Const In_Clamp_On = 14
    Const In_ESTOP = 15
    Const In_Clamp_Off = 16
    
    
    Const Axis_X = 2
    Const Axis_Y = 1

    Public Enum dtCo
        cNULL
        cX
        cY
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
    Private Enum atPosition
        atUnk
        atDrill
        atTap
    End Enum
    
    Private myPosition As atPosition
     
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
            
        'Create 6k movement control object
        myXY.init6k 2, 1, 32    'X-axis is motor 2, Y-axis is motor 1, acknowledge input is pin 32
        
        'Store calibration constants
        zeroOffset = myZeroOffset
        goDrill.X = 0
        goDrill.Y = 0
        goTap = myToolOffset
        
        'Initialize clamp functions
        myXY.initClamp Out_Clamp_Sol, Out_Clamp_Pump, In_Clamp_On, In_Clamp_Off
            
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
    
    Public Function getPos() As atPosition
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
    
    ''' setVel - Sets the internal velocity variables for drill/tap movement
    Public Sub setVel(JogVel As String, drillVel As String, TapVel As String)
    
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


    Public Function checkESTOP() As Boolean
            
        'Return false by default
        checkESTOP = False
        
        'Check estop; raise flag if true
        If myXY.chkESTOP(In_ESTOP) Then
            checkESTOP = True
        End If
        
    End Function
    
    Public Function resetESTOP()
        myXY.reset6k
    End Function



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
    
        checkClamp = myXY.checkOut(Out_Clamp_Sol)
    
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
                myXY.setIn In_Joy_X1, 2
                myXY.setIn In_Joy_X2, 2
                myXY.setIn In_Joy_Y1, 2
                myXY.setIn In_Joy_Y2, 2
            Case joyUp  'Move Up
                myXY.setIn In_Joy_X1, 2
                myXY.setIn In_Joy_X2, 2
                myXY.setIn In_Joy_Y1, 1
                myXY.setIn In_Joy_Y2, 2
            Case joyUpRight  'Move Up Right
                myXY.setIn In_Joy_X1, 1
                myXY.setIn In_Joy_X2, 2
                myXY.setIn In_Joy_Y1, 1
                myXY.setIn In_Joy_Y2, 2
            Case joyRight  'Move Right
                myXY.setIn In_Joy_X1, 1
                myXY.setIn In_Joy_X2, 2
                myXY.setIn In_Joy_Y1, 2
                myXY.setIn In_Joy_Y2, 2
            Case joyDownRight  'Move Down Right
                myXY.setIn In_Joy_X1, 1
                myXY.setIn In_Joy_X2, 2
                myXY.setIn In_Joy_Y1, 2
                myXY.setIn In_Joy_Y2, 1
            Case joyDown  'Move Down
                myXY.setIn In_Joy_X1, 2
                myXY.setIn In_Joy_X2, 2
                myXY.setIn In_Joy_Y1, 2
                myXY.setIn In_Joy_Y2, 1
            Case joyDownLeft  'Move Down Left
                myXY.setIn In_Joy_X1, 2
                myXY.setIn In_Joy_X2, 1
                myXY.setIn In_Joy_Y1, 2
                myXY.setIn In_Joy_Y2, 1
            Case joyLeft  'Move Left
                myXY.setIn In_Joy_X1, 2
                myXY.setIn In_Joy_X2, 1
                myXY.setIn In_Joy_Y1, 2
                myXY.setIn In_Joy_Y2, 2
            Case joyUpLeft  'Move Up Left
                myXY.setIn In_Joy_X1, 2
                myXY.setIn In_Joy_X2, 1
                myXY.setIn In_Joy_Y1, 1
                myXY.setIn In_Joy_Y2, 2
            Case Else
        End Select

    End Sub


'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=













'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---=== Advanced Functions - Drill & Tap cycles; Move to Position ---===---===---===---===---===---===---==

Public Function waitForMove() As Boolean

    waitForMove = False
    
    'If there is currently a move in progress, do not proceed. Else acknowledge prior move
    If myXY.checkMov Then Exit Function Else myXY.ackMov
    
    If goToDrill = True Then
        goToDrill = False
        myPosition = atDrill
    ElseIf goToTap = True Then
        goToTap = False
        myPosition = atTap
    Else
        myPosition = atUnk
    End If
    
    waitForMove = True

End Function

Public Function goToDrill() As Boolean

    goToDrill = False

    If joyOn Then Exit Function
    
    'If there is currently a move in progress, do not proceed. Else acknowledge prior move
    If myXY.checkMov Then Exit Function Else myXY.ackMov

    'Move to the drill position
    myXY.moveTo goDrill, movY0XY


    myPosition = atDrill
    
    goToDrill = True

End Function

Public Function goToTap() As Boolean
    
    goToTap = False
    
    If joyOn Then Exit Function
    
    'If there is currently a move in progress, do not proceed. Else acknowledge prior move
    If myXY.checkMov Then Exit Function Else myXY.ackMov

    'Move to the drill position
    myXY.moveTo goTap, movY0XY
    
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
    If myXY.checkMov Then Exit Function Else myXY.ackMov

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
    If myXY.checkMov Then Exit Function Else myXY.ackMov

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
    
    myXY.movePeck cycleDeep, 0.5, V_Tap, V_Tap
    
    
    'If the move is sucessfully initiated, increment cycle counter
    runTap = currentCycle + 1
    

End Function







'---===---===---=== Advanced Functions - Drill & Tap cycles; Move to Position ---===---===---===---===---===---===---==
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=


