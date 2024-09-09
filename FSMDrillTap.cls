VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFSMDrillTap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- Drill/Tap FSM Class Container -----
'|
'|
'|
'|====================================================================================================================|

Option Explicit


'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===

'--------------------------- Class Enumerations & Constants -----------------------------------------------------------

Public Enum fsmStates
    s0_Init
    s1_Stby
    s2_Joy
    s3_Maint
    s4_Move
    s5_Prep
    s6_Run
    s7_ESTOP
End Enum


Public Enum moveTo
    moveNone
    moveInProgress
    moveToDrillNow
    moveToTapNow
    moveToPositionNow
End Enum


Public Enum boolBehavior
    boolFalse
    boolTrue
    boolToggle
End Enum

'State Variable
Private fsmState As fsmStates

'Inner Variables for controlling State loops
Private innerState As Integer
Private Counter As Integer



'State request variables
Private joystickMode As Boolean
Private maintenanceMode As Boolean
Private runMode As Boolean
Private moveMode As moveTo
Private clearESTOPnow As Boolean





Public Sub initializeFSM()

    'Clear all variables
    joystickMode = False
    maintenanceMode = False
    runMode = False
    moveMode = moveNone
    clearESTOPnow = False
    
    clearESTOPnow = False
    
    'Set state variable to 0
    fsmState = 0
    
End Sub

Private Sub setState(state As fsmStates)

    fsmState = state
    Counter = 0
    innerState = 0

End Sub

'When Possible, toggle run state
Public Sub setGoStop(state As boolBehavior)

    Select Case state
        Case boolFalse
            runMode = False
        Case boolTrue
            runMode = True
        Case boolToggle
            runMode = Not runMode
    End Select

End Sub

'Set move destination; when possible, move to position
Public Sub setMoveDest(myDestination As moveTo)
    
    moveMode = myDestination

End Sub


'When possible, go to Joystick State
Public Sub setJoystick(state As boolBehavior)

    Select Case state
        Case boolFalse
            joystickMode = False
        Case boolTrue
            joystickMode = True
        Case boolToggle
            joystickMode = Not joystickMode
    End Select

End Sub

'When possible, go to Maintenance State
Public Sub setMaintenance(enable As Boolean)

    maintenanceMode = enable

End Sub


'Immediately go to ESTOP State
Public Sub setESTOP()

    myDrillTap.checkESTOP True, True
    
    
    'Clear FSM so it goes back to Standby after ESTOP is cleared
    joystickMode = False
    runMode = False
    moveMode = moveNone

    myUI.updStatus uiESTOP, ""

    setState s7_ESTOP
    
End Sub

Public Sub clearESTOP()

    clearESTOPnow = True

End Sub


Public Sub setHome()

    setState s0_Init

End Sub







Public Sub runFSM()

    runAlways
    
    Select Case fsmState
        Case s0_Init
            runState0
        Case s1_Stby
            runState1
        Case s2_Joy
            runState2
        Case s3_Maint
            runState3
        Case s4_Move
            runState4
        Case s5_Prep
            runState5
        Case s6_Run
            runState6
        Case s7_ESTOP
            runState7
    End Select

End Sub




Private Sub runState0()

    'Wait for simulated move to be cleared before switching from initialize to standby
    If myDrillTap.waitForMove Then Exit Sub
    
    If innerState = 0 Then
        Dim msgRtn As VbMsgBoxResult
        msgRtn = MsgBox("Preparing to Home machine. Check for obstructions!" & vbCrLf & "Clear to proceed?", vbYesNo, "Check Machine")
        If msgRtn = vbNo Then
            Unload frmUI
            Exit Sub
        Else
            innerState = 1
        End If
    ElseIf innerState = 1 Then
        If Not myDrillTap.Home("Y") Then Exit Sub
        innerState = 2
        myUI.updStatus uiHoming, "Y"
    ElseIf innerState = 2 Then
        If Not myDrillTap.Home("X") Then Exit Sub
        innerState = 3
        myUI.updStatus uiHoming, "X"
    ElseIf innerState = 3 Then
        'Initialize FSM on first run
        myUI.updStatus uiStandby, ""
        setState s1_Stby
    End If

End Sub

Private Sub runAlways()
    
    'Update maintenance when necessary
    If maintenanceOpen Then myDrillTap.maintPassive
    
    'Check ESTOP Unless in Active Maintenance Mode
    myDrillTap.checkESTOP ((Not frmMaintenance.frameMaintIO.Enabled))
    
    'Check user inputs
    Dim isRunning As Boolean
    If (fsmState = s6_Run) Then isRunning = True
    
    myDrillTap.checkInputs (isRunning)
    
    'Update the DRO
    myUI.updDRO

End Sub

Private Sub runState1()
    
    'Standby State - Waits for USER to tell the machine to go
    
    'Check each move-mode boolean and set state according to order of priority
    
        If joystickMode Then
            
            setState s2_Joy
        
        ElseIf maintenanceMode Then
            
            setState s3_Maint
        
        ElseIf moveMode Then
            
            setState s4_Move
        
        ElseIf runMode Then
        
            setState s5_Prep
            
        Else
            'If the state does not change, exit sub
            Exit Sub
        End If
        
End Sub

Private Sub runState2()

    'Joystick State - Operates the virtual joystick via the keyboard
    
    'If first run, enable joystick
    If (innerState = 0) Or (maintenanceMode And (innerState = 1)) Then
    
        myDrillTap.joyState True
        
        'Set UI state
        myUI.updStatus uiJoystick, ""
        
        innerState = 1
        If maintenanceMode Then innerState = 2
    
    End If
    
    'While joystick mode is on, run joystick
    If joystickMode Then
    
        'Run Joystick
        myDrillTap.joy myUI.getDir
        
        'If maintenance mode is also enabled, cycle to maintenance state
        If maintenanceMode Then fsmState = s3_Maint
        
    'When joystick mode is disabled, shut down joystick and exit state
    Else
    
        myDrillTap.joyState False
        
        If maintenanceMode Then
            fsmState = s3_Maint
            innerState = 1
        Else
            myUI.updStatus uiStandby, ""
            setState s1_Stby
        End If
        
    End If

End Sub

Private Sub runState3()

    'Mantenance State - Prevents all I/O from triggering program-related items. Allows for user actuated I/O. Runs joystick from a different location than state2

    'first run, set up I/O defaults
    If (innerState = 0) Or (joystickMode And (innerState = 1)) Then


        innerState = 1
        If joystickMode Then innerState = 2

    End If

    If maintenanceMode Then

        'If joystick is also enabled, cycle to joystick state
        If joystickMode Then fsmState = s2_Joy

    
    'When maintenance mode is disabled, shut down maintenance and exit state
    Else
        If joystickMode Then
            fsmState = s2_Joy
            innerState = 1
        Else
            myUI.updStatus uiStandby, ""
            setState s1_Stby
        End If
    End If


End Sub

Private Sub runState4()

    'Move State - Idles the FSM until current move is completed
    
    
    If Not CBool(moveMode) Then
    
        'Stop motion
        myDrillTap.stopNow
        
        'Go to standby state
        myUI.updStatus uiStandby, ""
        moveMode = moveNone
        setState s1_Stby
    
    End If
    
    
    If innerState = 0 Then
        myUI.updStatus uiWaiting, ""
        innerState = 1
    End If
    
    'If move is requested, determine destination and call move command. Continue calling command until confirmation is recieved.
    If moveMode > moveInProgress Then
        Select Case moveMode
            Case moveToDrillNow
                If myDrillTap.moveTo(atDrill) Then
                    moveMode = moveInProgress
                ElseIf myDrillTap.getPos = atDrill Then
                    moveMode = moveNone
                Else
                    Exit Sub
                End If
            Case moveToTapNow
                If myDrillTap.moveTo(atTap) Then
                    moveMode = moveInProgress
                ElseIf myDrillTap.getPos = atTap Then
                    moveMode = moveNone
                Else
                    Exit Sub
                End If
            Case moveToPositionNow
                If myDrillTap.moveTo(atUnk) Then
                    moveMode = moveInProgress
                Else
                    Exit Sub
                End If
        End Select
        
    End If
    
    'When any pending moves have been completed, exit to standby state
    If Not myDrillTap.waitForMove() Then
    
        'Go to standby state
        myUI.updStatus uiStandby, ""
        moveMode = moveNone
        setState s1_Stby
    End If

End Sub

Private Sub runState5()

    'Prep State - Verifies that the clamp is set
    
    'If the clamp is set, go to run state
    If myDrillTap.checkClamp Then
        setState s6_Run
        
    'If it is not set, update status message and return to standby
    Else
    
    myUI.updStatus uiClampOFF, ""
    
    setState s1_Stby
    setGoStop boolFalse
    
    End If

End Sub

Private Sub runState6()

    'Run State - Runs Drill & Tap Cycles As Requested
            
            
    If Not runMode Then

        'Stop motion
        myDrillTap.stopNow
        
        'Go to standby state
        myUI.updStatus uiStandby, ""
        moveMode = moveNone
        setState s1_Stby
    
    End If
            
    Static runCycles As Integer
    
    Select Case innerState
        Case 0  'Initialise Function
            Counter = 0
            
            'If the system is set to tap-only, then go to tap. Otherwise, go to drill
            If frmUI.optDrillTapAuto(1) Then
                innerState = 3
            Else
                innerState = 1
            End If
            
        Case 1, 3 'Move-To-Position - if necessary
        
            'Check to ensure the tool is in the correct position
            Select Case innerState
                Case 1
                    'If not at drill position, move to drill position
                    If myDrillTap.getPos() <> atDrill Then
                        myDrillTap.moveTo (atDrill)
                        myUI.updStatus uiWaiting, ""
                        Exit Sub
                    End If
                Case 3
                    'If not at tap position, move to tap position
                    If myDrillTap.getPos() <> atTap Then
                        myDrillTap.moveTo (atTap)
                        myUI.updStatus uiWaiting, ""
                        Exit Sub
                    End If
            End Select
                    
            'If the counter has not been reset, exit the run state
            If Counter > runCycles Then
            
                'Unless the tool has just been moved to the drill and auto cycle is enabled
                If innerState = 1 And frmUI.optDrillTapAuto(3).Value Then
                    innerState = 5
                    Counter = 0
                    Exit Sub
                End If
                
                innerState = 6
            
            'Otherwise, go to appropriate tool cycle
            Else
                Select Case innerState
                    Case 1
                        innerState = 2
                        '7 cycles + 2 prep cycles (0 and 1) + 2 conclusion cycles
                        runCycles = 7 + 3
                    Case 3
                        innerState = 4
                        '3 cycles + 2 prep cycles (0 and 1) + 2 conclusion cycles
                        runCycles = 3 + 3
                End Select
            End If
            
        Case 2, 4  'Run Tool
            
            'While the cycle counter is less than the desired number of cycles, run the tap
            If Counter <= runCycles Then
                
                Dim currentTool As myTool
            
                Select Case innerState
                    Case 2
                        'First loop, set UI
                        If Counter = 0 Then myUI.updStatus uiRunningDrill, ""
                        'Identify correct tool
                        currentTool = tDrill
                    Case 4
                        'First loop, set UI
                        If Counter = 0 Then myUI.updStatus uiRunningTap, ""
                        'Identify correct tool
                        currentTool = tTap
                End Select
                
                Counter = myDrillTap.runTool(currentTool, CDbl(frmUI.txtDeep(0)), Counter, runCycles, frmUI.chkDryRun.Value)
            
            'When the cycle counter exceeds the number of desired cycles, wait for current move to finish
            Else
                'If the move has been completed, determine next step
                If Not myDrillTap.waitForMove Then
                    Select Case innerState
                        Case 2
                            'If run drill only has been selected, move to drill then exit state
                            If frmUI.optDrillTapAuto(0).Value Then
                                innerState = 1
                            'If run Cycle or run Auto has been selected, proceed to tap cycle
                            ElseIf frmUI.optDrillTapAuto(2).Value Or frmUI.optDrillTapAuto(3).Value Then
                                innerState = 3
                                Counter = 0
                            End If
    
                        Case 4
                            'If run tap only has been selected, move to tap then exit state
                            If frmUI.optDrillTapAuto(1).Value Then
                                innerState = 3
                            'If run cycle or run auto was selected, move to drill
                            ElseIf frmUI.optDrillTapAuto(2).Value Or frmUI.optDrillTapAuto(3).Value Then
                                innerState = 1
                            End If
                    End Select
                End If
            End If
            
        Case 5  'Wait for Clamp, then start new cycle
        
            myDrillTap.checkInputs False
        
            Select Case Counter
                Case 0
                    'Update UI Msg
                    myUI.updStatus uiClampCycle1, ""
                    Counter = 1
                    Exit Sub
                Case 1
                    'Wait for clamp to be disengaged
                    If Not myDrillTap.checkClamp Then
                        Counter = 2
                        myUI.updStatus uiClampCycle2, ""
                    End If
                    
                    Exit Sub
                Case 2
                    'Wait for clamp to be reengaged
                    If myDrillTap.checkClamp Then
                        Counter = 0
                        innerState = 1
                    End If
                    
                    Exit Sub
            End Select
                
            
        Case 6  'Exit State
            
            'When any pending moves have been completed, exit to standby state
            If Not myDrillTap.waitForMove() Then
            
                'Exit run mode
                setGoStop False
            
                'Go to standby state
                myUI.updStatus uiStandby, ""
                setState s1_Stby
                
            End If
            
        
    End Select
    
    

End Sub

Private Sub runState7()

    'ESTOP State - Locks up everything except the UI and waits for user to clear ESTOP condition
    
    
    'If ESTOP cleared
    If clearESTOPnow Then
        
        myDrillTap.resetESTOP
        
        myUI.updStatus uiStandby, ""
        setState s1_Stby
        
        clearESTOPnow = False
    End If
    
    

End Sub

