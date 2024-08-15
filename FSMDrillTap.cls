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

'State Variable
Private fsmState As fsmStates

'Inner Variables for controlling State loops
Private innerState As Integer
Private counter As Integer



'State request variables
Private joystickMode As Boolean
Private maintenanceMode As Boolean
Private runMode As Boolean





Public Sub initializeFSM()

    'Clear all variables
    joystickMode = False
    maintenanceMode = False
    runMode = False
    
    
    'Set state variable to 0
    fsmState = 0
    
End Sub

Private Sub setState(state As fsmStates)

    fsmState = state
    counter = 0
    stateInner = 0

End Sub

'When Possible, toggle run state
Public Sub setGoStop()

    runMode = Not runMode

End Sub

'When possible, go to Joystick State
Public Sub setJoystick()

    joystickMode = True

End Sub

'When possible, go to Maintenance State
Public Sub setMaintenance()

    maintenanceMode = True

End Sub


'Immediately go to ESTOP State
Public Sub setESTOP()

    fsmState = s7_ESTOP
    
End Sub










Public Sub runFSM()

    runAlwaysFirst
    
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

    runAlwaysLast

End Sub




Private Sub runState0()

    'Initialize FSM on first run

End Sub

Private Sub runAlwaysFirst()
    
    'Update faststatus
    myDrillTap.update
    
    'Operations that should run every loop, regardless of state (i.e. ESTOP check)
    myDrillTap.checkESTOP

End Sub

Private Sub runAlwaysLast()
    
    'Update faststatus
    myDrillTap.update
    
    'At the end of every loop, the DRO should be updated
    myUI.updDRO

End Sub

Private Sub runState1()
    
    'Standby State - Waits for USER to tell the machine to go

End Sub

Private Sub runState2()

    'Joystick State - Operates the virtual joystick via the keyboard
    
    'Run Joystick
    myDrillTap.joy myUI.getDir


    innerState = 0

End Sub

Private Sub runState3()

    'Mantenance State - Prevents all I/O from triggering program-related items. Allows for user actuated I/O. Runs joystick from a different location than state2

End Sub

Private Sub runState4()

    'Move State - Idles the FSM until current move is completed

End Sub

Private Sub runState5()

    'Prep State - Verifies that the clamp is set
    

End Sub

Private Sub runState6()

    'Run State - Runs Drill & Tap Cycles As Requested
            
    Dim runCycles As Integer
    
    Select Case innerState
        Case 0  'Initialise Function
            counter = 0
            
            'If the system is set to tap-only, then go to tap. Otherwise, go to drill
            If frmUI.optDrillTapAuto(1) Then
                innerState = 3
            Else
                innerState = 1
            End If
            
        Case 1  'Move-To-Drill - if necessary
        
            'If not at drill position, move to tap position
            If myDrillTap.getPos() <> atDrill Then
                myDrillTap.goToDrill
                Exit Sub
            End If
                    
            'If the counter has not been reset, exit the run state
            If counter > runCycles Then
            
                'but only if it isn't set to auto cycle
                If frmUI.optDrillTapAuto(3).Value Then
                    innerState = 5
                    Exit Sub
                End If
                
                innerState = 6
            
            'Otherwise, go to drill cycle
            Else
                innerState = 2
                runCycles = 8
            End If
            
        Case 2  'Run Drill
            
            'While the cycle counter is less than the desired number of cycles, run the tap
            If counter <= runCycles Then
                
                counter = myDrillTap.runTap(CDbl(frmUI.txtDeep(0)), counter, runCycles, frmUI.chkDryRun.Value)
            
            'When the cycle counter exceeds the number of desired cycles, wait for current move to finish
            Else
                'If the move has been completed, determine next step
                If myDrillTap.waitForMove Then
                    'If run drill only has been selected, move to drill then exit state
                    If frmUI.optDrillTapAuto(0).Value Then
                        innerState = 1
                    'If run Cycle or run Auto has been selected, proceed to tap cycle
                    ElseIf frmUI.optDrillTapAuto(2).Value Or frmUI.optDrillTapAuto(3).Value Then
                        innerState = 4
                        counter = 0
                    End If
                End If
            End If
            
            
        Case 3  'Move-To-Tap - if necessary
        
            'If not at tap position, move to tap position
            If myDrillTap.getPos() <> atTap Then
                myDrillTap.goToTap
                Exit Sub
            End If
                    
            'If the last state was the Run Tap state, and the counter has not been reset, exit the run state
            If counter > runCycles Then
                innerState = 6
            
            'Otherwise, go to tap cycle
            Else
                innerState = 4
                runCycles = 3
            End If
                
        Case 4  'Run Tap
            
            'While the cycle counter is less than the desired number of cycles, run the tap
            If counter <= runCycles Then
                
                counter = myDrillTap.runTap(CDbl(frmUI.txtDeep(1)), counter, runCycles, frmUI.chkDryRun.Value)
            
            'When the cycle counter exceeds the number of desired cycles, wait for current move to finish
            Else
                'If the move has been completed, determine next step
                If myDrillTap.waitForMove Then
                    'If run tap only has been selected, move to tap then exit state
                    If frmUI.optDrillTapAuto(1).Value Then
                        innerState = 3
                    'If run cycle or run auto was selected, move to drill
                    ElseIf frmUI.optDrillTapAuto(2).Value Or frmUI.optDrillTapAuto(3).Value Then
                        innerState = 1
                    End If
                End If
            End If
            
        Case 5  'Wait for Clamp, then start new cycle
            counter = 0
            
            Select Case counter
                Case 0  'Wait for clamp to be disengaged
                    If Not myDrillTap.checkClamp Then counter = 1
                    Exit Sub
                Case 1 'Wait for clamp to be reengaged
                    If myDrillTap.checkClamp Then
                        counter = 0
                        stateInner = 1
                    End If
                    Exit Sub
            End Select
                
            
        Case 6  'Exit State
            
            'When any pending moves have been completed, exit to standby state
            If myDrillTap.waitForMove() Then
            
                'Go to standby state
                setState s1_Stby
            End If
            
        
    End Select
    
    

End Sub

Private Sub runState7()

    'ESTOP State - Locks up everything except the UI and waits for user to clear ESTOP condition

End Sub

