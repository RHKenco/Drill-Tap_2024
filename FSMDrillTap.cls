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

End Sub

Private Sub runState7()

    'ESTOP State - Locks up everything except the UI and waits for user to clear ESTOP condition

End Sub

