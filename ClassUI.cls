VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'|====================================================================================================================|
'|
'|              ----- Drill/Tap UI Update Class -----
'|
'|
'|
'|====================================================================================================================|

Option Explicit


'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===

'--------------------------- Class Enumerations & Constants -----------------------------------------------------------



Public Enum uiStatus
    uiLoading
    uiStandby
    uiWaiting
    uiJoystick
    uiClampON
    uiClampOFF
    uiESTOP
    uiRunningDrill
    uiRunningTap
    uiError
End Enum

Public Enum btnState
    btnInit
    btnStby
    btnRunning
    btnESTOP
End Enum

'Keypress booleans to track the state of multiple keys
Public uiKeyUp As Boolean
Public uiKeyRight As Boolean
Public uiKeyDown As Boolean
Public uiKeyLeft As Boolean
     
'--------------------------- Class Variables --------------------------------------------------------------------------


    

'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=








'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------
    Public Sub initUI()
            
    End Sub

'---------------------------- Read-Only Set & Returns -----------------------------------------------------------------
        

Public Function getDir() As joyDirection
    
    Dim keyDir As Integer
    Dim upDown As Integer
    Dim leftRight As Integer
    
    keyDir = 0
    upDown = 0
    leftRight = 0
    
    'Establish directional integers
    If uiKeyUp Xor uiKeyDown Then

        If uiKeyUp Then upDown = 3
        If uiKeyUp Then upDown = 6
        
    End If
    
    If uiKeyRight Xor uiKeyLeft Then
    
        If uiKeyRight Then leftRight = 1
        If uiKeyRight Then leftRight = 2
        
    End If
    
    keyDir = upDown + leftRight
    
    Select Case keyDir
        Case 3
            getDir = joyUp
        Case 4
            getDir = joyUpRight
        Case 2
            getDir = joyRight
        Case 8
            getDir = joyDownRight
        Case 6
            getDir = joyDown
        Case 7
            getDir = joyDownLeft
        Case 1
            getDir = joyLeft
        Case 4
            getDir = joyUpLeft
        Case Else
            getDir = joyNeutral
    End Select
    
End Function


'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=













'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---=== Advanced Functions - Status Message,  ---===---===---===---===---===---===---==

Public Sub go()

    'If the estop is not set
    If Not myXY.estopSet Then
    
        'Call the FSM to go to the run state
        myFSM.setGoStop
    
    'If estip is set
    Else
        'Query the user to reset the ESTOP flag
        Dim messageReturn As VbMsgBoxResult
        messageReturn = MsgBox("Check machine before clearing ESTOP." & vbCrLf & vbCrLf & "Clear Estop?", vbOKCancel)
        
        'If the user acknowledges the reset
        If messageReturn = vbOK Then
        
            'Call the FSM to clear the ESTOP
            myFSM.setClrESTOP
            
        End If
    End If
    
End Sub


Public Sub btnState(newState As btnState)

    Select Case btnState
        Case btnInit
        
            frmUI.cmdGO.Caption = "GO"
            frmUI.cmdGO.Enabled = False
            
            frmUI.cmdGoToDrillTap(0).Enabled = False
            frmUI.cmdGoToDrillTap(1).Enabled = False
            
            frmUI.cmdResetCount(0).Enabled = False
            frmUI.cmdResetCount(0).Enabled = False
            
        Case btnStby
        
            frmUI.cmdGO.Caption = "GO"
            frmUI.cmdGO.Enabled = True
            
            frmUI.cmdGoToDrillTap(0).Enabled = True
            frmUI.cmdGoToDrillTap(1).Enabled = True
            
            frmUI.cmdResetCount(0).Enabled = True
            frmUI.cmdResetCount(0).Enabled = True
            
        Case btnRunning
        
            frmUI.cmdGO.Caption = "STOP"
            frmUI.cmdGO.Enabled = True
            
            frmUI.cmdGoToDrillTap(0).Enabled = False
            frmUI.cmdGoToDrillTap(1).Enabled = False
            
            frmUI.cmdResetCount(0).Enabled = False
            frmUI.cmdResetCount(0).Enabled = False
            
        Case btnESTOP

            frmUI.cmdGO.Caption = "CLEAR ESTOP"
            frmUI.cmdGO.Enabled = True
            
            frmUI.cmdGoToDrillTap(0).Enabled = False
            frmUI.cmdGoToDrillTap(1).Enabled = False
            
            frmUI.cmdResetCount(0).Enabled = False
            frmUI.cmdResetCount(0).Enabled = False
            
        Case Else
    End Select

End Sub




Public Sub updStatus(myStatus As uiStatus, strInput As String)

    Dim myMsg As String

    Select Case myStatus
        Case uiLoading
            myMsg = "System Loading, Please Wait"
        Case uiStandby
            myMsg = "Waiting for Input"
        Case uiWaiting
            myMsg = "Motion in Progress, Please Wait"
        Case uiJoystick
            myMsg = "Joystick Enabled"
        Case uiClampON
            myMsg = "Clamp On, Waiting for Input"
        Case uiClampOFF
            myMsg = "Clamp Off - Remove Shank"
        Case uiESTOP
            myMsg = "ESTOP Triggered - Check Machine before Clearing"
        Case uiRunningDrill
            myMsg = "Running Drill Cycle"
        Case uiRunningTap
            myMsg = "Running Tap Cycle"
        Case uiError
            myMsg = "ERROR: " & strInput
    End Select

    frmUI.lblSystemStatus.Caption = myMsg
    frmUI.lblSystemStatus.Refresh

End Sub



'---===---===---=== Advanced Functions -  ---===---===---===---===---===---===---==
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=



