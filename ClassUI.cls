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

Public Enum uiCo
    cNULL
    cX
    cY
End Enum
    

Public Enum uiStatus
    uiLoading
    uiStandby
    uiWaiting
    uiJoystick
    uiClampOFF
    uiClampOFFcycle
    uiESTOP
    uiRunningDrill
    uiRunningTap
    uiMaintIdle
    uiMaintEnabled
    uiError
End Enum

Public Enum btnState
    btnInit
    btnStby
    btnRunning
    btnJoystick
    btnESTOP
    btnMaintOn
End Enum

Public Enum uiTxtbxs
    txtDrillSpeed
    txtTapSpeed
    txtJogSpeed
    txtDrillTo
    txtTapTo
    txtMaintX
    txtMaintY
End Enum

Public Enum uiCounter
    countNull
    countDrill
    countTap
    countBoth
End Enum

Const savePath = "C:\Drill-Tap_2024\savedData\cycles.txt"


'Keypress booleans to track the state of multiple keys
Public uiKeyUp As Boolean
Public uiKeyRight As Boolean
Public uiKeyDown As Boolean
Public uiKeyLeft As Boolean

'Store Drill & Tap Cycles
Private drillCycles As Integer
Private tapCycles As Integer
Private drillAvg As Double
Private tapAvg As Double
Private drillAvgCy As Integer
Private tapAvgCy As Integer


'Validation strings for textbox numeric entry
Private txtbxDrillSpeedCHK As String
Private txtbxTapSpeedCHK As String
Private txtbxJogSpeedCHK As String
Private txtbxDrillToCHK As String
Private txtbxTapToChk As String
Private txtbxMaintXCHK As String
Private txtbxMaintYCHK As String
     
'--------------------------- Class Variables --------------------------------------------------------------------------


    

'---===---===---===---===--- Setup - Constants & Variables ---===---===---===---===---===---===---===---===---===---===
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=








'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--

'--------------------------- Init & Clear Functions -------------------------------------------------------------------
    Public Sub initUI()
            
        'Set initial UI status
        updStatus uiLoading, ""
                
        'Load text variable defaults
        txtbxDrillSpeedCHK = frmUI.txtSpeeds(0).Text
        txtbxTapSpeedCHK = frmUI.txtSpeeds(1).Text
        txtbxJogSpeedCHK = frmUI.txtSpeeds(2).Text
        txtbxDrillToCHK = frmUI.txtDeep(0).Text
        txtbxTapToChk = frmUI.txtDeep(1).Text
        txtbxMaintXCHK = frmMaintenance.txtMaintPos(0).Text
        txtbxMaintYCHK = frmMaintenance.txtMaintPos(1).Text
        
        'Load drill & tap counters
        loadCycles
            
            
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
        If uiKeyDown Then upDown = 6
        
    End If
    
    If uiKeyRight Xor uiKeyLeft Then
    
        If uiKeyRight Then leftRight = 1
        If uiKeyLeft Then leftRight = 2
        
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

Public Function getCoords(coordinate As uiCo) As Double

    Select Case coordinate
        Case cX
            getCoords = CDbl(frmMaintenance.txtMaintPos(0))
        Case cY
            getCoords = CDbl(frmMaintenance.txtMaintPos(1))
        Case Else
    End Select

End Function


'---===---===---===---===--- Basic Functions - Initialize, Read-Only, Etc. ---===---===---===---===---===---===---===--
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=













'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=
'---===---===---=== Advanced Functions - Status Message,  ---===---===---===---===---===---===---==

Public Sub validate(textbox As uiTxtbxs, inputText As String)
    
    'Check if the input value is numeric
    Dim tmpChk As Boolean
    tmpChk = False
    
    If IsNumeric(inpuText) Then tmpChk = True

    'Determine which textbox is being validated. If the text is numeric, save current state. Else reset to previous
    Select Case textbox
        Case txtDrillSpeed
            If tmpChk Then txtbxDrillSpeedCHK = inputText Else frmUI.txtSpeeds(0).Text = txtbxDrillSpeedCHK
        Case txtTapSpeed
            If tmpChk Then txtbxTapSpeedCHK = inputText Else frmUI.txtSpeeds(1).Text = txtbxTapSpeedCHK
        Case txtJogSpeed
            If tmpChk Then txtbxJogSpeedCHK = inputText Else frmUI.txtSpeeds(2).Text = txtbxJogSpeedCHK
        Case txtDrillTo
            If tmpChk Then txtbxDrillToCHK = inputText Else frmUI.txtDeep(0).Text = txtbxDrillToCHK
        Case txtTapTo
            If tmpChk Then txtbxTapToChk = inputText Else frmUI.txtDeep(1).Text = txtbxTapToChk
        Case txtMaintX
            If tmpChk Then txtbxMaintXCHK = inputText Else frmUI.txtMaintPos(0).Text = txtbxMaintXCHK
        Case txtMaintY
            If tmpChk Then txtbxMaintXCHK = inputText Else frmUI.txtMaintPos(1).Text = txtbxMaintXCHK
    End Select

End Sub

Public Function reformat(inputText As String)

    Dim outputText As String

    If Abs(CDbl(inputText)) < 10 Then outputText = Format(inputText, "###0.000")
    If Abs(CDbl(inputText)) >= 10 Then outputText = Format(inputText, "###0.00")

    reformat = outputText

End Function


Public Sub go()

    'If the estop is not set
    If Not myDrillTap.ESTOPset Then
        
        'If in active maintenance mode, disable
        If frmMaintenance.frameMaintIO.Enabled Then
            'Disable active maintenance mode
            myDrillTap.maintActive 0, mActiveOff
        Else
        
            If myDrillTap.isJoyOn Then
                'If the joystick is on, clear joystick
                myFSM.setJoystick 0
            Else
                'Call the FSM to go to the run state
                myFSM.setGoStop
            End If
        End If
        
        
    'If estip is set
    Else
        'Query the user to reset the ESTOP flag
        Dim messageReturn As VbMsgBoxResult
        messageReturn = MsgBox("Check machine before clearing ESTOP." & vbCrLf & vbCrLf & "Clear Estop?", vbOKCancel)
        
        'If the user acknowledges the reset
        If messageReturn = vbOK Then
        
            'Call the FSM to clear the ESTOP
            myFSM.clearESTOP
            
        End If
    End If
    
End Sub


Private Sub btnState(newState As btnState)

    Select Case newState
        Case btnInit
        
            frmUI.cmdGO.Caption = "GO"
            frmUI.cmdGO.Enabled = False
            
            frmUI.frameSetup.Enabled = False
            frmUI.frameDRO.Enabled = False
            frmUI.frameToolCount.Enabled = False
            frmMaintenance.frameMaintIO.Enabled = False
            frmMaintenance.frameMaintPos.Enabled = False
            
            frmUI.topbarJoy.Enabled = False
            frmUI.topbarMaint.Enabled = False
            frmUI.topbarSet0.Enabled = False
            frmMaintenance.topbarEnIO.Enabled = True
            
        Case btnStby
        
            frmUI.cmdGO.Caption = "GO"
            frmUI.cmdGO.Enabled = True
            
            frmUI.frameSetup.Enabled = True
            frmUI.frameDRO.Enabled = True
            frmUI.frameToolCount.Enabled = True
            frmMaintenance.frameMaintIO.Enabled = False
            frmMaintenance.frameMaintPos.Enabled = False

            frmUI.topbarJoy.Enabled = True
            frmUI.topbarMaint.Enabled = True
            frmUI.topbarSet0.Enabled = True
            frmMaintenance.topbarEnIO.Enabled = True
            
        Case btnRunning
        
            frmUI.cmdGO.Caption = "STOP"
            frmUI.cmdGO.Enabled = True
            
            frmUI.frameSetup.Enabled = False
            frmUI.frameDRO.Enabled = False
            frmUI.frameToolCount.Enabled = False
            frmMaintenance.frameMaintIO.Enabled = False
            frmMaintenance.frameMaintPos.Enabled = False
            
            frmUI.topbarJoy.Enabled = False
            frmUI.topbarMaint.Enabled = False
            frmUI.topbarSet0.Enabled = False
            frmMaintenance.topbarEnIO.Enabled = True
            
        Case btnJoystick
        
            frmUI.cmdGO.Caption = "STOP JOYSTICK"
            frmUI.cmdGO.Enabled = True
            
            frmUI.frameSetup.Enabled = False
            frmUI.frameDRO.Enabled = False
            frmUI.frameToolCount.Enabled = False
            frmMaintenance.frameMaintIO.Enabled = False
            frmMaintenance.frameMaintPos.Enabled = False
            
            frmUI.topbarJoy.Enabled = False
            frmUI.topbarMaint.Enabled = True
            frmUI.topbarSet0.Enabled = True
            frmMaintenance.topbarEnIO.Enabled = True
            
        Case btnESTOP

            frmUI.cmdGO.Caption = "CLEAR ESTOP"
            frmUI.cmdGO.Enabled = True
            
            frmUI.frameSetup.Enabled = False
            frmUI.frameDRO.Enabled = False
            frmUI.frameToolCount.Enabled = False
            frmMaintenance.frameMaintIO.Enabled = False
            frmMaintenance.frameMaintPos.Enabled = False
            
            frmUI.topbarJoy.Enabled = False
            frmUI.topbarMaint.Enabled = False
            frmUI.topbarSet0.Enabled = False
            frmMaintenance.topbarEnIO.Enabled = True
            
        Case btnMaintOn

            frmUI.cmdGO.Caption = "DISABLE I/O"
            frmUI.cmdGO.Enabled = True
            
            frmUI.frameSetup.Enabled = False
            frmUI.frameDRO.Enabled = False
            frmUI.frameToolCount.Enabled = False
            frmMaintenance.frameMaintIO.Enabled = True
            frmMaintenance.frameMaintPos.Enabled = True
            
            frmUI.topbarJoy.Enabled = True
            frmUI.topbarMaint.Enabled = False
            frmUI.topbarSet0.Enabled = True
            frmMaintenance.topbarEnIO.Enabled = False
            
        Case Else
    End Select

End Sub




Public Sub updStatus(myStatus As uiStatus, strInput As String)

    Dim myMsg As String

    Select Case myStatus
        Case uiLoading
        
            myMsg = "System Loading, Please Wait"
            
            btnState btnInit
            
        Case uiStandby
        
            myMsg = "Waiting for Input"
            
            btnState btnStby
            
        Case uiWaiting
        
            myMsg = "Motion in Progress, Please Wait"
            
            btnState btnRunning
            
        Case uiJoystick
        
            myMsg = "Joystick Enabled"
        
            btnState btnJoystick
            
        Case uiClampOFF
        
            myMsg = "Clamp Off - Load Shank before pressing GO"
        
            btnState btnStby
        
        Case uiClampOFFcycle
            
            myMsg = "Clamp Off - Remove Shank and Reload"
        
            btnState btnRunning
            
        Case uiESTOP
        
            myMsg = "ESTOP Triggered - Check Machine before Clearing"
            
            btnState btnESTOP
            
        Case uiRunningDrill
        
            myMsg = "Running Drill Cycle"
            
            btnState btnRunning
            
        Case uiRunningTap
        
            myMsg = "Running Tap Cycle"
            
            btnState btnRunning
            
        Case uiMaintIdle
            myMsg = "Maintenance Window Activated - Press Enable I/O or Run As Normal"
            
            btnState btnStby
            
        Case uiMaintEnabled
            myMsg = "Maintenance Window Activated - Normal Functions Disabled"
            
            btnState btnMaintOn
            
        Case uiError
        
            myMsg = "ERROR: " & strInput
            
            btnState btnStby
            
    End Select

    frmUI.lblSystemStatus.Caption = myMsg
    frmUI.Refresh
    frmMaintenance.Refresh

End Sub

Public Sub updDRO()

Dim currentPosition As myCoordinate

currentPosition.X = myDrillTap.getCoords(cX)
currentPosition.Y = myDrillTap.getCoords(cY)

frmUI.lblDRO(0).Caption = "X:     " & Format(currentPosition.X, "###0.000") & " in"
frmUI.lblDRO(1).Caption = "Y:     " & Format(currentPosition.Y, "###0.000") & " in"

End Sub

Private Sub loadCycles()

    'If a previously saved drill / tap cycle file exists, then load data
    If Dir(savePath) <> "" Then
        
        'Load previous data
        Open savePath For Input As #1
        Input #1, drillCycles, tapCycles, drillAvg, drillAvgCy, tapAvg, tapAvgCy
        Close #1
        
        updCycles
            
    'If the file doesn't exist, ensure variables are cleared.
    Else
        
        drillCycles = 0
        tapCycles = 0
        drillAvg = 0
        drillAvgCy = 0
        tapAvg = 0
        tapAvgCy = 0
    
    End If

End Sub

Private Sub updCycles()

    'Update the UI
    frmUI.lblDrillTapCount(0).Caption = "Drill:     " & Str(drillCycles)
    frmUI.lblDrillTapCount(0).Caption = "Tap :     " & Str(tapCycles)
    
    'Save to file
    Open savePath For Output As #1
    Write #1, drillCycles, tapCycles, drillAvg, drillAvgCy, tapAvg, tapAvgCy
    Close #1

End Sub

Public Sub increment(myTool As uiCounter)
    
    'Increment counters
    Select Case myTool
        Case countDrill, countBoth
            drillCycles = drillCycles + 1
        Case countTap, countBoth
            tapCycles = tapCycles + 1
        Case Else
            Exit Sub
    End Select
    
    'Update UI
    updCycles
    
End Sub

Public Sub resetCounter(resetTool As uiCounter)
    
    'Set up variables to the appropriate tool
    Dim toolStr As String
    Dim myAvg As Double
    Dim myCounter As Integer
    Dim myCycles As Integer
    
    Select Case uiCounter
        Case countDrill
            myCycles = drillCycles
            myAvg = drillAvg
            myCounter = drillAvgCy
            toolStr = "Drill"
        Case countTap
            myCycles = tapCycles
            myAvg = tapAvg
            myCounter = tapAvgCy
            toolStr = "Tap"
        Case Else
            Exit Sub
    End Select
    
    'Verify that the user wants to reset
    Dim boxReturn As VbMsgBoxResult
    boxReturn = MsgBox(toolStr & " has been used for " & Str(myCycles) & " Cycles." & vbCrLf & "Average is " & Format(myAvg, "###0.0") & " Cycles. Reset Counter?", vbYesNo, "Reset Drill Counter?")
    
    'If the user confirms the reset
    If boxReturn = vbYes Then
        
        'Compute the rolling average
        myAvg = ((myAvg * (myCounter)) / (myCounter + 1)) + (myCycles / (myCounter + 1))
        
        
        'Reset the counter
        Select Case uiCounter
            Case countDrill
                drillCycles = 0
                drillAvg = myAvg
                drillAvgCy = myCounter + 1
            Case countTap
                tapCycles = 0
                tapAvg = myAvg
                tapAvgCy = myCounter + 1
            Case Else
                Exit Sub
        End Select
        
        'Update UI
        updCycles
        
    End If
End Sub

'---===---===---=== Advanced Functions -  ---===---===---===---===---===---===---==
'---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---===---=



