Attribute VB_Name = "terminal"
'============================================================================
'
' FILE:       TERMINAL.BAS
'
' PURPOSE:    Global declarations for VB example using Com6srvr.exe
'
' COPYRIGHT:  Copyright(c) 1992-2000, All Rights Reserved
'             Parker Hannifin Corporation
'             Compumotor Corporation
'             5500 Business Park Drive
'             Rohnert Park, California 94928
'             Applications Engineering:  (800)358-9070
'
' DISCLAIMER: THIS SOFTWARE IS PROVIDED FREE OF CHARGE AND WITHOUT
'             WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED.  IN NO
'             EVENT WILL PARKER HANNIFIN CORPORATION BE LIABLE FOR ANY
'             DAMAGES, INCLUDING ANY LOST PROFITS, LOST SAVINGS, OR OTHER
'             INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE USE
'             OR INABILITY TO USE THIS SOFTWARE.
'
' NOTE:       Additional information regarding the Com6srvr.exe driver is
'             avaliable in the readme.txt file in the Com6srvr directory.
'
'============================================================================

Option Explicit

'------------------------------------------------------------------------------------------------------------------------------------------
' Faststatus Information Data Structure
'------------------------------------------------------------------------------------------------------------------------------------------
Type FastStatusInfo
    UpdateID As Integer               ' Reserved for internal use
    Counter As Integer                ' time frame counter (2ms per count)
    MotorPos(1 To 8) As Long          ' commanded position (counts)
    EncoderPos(1 To 8) As Long        ' actual position (counts)
    MotorVel(1 To 8) As Long          ' commanded velocity (counts/sec)
    AxisStatus(1 To 8) As Long        ' axis status (TAS)
    SysStatus As Long                 ' system status (TSS)
    ErrorStatus As Long               ' user status (TER)
    UserStatus As Long                ' user status (TUS)
    Timer As Long                     ' timer value (TIM - milliseconds)
    Limits As Long                    ' limit status (TLIM)
    ProgIn(0 To 3) As Long            ' programmable input status (TIN)
    ProgOut(0 To 3) As Long           ' programmable output status (TOUT)
    Triggers As Long                  ' trigger interrupt status (TTRIG)
    Analog(1 To 2) As Integer         ' lo-res analog input voltage (TANV)
    VarB(1 To 10) As Long             'VARB1 - VARB10
    VarI(1 To 10) As Long             'VARI1 - VARI10
    Reserved As Long                  'Reserved for internal use
    CmdCount As Long                  'Command Count (from communications port)
End Type


'------------------------------------------------------------------------------------------------------------------------------------------
' Win API Function declarations
'------------------------------------------------------------------------------------------------------------------------------------------
Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal numbytes As Long)


'------------------------------------------------------------------------------------------------------------------------------------------
' Global variables
'------------------------------------------------------------------------------------------------------------------------------------------
Global c6k As Object            ' comm server object (RS232 or Ethernet)
Global fsinfo As FastStatusInfo 'fast status information
Global connected As Boolean     ' flag to indicate connection state - could also use the state of timer1.enabled for this
Global alarms As Long 'alarm status

'Control Classes
Global myDrillTap As New ClassDrillTap
Global myUI As New ClassUI
Global myFSM As New ClassFSMDrillTap

Public Const Input9 As Long = (2 ^ (9 - 1))
Public Const Input10 As Long = (2 ^ (10 - 1))
Public Const Input11 As Long = (2 ^ (11 - 1))
Public Const Input12 As Long = (2 ^ (12 - 1))
Public Const Input13 As Long = (2 ^ (13 - 1))
Public Const Input14 As Long = (2 ^ (14 - 1))
Public Const Input15 As Long = (2 ^ (15 - 1))
Public Const Input16 As Long = (2 ^ (16 - 1))
Public Const Input17 As Long = (2 ^ (17 - 1))
Public Const Input18 As Long = (2 ^ (18 - 1))
Public Const Input19 As Long = (2 ^ (19 - 1))
Public Const Input20 As Long = (2 ^ (20 - 1))
Public Const Input21 As Long = (2 ^ (21 - 1))
Public Const Input22 As Long = (2 ^ (22 - 1))
Public Const Input23 As Long = (2 ^ (23 - 1))
Public Const Input24 As Long = (2 ^ (24 - 1))
Public Const Drive1 As Long = (2 ^ (13 - 1))
Public Const JoyStick As Long = (2 ^ (9 - 1))
Public Input23Mask
Public Pump As Boolean
Public Last_Pcut_State As Integer

Public Type myCoordinate
    X As Double
    Y As Double
End Type




Function BitText32(v As Long) As String
' this fucntion takes a long value and returns a string
' representing the 32 bit binary value LSB->MSB left to right
       
' this code is less compact than the BitText32Ex function but takes
' only .15 milliseconds to run
    
    Dim temp$
    temp = temp & (v And 1) \ 1
    temp = temp & (v And 2) \ 2
    temp = temp & (v And 4) \ 4
    temp = temp & (v And 8) \ 8
    temp = temp & "_"
    temp = temp & (v And 16) \ 16
    temp = temp & (v And 32) \ 32
    temp = temp & (v And 64) \ 64
    temp = temp & (v And 128) \ 128
    temp = temp & "_"
    temp = temp & (v And 256) \ 256
    temp = temp & (v And 512) \ 512
    temp = temp & (v And 1024) \ 1024
    temp = temp & (v And 2048) \ 2048
    temp = temp & "_"
    temp = temp & (v And 4096) \ 4096
    temp = temp & (v And 8192) \ 8192
    temp = temp & (v And 16384) \ 16384
    temp = temp & (v And 32768) \ 32768
    temp = temp & "_"
    temp = temp & (v And 65536) \ 65536
    temp = temp & (v And 131072) \ 131072
    temp = temp & (v And 262144) \ 262144
    temp = temp & (v And 524288) \ 524288
    temp = temp & "_"
    temp = temp & (v And 1048572) \ 1048572
    temp = temp & (v And 2097152) \ 2097152
    temp = temp & (v And 4194304) \ 4194304
    temp = temp & (v And 8388608) \ 8388608
                  If (v And 8388608) \ 8388608 = 1 Then
                  Input23Mask = "24 On"
                  Else
                  Input23Mask = "24 Off"
                  End If
    temp = temp & "_"
    temp = temp & (v And 16777216) \ 16777216
    temp = temp & (v And 33554432) \ 33554432
    temp = temp & (v And 67108864) \ 67108864
    temp = temp & (v And 134217728) \ 134217728
    temp = temp & "_"
    temp = temp & (v And 268435456) \ 268435456
    temp = temp & (v And 536870912) \ 536870912
    temp = temp & (v And 1073741824) \ 1073741824
    If v < 0 Or v >= 2147483647 Then
        temp = temp & "1"
    Else
        temp = temp & "0"
    End If
    BitText32 = temp

End Function


Sub Disconnect()
'global error routine
'ensure that commserver objects are released on any unexpected errors.

    Dim msg$
    
    msg = "An unexpected error was encountered." & vbCr
    msg = msg & "Closing down the application." & vbCr & vbCr
    msg = msg & "Error: " & CStr(Err) & vbCr
    msg = msg & Error$
    Call MsgBox(msg, 0, "Unexpected Error")     'display error message
    frmMain!Timer1.Enabled = False              'diable timer
    Unload frmFastStatus                        'unload status display if loaded
    Set c6k = Nothing                           'release the commserver
    End
End Sub

Function BitText32Ex(v As Long) As String
' this fucntion takes a long value and returns a string
' representing the 32 bit binary value LSB->MSB left to right
       
' this code is more compact than the BitText32 function but takes
' roughly .31 milliseconds to run - about double that of BitText32
    
    Dim temp$
    Dim n%, i%, mask&
    
    i = 0
    For n = 0 To 30
        i = i + 1
        mask = 2 ^ n
        temp = temp & CStr((v And mask) \ mask)
        If i = 4 Then
            temp = temp & "_"
            i = 0
        End If
    Next 'n
    
    If v >= 2147483647 Then
        temp = temp & "1"
    Else
        temp = temp & "0"
    End If
    BitText32Ex = temp
    
End Function




Public Sub JoyStickOn()
c6k.Write "1ANIRNG.25=3:1INFNC23-O:1INFNC24-M:JOYAXH1-25:JOYVH50,:JOYVL10,:JOYA100:" & Chr$(13)
c6k.Write "1JOYEDB.25=2:1JOYZ.25=1:JOY11" & Chr$(13) 'send commands to 6k
End Sub
