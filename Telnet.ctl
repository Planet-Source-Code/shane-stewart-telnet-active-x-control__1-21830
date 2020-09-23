VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Telnet 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleMode       =   0  'User
   ScaleWidth      =   420
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4080
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Telnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'This code was writen by and is the property of Shane Stewart
'Copyright 2000 by Shane Stewart
'I grant permission to reuse this code in any manner you wish

'Please give credit to the author when using this code
'The constants and enum type iacoption were code borrowed from Ian Storrs
'       but has been significanty altered
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'I made this because I could not find any Telnet implementations that are
'compliant with rfc854 in the way that they negotiate options. I used rfc1143
'as the basis for this implementation.

'I welcome any comments on this code any any sugestions on how to improve
'the implementation or features to add.

'I would love to see the things done with this code or features added to it.
'Please send ideas and your projects to sstewart@networld.com
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'This control presents several events and properties to the developer. Some of
'these events must be used in step with a property for proper functionality.
'Please see the documentation for more details.

Option Explicit

Enum Qstate
qNo = 0
qYes = 1
qWantno = 3
qWantyes = 4
End Enum

Enum OptState
OptDisable = 0
OptEnable = 1
End Enum

Enum Queue
qEmpty = 0
qOpposite = 1
End Enum

Enum OptWaiting
WaitingNo = 0
WaitingYes = 1
End Enum

Private Type rfc1143
Us As Qstate
Him As Qstate
Usq As Queue
Himq As Queue
UsWaiting As OptWaiting
HimWaiting As OptWaiting
End Type

Enum IACOption
BINARYt = 0                 'rfc856
ECHO = 1                    'rfc857
RECONNECT = 2
SGA = 3                     'rfc858
AMSN = 4
STATUSt = 5                 'rfc859
TIMINGMARK = 6              'rfc860
RCTAN = 7
OLW = 8
OPS = 9
OCRD = 10
OHTS = 11
OHTD = 12
OFFD = 13
OVTS = 14
OVTD = 15
OLFD = 16
XASCII = 17
LOGOUT = 18
BYTEM = 19
DET = 20
SUPDUP = 21
SUPDUPOUT = 22
SENDLOC = 23
TERMTYPE = 24               'rfc1091
EOR = 25
TACACSUID = 26
OUTPUTMARK = 27
TERMLOCNUM = 28
REGIME3270 = 29
X3PAD = 30
NAWS = 31
TERMSPEED = 32
TFLOWCNTRL = 33
LINEMODE = 34
DISPLOC = 35
ENVIRONt = 36               'rfc1408
AUTHENTICATION = 37
NEWENVIRON = 39             'rfc1572
EXTENDED_OPTIONS_LIST = 255
End Enum

Dim IACState(255) As rfc1143
Dim TermData As String
Dim TimeToWait As Double
Dim DoneSending As Boolean
Dim Processing As Boolean
Dim Sending As Boolean
Dim InQueue As New Collection
Dim OutQueue As New Collection
Dim DataLength As Long

Const SUSP = 237
Const ABORT = 238      'Abort
Const SE = 240         'End of Subnegotiation
Const NOP = 241        'No Opperation
Const DM = 242         'Data Mark
Const BREAK = 243      'BREAK
Const IP = 244         'Interrupt Process
Const AO = 245         'Abort Output
Const AYT = 246        'Are you there
Const EC = 247         'Erase Character
Const EL = 248         'Erase Line
Const GOAHEAD = 249    'Go Ahead
Const SB = 250         'What follows is subnegotiation
Const IACWILL = 251    'WILL
Const IACWONT = 252    'WONT
Const IACDO = 253      'DO
Const IACDONT = 254    'DONT
Const IAC = 255        'Is A Command

Public Event HimStateChange(ByVal OptionNumber As IACOption)
Public Event UsStateChange(ByVal OptionNumber As IACOption)
Public Event HeRequestHimEnable(ByVal OptionNumber As IACOption)
Public Event HeRequestUsEnable(ByVal OptionNumber As IACOption)
Public Event SubNegotiationRecieved(SBOption As String)

Public Event RecievedSuspend()
Public Event RecievedAbort()
Public Event RecievedDataMark()
Public Event RecievedBreak()
Public Event RecievedInterruptProcess()
Public Event RecievedAbortOutput()
Public Event RecievedAreYouThere()
Public Event RecievedEraseCharacter()
Public Event RecievedEraseLine()
Public Event RecievedGoAhead()

Public Event DebugMsgRecieve(ByVal Message As String)
Public Event TermDataArive(ByVal DataLength As Long)

Public Event ConnectionClosed()
Public Event ConnectionOpened()

Private Sub IAC_RecieveCommand(ByVal IACCommand As Integer, ByVal OptionNumber As Integer)

Select Case IACCommand
Case IACWILL
RaiseEvent DebugMsgRecieve("Recieved WILL: " & OptionNumber)
'Upon receipt of WILL, we choose based upon him and himq:
         'NO            If we agree that he should enable, him=YES, send
         '              DO; otherwise, send DONT.
         'YES           Ignore.
         'WANTNO  EMPTY Error: DONT answered by WILL. him=NO.
         '     'OPPOSITE Error: DONT answered by WILL. him=YES*,
         '              himq=EMPTY.
         'WANTYES EMPTY him=YES.
         '     OPPOSITE him=WANTNO, himq=EMPTY, send DONT.
    Select Case IACState(OptionNumber).Him
        Case qNo
            IACState(OptionNumber).HimWaiting = WaitingYes
            RaiseEvent HeRequestHimEnable(OptionNumber)
            Pause (TimeToWait)
            If IACState(OptionNumber).HimWaiting = WaitingYes Then
            IACState(OptionNumber).HimWaiting = WaitingNo
            Call IAC_SendDONT(OptionNumber)
            RaiseEvent DebugMsgRecieve("Option timed out WILL: " & OptionNumber)
            End If
        Case qYes
            RaiseEvent DebugMsgRecieve("recieve WILL when him = qyes")
        Case qWantno
            If IACState(OptionNumber).Himq = qEmpty Then
            IACState(OptionNumber).Him = qNo
            RaiseEvent HimStateChange(OptionNumber)
            Else
            IACState(OptionNumber).Him = qYes
            IACState(OptionNumber).Himq = qEmpty
            RaiseEvent HimStateChange(OptionNumber)
            End If
        Case qWantyes
            If IACState(OptionNumber).Himq = qEmpty Then
            IACState(OptionNumber).Him = qYes
            RaiseEvent HimStateChange(OptionNumber)
            Else
            IACState(OptionNumber).Himq = qEmpty
            IACState(OptionNumber).Him = qWantno
            End If
    End Select

Case IACWONT
RaiseEvent DebugMsgRecieve("Recieved WONT: " & OptionNumber)
'Upon receipt of WONT, we choose based upon him and himq:
         'NO            Ignore.
         'YES           him=NO, send DONT.
         'WANTNO  EMPTY him=NO.
         '     OPPOSITE him=WANTYES, himq=EMPTY, send DO.
         'WANTYES EMPTY him=NO.*
         '     OPPOSITE him=NO, himq=EMPTY.**
    Select Case IACState(OptionNumber).Him
        Case qNo
            'Ignore
        Case qYes
            IACState(OptionNumber).Him = qNo
            Call IAC_SendDONT(OptionNumber)
            RaiseEvent HimStateChange(OptionNumber)
        Case qWantno
            If IACState(OptionNumber).Himq = qEmpty Then
                IACState(OptionNumber).Him = qNo
                RaiseEvent HimStateChange(OptionNumber)
            Else
                IACState(OptionNumber).Him = qWantyes
                IACState(OptionNumber).Himq = qEmpty
                Call IAC_SendDO(OptionNumber)
            End If
        Case qWantyes
            If IACState(OptionNumber).Himq = qEmpty Then
                IACState(OptionNumber).Him = qNo
                RaiseEvent HimStateChange(OptionNumber)
            Else
                IACState(OptionNumber).Him = qNo
                IACState(OptionNumber).Himq = qEmpty
                RaiseEvent HimStateChange(OptionNumber)
            End If
    End Select

Case IACDO
RaiseEvent DebugMsgRecieve("Recieved DO: " & OptionNumber)
'Upon receipt of DO, we choose based upon us and usq:
         'NO            If we agree that we should enable, us=YES, send
         '              WILL; otherwise, send WONT.
         'YES           Ignore.
         'WANTNO  EMPTY Error: WONT answered by DO. us=NO.
         '     'OPPOSITE Error: WONT answered by DO. us=YES*,
         '              usq=EMPTY.
         'WANTYES EMPTY us=YES.
         '     OPPOSITE us=WANTNO, usq=EMPTY, send WONT.
    Select Case IACState(OptionNumber).Us
        Case qNo
            IACState(OptionNumber).UsWaiting = WaitingYes
            RaiseEvent HeRequestUsEnable(OptionNumber)
            Pause (TimeToWait)
            If IACState(OptionNumber).UsWaiting = WaitingYes Then
                IACState(OptionNumber).UsWaiting = WaitingNo
                Call IAC_SendWONT(OptionNumber)
                RaiseEvent DebugMsgRecieve("Option timed out DO: " & OptionNumber)
            End If
        Case qYes
            RaiseEvent DebugMsgRecieve("error: Already enabled " & OptionNumber)
        Case qWantno
            If IACState(OptionNumber).Usq = qEmpty Then
                IACState(OptionNumber).Us = qNo
                RaiseEvent UsStateChange(OptionNumber)
            Else
                IACState(OptionNumber).Us = qYes
                    IACState(OptionNumber).Usq = qEmpty
            RaiseEvent UsStateChange(OptionNumber)
            End If
        Case qWantyes
            If IACState(OptionNumber).Usq = qEmpty Then
                IACState(OptionNumber).Us = qYes
                RaiseEvent UsStateChange(OptionNumber)
            Else
                IACState(OptionNumber).Us = qWantno
                IACState(OptionNumber).Usq = qEmpty
                Call IAC_SendWONT(OptionNumber)
            End If
    End Select

Case IACDONT
RaiseEvent DebugMsgRecieve("Recieved DONT: " & OptionNumber)
'Upon receipt of DONT, we choose based upon us and usq:
         'NO            Ignore.
         'YES           us=NO, send WONT.
         'WANTNO  EMPTY us=NO.
         '     OPPOSITE us=WANTYES, usq=EMPTY, send WILL.
         'WANTYES EMPTY us=NO.*
         '     OPPOSITE us=NO, usq=EMPTY.**
    Select Case IACState(OptionNumber).Us
        Case qNo
            RaiseEvent DebugMsgRecieve("error: already disabled " & OptionNumber)
        Case qYes
            IACState(OptionNumber).Us = qNo
            Call IAC_SendWONT(OptionNumber)
            RaiseEvent UsStateChange(OptionNumber)
        Case qWantno
            If IACState(OptionNumber).Usq = qEmpty Then
                IACState(OptionNumber).Us = qNo
                RaiseEvent UsStateChange(OptionNumber)
            Else
                IACState(OptionNumber).Us = qWantyes
                IACState(OptionNumber).Usq = qEmpty
                Call IAC_SendWILL(OptionNumber)
            End If
        Case qWantyes
            If IACState(OptionNumber).Usq = qEmpty Then
                IACState(OptionNumber).Us = qNo
                RaiseEvent UsStateChange(OptionNumber)
            Else
                IACState(OptionNumber).Us = qNo
                IACState(OptionNumber).Usq = qEmpty
                RaiseEvent UsStateChange(OptionNumber)
            End If
    End Select
    
End Select
End Sub

Private Sub UserControl_Initialize()
TimeToWait = 0.1
DoneSending = True
End Sub

Private Sub UserControl_Resize()
Size 420, 420
End Sub

Private Sub Pause(Dur As Double)
'Standard pause routine. Used Double to allow small increments.
Dim Tim1 As Double
Tim1 = Timer
Do Until Tim1 + Dur <= Timer: DoEvents: Loop
End Sub

Public Property Get UsOptionState(ByVal OptionNumber As IACOption) As OptState
If IACState(OptionNumber).Us = qYes Then
UsOptionState = OptEnable
Else
UsOptionState = OptDisable
End If
End Property

Public Property Let UsOptionState(ByVal OptionNumber As IACOption, ByVal NewState As OptState)
If Winsock1.State <> 7 Then Exit Property    'if we arent connected then ignor request
'if we were waiting for application to decide if we should enable
Dim iOptionNumber As Integer
iOptionNumber = OptionNumber
If IACState(OptionNumber).UsWaiting = WaitingYes Then 'we were waiting
    IACState(OptionNumber).UsWaiting = WaitingNo 'reset the wait state
    Select Case NewState
        Case OptDisable 'decided we should not enable
            IACState(OptionNumber).Us = qNo    'just for good measure
            Call IAC_SendWONT(iOptionNumber)
        Case OptEnable  'decided we should enable
            IACState(OptionNumber).Us = qYes   'see reciept of DO in IAC_RecieveCommand
            Call IAC_SendWILL(iOptionNumber)
            RaiseEvent UsStateChange(OptionNumber)
    End Select  'only those two choices
Else 'we were not waiting
    Select Case NewState
        Case OptDisable 'we are telling him we are going to disable
            'If we decide to ask for us to disable:
            'NO Error: Already disabled.
            'YES us=WANTNO, send WONT.
            'WANTNO EMPTY Error: Already negotiating for disable.
            '     OPPOSITE usq=EMPTY.
            'WANTYES EMPTY If we are queueing requests, usq=OPPOSITE; otherwise, Error: Cannot initiate new request in the middle of negotiation.
            '    OPPOSITE Error: Already queued a disable request.
            Select Case IACState(OptionNumber).Us
                Case qNo
                RaiseEvent DebugMsgRecieve("error: already disabled " & OptionNumber)
                Case qYes
                IACState(OptionNumber).Us = qWantno
                Call IAC_SendWONT(iOptionNumber)
                Case qWantno
                    If IACState(OptionNumber).Usq = qEmpty Then
                    RaiseEvent DebugMsgRecieve("error: already negotiating for disable " & OptionNumber)
                    Else
                    IACState(OptionNumber).Usq = qEmpty
                    End If
                Case qWantyes
                    If IACState(OptionNumber).Usq = qEmpty Then
                    IACState(OptionNumber).Usq = qOpposite
                    Else
                    RaiseEvent DebugMsgRecieve("error: already queued a disable request " & OptionNumber)
                    End If
            End Select
            
        Case OptEnable  'we are offering to enable
            'If we decide to ask for us to enable:
            'NO us=WANTYES, send WILL.
            'YES Error: Already enabled.
            'WANTNO EMPTY If we are queueing requests, usq=OPPOSITE; otherwise, Error: Cannot initiate new request in the middle of negotiation.
            '    OPPOSITE Error: Already queued an enable request.
            'WANTYES EMPTY Error: Already negotiating for enable.
            '    OPPOSITE usq=EMPTY.
            Select Case IACState(OptionNumber).Us
                Case qNo
                IACState(OptionNumber).Us = qWantyes
                Call IAC_SendWILL(iOptionNumber)
                Case qYes
                RaiseEvent DebugMsgRecieve("error: Already enabled " & OptionNumber)
                Case qWantno
                    If IACState(OptionNumber).Usq = qEmpty Then
                    IACState(OptionNumber).Usq = qOpposite
                    Else
                    RaiseEvent DebugMsgRecieve("error: Already queued an enable request " & OptionNumber)
                    End If
                Case qWantyes
                    If IACState(OptionNumber).Usq = qEmpty Then
                    RaiseEvent DebugMsgRecieve("error: Already negotiating for enable " & OptionNumber)
                    Else
                    IACState(OptionNumber).Usq = qEmpty
                    End If
            End Select
    End Select
End If
End Property

Public Property Get HimOptionState(ByVal OptionNumber As IACOption) As OptState
If IACState(OptionNumber).Him = qYes Then
HimOptionState = OptEnable
Else
HimOptionState = OptDisable
End If
End Property
Public Function GetTermData() As String
GetTermData = TermData
TermData = ""
DataLength = 0
End Function

Private Sub SendOut()

Dim Temp As String

Do While OutQueue.Count > 0
Temp = OutQueue.Item(1)
OutQueue.Remove (1)
If Winsock1.State = 7 Then
Do Until DoneSending = True: DoEvents: Loop
DoneSending = False
Winsock1.SendData Temp
RaiseEvent DebugMsgRecieve("Sent " & Len(Temp) & " bytes of data")
Else

RaiseEvent DebugMsgRecieve(Len(Temp) & " bytes not sent because host is disconnected")
End If
DoEvents
Loop
Sending = False
End Sub

Public Property Let HimOptionState(ByVal OptionNumber As IACOption, ByVal NewState As OptState)
If Winsock1.State <> 7 Then Exit Property    'if we arent connected then ignor request
'if we were waiting for application to decide if he should enable
Dim iOptionNumber As Integer
iOptionNumber = OptionNumber
If IACState(OptionNumber).HimWaiting = WaitingYes Then 'we were waiting
    IACState(OptionNumber).HimWaiting = WaitingNo 'reset the wait state
    Select Case NewState
        Case OptDisable 'decided he should not enable
            IACState(OptionNumber).Him = qNo    'just for good measure
            Call IAC_SendDONT(iOptionNumber)
        Case OptEnable  'decided he should enable
            IACState(OptionNumber).Him = qYes   'see reciept of will in IAC_RecieveCommand
            Call IAC_SendDO(iOptionNumber)
            RaiseEvent HimStateChange(OptionNumber)
    End Select  'only those two choices
Else 'we were not waiting
    Select Case NewState
        Case OptDisable 'we are asking him to disable
            'If we decide to ask him to disable:
            'NO Error: Already disabled.
            'YES him=WANTNO, send DONT.
            'WANTNO EMPTY Error: Already negotiating for disable.
            '    OPPOSITE himq=EMPTY.
            'WANTYES EMPTY If we are queueing requests, himq=OPPOSITE; otherwise, Error: Cannot initiate new request in the middle of negotiation.
            '    OPPOSITE Error: Already queued a disable request.
            Select Case IACState(OptionNumber).Him
                Case qNo
                RaiseEvent DebugMsgRecieve("Error: Already disabled " & OptionNumber)
                Case qYes
                IACState(OptionNumber).Him = qWantno
                Call IAC_SendDONT(iOptionNumber)
                Case qWantno
                    If IACState(OptionNumber).Himq = qEmpty Then
                    RaiseEvent DebugMsgRecieve("error: Already negotiating for disable " & OptionNumber)
                    Else
                    IACState(OptionNumber).Himq = qEmpty
                    End If
                Case qWantyes
                    If IACState(OptionNumber).Himq = qEmpty Then
                    IACState(OptionNumber).Himq = qOpposite
                    Else
                    RaiseEvent DebugMsgRecieve("error: Already queued a disable request " & OptionNumber)
                    End If
            End Select
            
        Case OptEnable  'we are asking him to enable
            'If we decide to ask him to enable:
            'NO him=WANTYES, send DO.
            'YES Error: Already enabled.
            'WANTNO EMPTY If we are queueing requests, himq=OPPOSITE; otherwise, Error: Cannot initiate new request in the middle of negotiation.
            '   OPPOSITE Error: Already queued an enable request.
            'WANTYES EMPTY Error: Already negotiating for enable.
            '   OPPOSITE himq=EMPTY.
            Select Case IACState(OptionNumber).Him
                Case qNo
                IACState(OptionNumber).Him = qWantyes
                Call IAC_SendDO(iOptionNumber)
                Case qYes
                RaiseEvent DebugMsgRecieve("error: Already enabled " & OptionNumber)
                Case qWantno
                    If IACState(OptionNumber).Himq = qEmpty Then
                    IACState(OptionNumber).Himq = qOpposite
                    Else
                    RaiseEvent DebugMsgRecieve("error: Already queued an enable request " & OptionNumber)
                    End If
                Case qWantyes
                    If IACState(OptionNumber).Himq = qEmpty Then
                    RaiseEvent DebugMsgRecieve("error: Already negotiating for enable " & OptionNumber)
                    Else
                    IACState(OptionNumber).Himq = qEmpty
                    End If
            End Select
    End Select
End If
End Property

Private Sub IAC_SendDO(OptionNumber As Integer)
RaiseEvent DebugMsgRecieve("Sent DO " & OptionNumber)
OutQueue.Add Chr$(IAC) & Chr$(IACDO) & Chr$(OptionNumber)
If Not Sending Then
Sending = True
SendOut
End If
End Sub

Private Sub IAC_SendDONT(OptionNumber As Integer)
RaiseEvent DebugMsgRecieve("Sent DONT " & OptionNumber)
OutQueue.Add Chr$(IAC) & Chr$(IACDONT) & Chr$(OptionNumber)
If Not Sending Then
Sending = True
SendOut
End If
End Sub

Private Sub IAC_SendWILL(OptionNumber As Integer)
RaiseEvent DebugMsgRecieve("Sent WILL " & OptionNumber)
OutQueue.Add Chr$(IAC) & Chr$(IACWILL) & Chr$(OptionNumber)
If Not Sending Then
Sending = True
SendOut
End If
End Sub

Private Sub IAC_SendWONT(OptionNumber As Integer)
RaiseEvent DebugMsgRecieve("Sent WONT " & OptionNumber)
OutQueue.Add Chr$(IAC) & Chr$(IACWONT) & Chr$(OptionNumber)
If Not Sending Then
Sending = True
SendOut
End If
End Sub

Public Function SendSuspend()
OutQueue.Add Chr$(IAC) & Chr$(SUSP)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendAbort()
OutQueue.Add Chr$(IAC) & Chr$(ABORT)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendDataMark()
OutQueue.Add Chr$(IAC) & Chr$(DM)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendBreak()
OutQueue.Add Chr$(IAC) & Chr$(BREAK)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendInterruptProcess()
OutQueue.Add Chr$(IAC) & Chr$(IP)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendAbortOutput()
OutQueue.Add Chr$(IAC) & Chr$(AO)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendAreYouThere()
OutQueue.Add Chr$(IAC) & Chr$(AYT)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendEraseCharacter()
OutQueue.Add Chr$(IAC) & Chr$(EC)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendEraseLine()
OutQueue.Add Chr$(IAC) & Chr$(EL)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendGoAhead()
OutQueue.Add Chr$(IAC) & Chr$(GOAHEAD)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SubNegotiationSend(ByVal SBOption As String)
OutQueue.Add Chr$(IAC) & Chr$(SB) & SBOption & Chr$(IAC) & Chr$(SE)
If Not Sending Then
Sending = True
SendOut
End If
End Function

Public Function SendData(ByVal DataToSend As String)
OutQueue.Add DataToSend
If Not Sending Then
Sending = True
SendOut
End If
End Function

Private Sub Winsock1_Close()
Winsock1.Close
RaiseEvent ConnectionClosed
CleanUp
RaiseEvent DebugMsgRecieve("Forced disconnect from " & Winsock1.RemoteHostIP)
End Sub

Private Sub Winsock1_Connect()
RaiseEvent ConnectionOpened
RaiseEvent DebugMsgRecieve("Connected to " & Winsock1.RemoteHostIP)
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)

Dim BytesIn As String

Winsock1.GetData BytesIn, vbString, BytesTotal
InQueue.Add BytesIn

If Not Processing Then
Processing = True
ProcessInData
End If

End Sub

Private Sub ProcessInData()

Dim BufferIndex As Integer
Dim ByteBuffer() As Byte
Dim ProcessingChar As Integer
Dim ProcessingCommand As Integer
Dim IsSB As Boolean
Dim SBOption As String
Dim Temp As String
Dim x As Integer
Dim BytesTotal As Long
Dim OptionNumber As Integer

Do While InQueue.Count > 0

    Temp = InQueue.Item(1)
    InQueue.Remove (1)
    BytesTotal = Len(Temp)
    ReDim ByteBuffer(BytesTotal)
    For x = 1 To BytesTotal
        ByteBuffer(x) = AscB(Mid(Temp, x, 1))
    Next x

    RaiseEvent DebugMsgRecieve("There are " & InQueue.Count & " packets in the queue")
    RaiseEvent DebugMsgRecieve("Size of packet is: " & BytesTotal & " bytes")
    BufferIndex = 1
    Do While BufferIndex <= BytesTotal
        ProcessingChar = ByteBuffer(BufferIndex)
        If ProcessingChar = IAC Then
            BufferIndex = BufferIndex + 1
            ProcessingCommand = ByteBuffer(BufferIndex)
        
            Select Case ProcessingCommand

                Case IACWILL, IACWONT, IACDO, IACDONT
                    BufferIndex = BufferIndex + 1
                    OptionNumber = ByteBuffer(BufferIndex)
                    IAC_RecieveCommand ProcessingCommand, OptionNumber

                Case SUSP
                RaiseEvent RecievedSuspend

                Case ABORT
                RaiseEvent RecievedAbort

                Case SE
                    IsSB = False
                    'put code to process sb here
                    RaiseEvent DebugMsgRecieve("Recieved sub negotiation: " & SBOption)
                    RaiseEvent SubNegotiationRecieved(SBOption)
                
                Case NOP
                'no operation

                Case DM
                RaiseEvent RecievedDataMark

                Case BREAK
                RaiseEvent RecievedBreak

                Case IP
                RaiseEvent RecievedInterruptProcess

                Case AO
                RaiseEvent RecievedAbortOutput

                Case AYT
                RaiseEvent RecievedAreYouThere

                Case EC
                RaiseEvent RecievedEraseCharacter

                Case EL
                RaiseEvent RecievedEraseLine

                Case GOAHEAD
                RaiseEvent RecievedGoAhead

                Case SB
                    IsSB = True
                    SBOption = ""

                Case IAC
                    'if its another IAC then its really a character
                    If IACState(1).Us = qYes Then OutQueue.Add Chr$(ProcessingChar)
                    TermData = TermData & Chr$(ProcessingChar)
                    DataLength = DataLength + 1
                
                Case Else
                'code goes here

            End Select

        Else
            'we are in SB mode
            If IsSB = True Then
                'will probably want to change type of SBOption and code here
                SBOption = SBOption & ProcessingChar & " "
            Else
                'it was not IAC so its a character
                If IACState(1).Us = qYes Then OutQueue.Add Chr$(ProcessingChar)
                    TermData = TermData & Chr$(ProcessingChar)
                    DataLength = DataLength + 1
            End If
        End If
        BufferIndex = BufferIndex + 1
    DoEvents
    Loop

    If DataLength > 0 Then 'just in case there are still characters in queue
        RaiseEvent TermDataArive(DataLength)
    End If
DoEvents
Loop
Processing = False
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,RemoteHost
Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "Returns/Sets the name used to identify the remote computer"
    RemoteHost = Winsock1.RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    Winsock1.RemoteHost() = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,RemoteHostIP
Public Property Get RemoteHostIP() As String
Attribute RemoteHostIP.VB_Description = "Returns the remote host IP address"
    RemoteHostIP = Winsock1.RemoteHostIP
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,RemotePort
Public Property Get RemotePort() As Long
Attribute RemotePort.VB_Description = "Returns/Sets the port to be connected to on the remote computer"
    RemotePort = Winsock1.RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Long)
    Winsock1.RemotePort() = New_RemotePort
    PropertyChanged "RemotePort"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,Connect
Public Sub Connect(Optional ByVal RemoteHost As Variant, Optional ByVal RemotePort As Variant)
Attribute Connect.VB_Description = "Connect to the remote computer"
Erase IACState

If Winsock1.State = 0 Then
If IsMissing(RemoteHost) Or IsMissing(RemotePort) Then
    Winsock1.Connect
Else
    Winsock1.Connect RemoteHost, RemotePort
End If
Else: Exit Sub
End If
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Winsock1.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    Winsock1.RemotePort = PropBag.ReadProperty("RemotePort", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("RemoteHost", Winsock1.RemoteHost, "")
    Call PropBag.WriteProperty("RemotePort", Winsock1.RemotePort, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,State
Public Property Get ConnectState() As Integer
Attribute ConnectState.VB_Description = "Returns the state of the socket connection"
    ConnectState = Winsock1.State
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Winsock1,Winsock1,-1,Close
Public Sub Disconnect()
    Winsock1.Close
CleanUp
RaiseEvent ConnectionClosed
RaiseEvent DebugMsgRecieve("We disconnected remote host " & Winsock1.RemoteHostIP)
End Sub

Private Sub Winsock1_SendComplete()
DoneSending = True
RaiseEvent DebugMsgRecieve("Winsock reports send complete")
End Sub

Private Sub CleanUp()
Erase IACState
Do Until OutQueue.Count = 0: OutQueue.Remove (1): Loop
'Do Until InQueue.Count = 0: InQueue.Remove (1): Loop
RaiseEvent DebugMsgRecieve("Clean up initialized")
End Sub
