VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "xtelnet"
   ClientHeight    =   7695
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9705
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin Xtelnet.VT100 VT1001 
      Height          =   3615
      Left            =   2400
      TabIndex        =   1
      Top             =   2040
      Width           =   4815
      _extentx        =   8493
      _extenty        =   6376
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2560
            MinWidth        =   2560
            Text            =   "Disconnected"
            TextSave        =   "Disconnected"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin Xtelnet.Telnet Telnet1 
      Left            =   9000
      Top             =   6480
      _extentx        =   741
      _extenty        =   741
   End
   Begin VB.Menu MnuConnect 
      Caption         =   "&Connect"
      Begin VB.Menu MnuRemote 
         Caption         =   "Remote System"
      End
      Begin VB.Menu MnuDisconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConnectBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConnectBar4 
         Caption         =   "-"
      End
      Begin VB.Menu host 
         Caption         =   "&1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&8"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&9"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu host 
         Caption         =   "&10"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu MnuConnectBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu MnuOption 
         Caption         =   "&Options"
      End
      Begin VB.Menu MnuClearScreen 
         Caption         =   "&Clear Screen"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HostArray(15) As String
Dim PortArray(15) As String

Private Sub Form_Load()
Dim x As Integer

Me.Width = GetSetting(App.Title, "Settings", "Width", 10500)
Me.Height = GetSetting(App.Title, "Settings", "Height", (Screen.Height - 500))
Me.Top = GetSetting(App.Title, "Settings", "Top", (((Screen.Height - Me.Height) / 2) - 250))
Me.Left = GetSetting(App.Title, "Settings", "Left", ((Screen.Width - Me.Width) / 2))
VT1001.Left = 0
VT1001.Top = 0
VT1001.Height = ScaleHeight - StatusBar.Height
VT1001.Width = ScaleWidth

Me.Caption = "Xtelnet -(" & "none" & ")"

If GetSetting(App.Title, "Settings", "VT100", 1) = "True" Then
SaveSetting App.Title, "Settings", "VT100", 0 'should be 1 when arrows work
End If
If GetSetting(App.Title, "Settings", "VT100", 1) = "False" Then
SaveSetting App.Title, "Settings", "VT100", 0
End If

VT1001.ScrollBuffer = GetSetting(App.Title, "Settings", "Buffersize", 25)

For x = 1 To 15
Hosts.Add GetSetting(App.Title, "Hosts", x, "")
Port.Add GetSetting(App.Title, "Port", x, "")
Next x

load_hostmnu

End Sub

Private Sub Form_Resize()

If ScaleWidth > 0 And ScaleHeight > 0 Then
If ScaleWidth > 12500 Then Me.Width = 12500
If ScaleWidth < 2000 Then Me.Width = 2000
If ScaleHeight < 2000 Then Me.Height = 2000
VT1001.Height = ScaleHeight - StatusBar.Height
VT1001.Top = 0
VT1001.Width = ScaleWidth

End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim x As Integer

SaveSetting App.Title, "Settings", "Buffersize", VT1001.ScrollBuffer
SaveSetting App.Title, "Settings", "Width", Me.Width
SaveSetting App.Title, "Settings", "Height", Me.Height
SaveSetting App.Title, "Settings", "Top", Me.Top
SaveSetting App.Title, "Settings", "Left", Me.Left

For x = 1 To 15
SaveSetting App.Title, "Hosts", x, Hosts(x)
SaveSetting App.Title, "Port", x, Port(x)
Next x

For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
        
        End
Next
End Sub
Public Sub load_hostmnu()
Dim y As Integer
Dim x As Integer

For x = 1 To host.Count
host(x).Visible = False
host(x).Caption = ""
Next x

y = 1
For x = 15 To 1 Step -1
If Hosts(x) <> "" And Port(x) <> "" Then
HostArray(y) = Hosts(x)
PortArray(y) = Port(x)
host(y).Visible = True
host(y).Caption = "&" & y & "  " & HostArray(y) & ":" & PortArray(y)
y = y + 1
If y > 9 Then Exit Sub
End If
Next x

End Sub

Private Sub host_Click(Index As Integer)

If Telnet1.ConnectState <> 0 Then Exit Sub
If HostArray(Index) = "" Or PortArray(Index) = "" Then Exit Sub
If PortArray(Index) < 1 Or PortArray(Index) > 65535 Then Exit Sub
FrmMain.VT1001.ClearScreen
FrmMain.Telnet1.Connect HostArray(Index), PortArray(Index)
FrmMain.MnuDisconnect.Enabled = True
FrmMain.MnuRemote.Enabled = False

End Sub

Private Sub MnuClearScreen_Click()
VT1001.ClearScreen
End Sub


Private Sub MnuDisconnect_Click()
Telnet1.Disconnect
End Sub

Private Sub MnuExit_Click()
Unload Me
End Sub

Private Sub MnuMIB_Click()
FrmMIBs.Show
End Sub

Private Sub MnuOption_Click()
frmOptions.Show
End Sub

Private Sub MnuRemote_Click()
FrmConnect.Show
End Sub

Private Sub Telnet1_ConnectionClosed()
Dim x As Integer

MnuDisconnect.Enabled = False
MnuRemote.Enabled = True

For x = 1 To host.Count
host(x).Enabled = True
Next x

StatusBar.Panels(1).Text = "Disconnected"
StatusBar.Panels(1).ToolTipText = "Not connected"
Me.Caption = "Xtelnet -(" & "none" & ")"
End Sub

Private Sub Telnet1_ConnectionOpened()
Dim x As Integer
MnuDisconnect.Enabled = True
MnuRemote.Enabled = False

For x = 1 To host.Count
host(x).Enabled = False
Next x

Telnet1.UsOptionState(SGA) = OptEnable
StatusBar.Panels(1).Text = "Connected"
StatusBar.Panels(1).ToolTipText = Telnet1.RemoteHostIP
Me.Caption = "Xtelnet -(" & Telnet1.RemoteHostIP & ")"
End Sub

Private Sub Telnet1_HeRequestHimEnable(ByVal OptionNumber As IACOption)
Select Case OptionNumber
Case ECHO
Telnet1.HimOptionState(ECHO) = OptEnable
Case SGA
Telnet1.HimOptionState(SGA) = OptEnable
Case Else
Telnet1.HimOptionState(OptionNumber) = OptDisable
End Select
End Sub

Private Sub Telnet1_HeRequestUsEnable(ByVal OptionNumber As IACOption)
Select Case OptionNumber
Case SGA
Telnet1.UsOptionState(SGA) = OptEnable
Case Else
Telnet1.UsOptionState(OptionNumber) = OptDisable
End Select
End Sub

Private Sub Telnet1_TermDataArive(ByVal DataLength As Long)
Dim InMessage As String
InMessage = Telnet1.GetTermData
VT1001.MessageIn InMessage
End Sub

Private Sub VT1001_MessageOut(OutMessage As String)
Telnet1.SendData OutMessage
End Sub
