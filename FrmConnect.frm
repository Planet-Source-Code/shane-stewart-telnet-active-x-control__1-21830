VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   ControlBox      =   0   'False
   Icon            =   "FrmConnect.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox TxtHost 
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      Format          =   "0"
      Mask            =   "99999"
      PromptChar      =   " "
   End
   Begin VB.ComboBox CmbTerm 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "vt100"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox CmbHost 
      Height          =   1740
      ItemData        =   "FrmConnect.frx":0ECA
      Left            =   1440
      List            =   "FrmConnect.frx":0ECC
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Term Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Host Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "FrmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmbHost_Click()
TxtHost.Text = CmbHost.ItemData(CmbHost.ListIndex)

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub CmdConnect_Click()
If CmbHost.Text = "" Or TxtHost.Text = "" Then Exit Sub
If Val(TxtHost.Text) < 1 Or Val(TxtHost.Text) > 65535 Then Exit Sub

FrmMain.VT1001.ClearScreen
FrmMain.Telnet1.Connect CmbHost, TxtHost
FrmMain.MnuDisconnect.Enabled = True
FrmMain.MnuRemote.Enabled = False

If CmbHost.ListIndex = -1 Then
If Hosts.Count >= 15 Then
Hosts.Remove (1)
Port.Remove (1)
End If
Hosts.Add CmbHost.Text
Port.Add TxtHost.Text
End If

FrmMain.load_hostmnu
Unload Me
End Sub

Private Sub Form_Deactivate()
FrmConnect.SetFocus
End Sub

Private Sub Form_Load()
Dim x As Integer
Dim y As Integer

y = 0
CmbHost.Clear

Me.Top = ((Screen.Height - Me.Height) / 2) - 250
Me.Left = (Screen.Width - Me.Width) / 2

For x = 15 To 1 Step -1
If Hosts(x) <> "" Then
CmbHost.AddItem Hosts(x), y
CmbHost.ItemData(y) = Port(x)
y = y + 1
End If
Next x

TxtHost.Text = 23

End Sub

Private Sub Form_LostFocus()
'FrmConnect.SetFocus
End Sub

Private Sub TxtHost_GotFocus()
TxtHost.SelStart = 0
TxtHost.SelLength = Len(TxtHost.Text)
End Sub
