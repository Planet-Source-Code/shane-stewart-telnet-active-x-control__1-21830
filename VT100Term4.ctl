VERSION 5.00
Begin VB.UserControl VT100 
   BackColor       =   &H8000000A&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox Term 
      ForeColor       =   &H00000040&
      Height          =   495
      HideSelection   =   0   'False
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox Term2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   1455
      Left            =   1440
      ScaleHeight     =   1455
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   120
      Min             =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3360
      Value           =   1
      Width           =   1215
   End
   Begin VB.Frame Corner 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   3360
      Width           =   255
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1215
      Left            =   4560
      Min             =   1
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Value           =   1
      Width           =   255
   End
End
Attribute VB_Name = "VT100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_ScrollBuffer = 25
Const m_def_CharPerLine = 80
'Property Variables:
Dim m_ScrollBuffer As Long
Dim m_CharPerLine As Long

Dim Processing As Boolean
Dim StringIn As New Collection
Dim fTextHeight As Integer
Dim fTextWidth As Integer
Dim ScreenImage As New Collection
Dim CurX As Integer
Dim CurY As Integer
Dim tempstring As String

Public Event MessageOut(OutMessage As String)

Public Sub MessageIn(InMessage As String)

StringIn.Add InMessage
If Not Processing Then
Processing = True

Process_Text
End If
End Sub

Private Sub TxtTerm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then
RaiseEvent MessageOut(Chr$(KeyCode))
KeyCode = 0
End If

End Sub

Private Sub TxtTerm_KeyPress(keyascii As Integer)
RaiseEvent MessageOut(Chr$(keyascii))
keyascii = 0
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,25
Public Property Get ScrollBuffer() As Integer
Attribute ScrollBuffer.VB_Description = "Number of lines to buffer in memory"
    ScrollBuffer = m_ScrollBuffer
End Property

Public Property Let ScrollBuffer(ByVal New_ScrollBuffer As Integer)
    m_ScrollBuffer = New_ScrollBuffer
    PropertyChanged "ScrollBuffer"
    Init_Controls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,80
Public Property Get CharPerLine() As Integer
Attribute CharPerLine.VB_Description = "The number of characters on a single line"
    CharPerLine = m_CharPerLine
End Property

Public Property Let CharPerLine(ByVal New_CharPerLine As Integer)
    m_CharPerLine = New_CharPerLine
    PropertyChanged "CharPerLine"
    Init_Controls
End Property

Private Sub Term_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case 37

Case 38
'RaiseEvent MessageOut(Chr$(27) & Chr$(91) & Chr$(65))

Case 39

Case 40

Case Else

End Select
KeyCode = 0
End Sub

Private Sub Term_KeyPress(keyascii As Integer)

Select Case keyascii
Case 32 To 125
RaiseEvent MessageOut(Chr$(keyascii))

Case 8
RaiseEvent MessageOut(Chr$(keyascii))

Case 13
RaiseEvent MessageOut(Chr$(keyascii))

End Select
keyascii = 0
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_ForeColor = m_def_ForeColor

    m_ScrollBuffer = m_def_ScrollBuffer
    m_CharPerLine = m_def_CharPerLine
    CurX = 1
    CurY = 1
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)


    m_ScrollBuffer = PropBag.ReadProperty("ScrollBuffer", m_def_ScrollBuffer)
    m_CharPerLine = PropBag.ReadProperty("CharPerLine", m_def_CharPerLine)
End Sub

Private Sub UserControl_Resize()
Init_Controls
End Sub

Private Sub UserControl_Show()
Init_Controls
CurX = 1
CurY = 1
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)


    Call PropBag.WriteProperty("ScrollBuffer", m_ScrollBuffer, m_def_ScrollBuffer)
    Call PropBag.WriteProperty("CharPerLine", m_CharPerLine, m_def_CharPerLine)
End Sub

Public Function ClearScreen()
Dim x As Integer
Term.Text = ""
For x = 1 To ScreenImage.Count
ScreenImage.Remove (1)
Next x
CurY = 1
VScroll.Value = 1
HScroll.Value = 1
End Function

Private Sub Init_Controls()

fTextHeight = Term2.TextHeight("M")
fTextWidth = Term2.TextWidth("M")

VScroll.Top = 0
VScroll.Left = Width - VScroll.Width
HScroll.Left = 0
HScroll.Top = Height - HScroll.Height
Corner.Top = Height - Corner.Height
Corner.Left = Width - Corner.Width
VScroll.Height = Height - Corner.Height
HScroll.Width = Width - Corner.Width

Term.Top = 0
Term.Left = 0
Term.Height = Height - HScroll.Height
Term.Width = Width - VScroll.Width

Term2.Top = 0
Term2.Left = 0
Term2.Height = Height - HScroll.Height
Term2.Width = Width - VScroll.Width

If MaxY < ScrollBuffer Then
VScroll.Enabled = True
Else: VScroll.Enabled = False
End If

If MaxX < CharPerLine Then
HScroll.Enabled = True
Else: HScroll.Enabled = False
End If

VScroll.Max = ScrollBuffer - MaxY
HScroll.Max = CharPerLine - MaxX

VScroll.LargeChange = ScrollBuffer / 20

End Sub

Private Sub Process_Text()

Dim Temp As String
Dim x As Long
Dim displaychar As String
Dim NewPos As Integer

Do While StringIn.Count > 0
Temp = StringIn(1)
StringIn.Remove (1)

For x = 1 To Len(Temp)
displaychar = Mid$(Temp, x, 1)

Select Case displaychar
Case Chr(7)
Beep
Case Chr(8)
CurX = CurX - 1

Case vbLf

Case vbCr
Add_Character vbCr
Add_Character vbLf
Do Until ScreenImage.Count < ScrollBuffer
ScreenImage.Remove (1)
CurY = ScreenImage.Count
Loop
CurY = CurY + 1
CurX = 1

Case Chr(20) To Chr(126)
Add_Character displaychar

Case Chr(0)


Case Else


End Select
Next x

If CurY > MaxY Then NewPos = (CurY - MaxY) Else NewPos = 1
If VScroll.Value = NewPos Then
Paint_Screen
Else: VScroll.Value = NewPos
End If

DoEvents
Loop
Processing = False

End Sub

Private Sub Paint_Screen()
Dim x As Integer
Term.Text = ""

For x = (VScroll.Value) To (VScroll.Value + MaxY)
If x > ScreenImage.Count Then
Else
Term.Text = Term.Text & Mid$(ScreenImage(x), HScroll.Value, MaxX)
End If
Next x
If VScroll.Value + MaxY >= ScreenImage.Count Then
Term.SelStart = Len(Term.Text)
Else

End If

End Sub

Private Function MaxX() As Integer
MaxX = (Term.Width \ fTextWidth) - 1
End Function

Private Function MaxY() As Integer
MaxY = (Term.Height \ fTextHeight) - 1
End Function

Private Sub VScroll_Change()
Paint_Screen
End Sub

Private Sub Add_Character(displaychar As String)
Do Until ScreenImage.Count >= CurY
ScreenImage.Add ""
Loop
tempstring = ScreenImage(CurY)
tempstring = Left$(tempstring, CurX - 1) & displaychar & Mid$(tempstring, CurX + 1)
ScreenImage.Add tempstring, , CurY
ScreenImage.Remove (CurY + 1)
CurX = CurX + 1
End Sub



