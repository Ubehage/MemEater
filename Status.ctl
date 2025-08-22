VERSION 5.00
Begin VB.UserControl Status 
   AutoRedraw      =   -1  'True
   BackColor       =   &H002A2A2A&
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   ScaleHeight     =   3195
   ScaleWidth      =   6720
   Begin VB.Label lStatus1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   420
      TabIndex        =   3
      Top             =   915
      Width           =   630
   End
   Begin VB.Label lStatus2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   390
      TabIndex        =   2
      Top             =   1185
      Width           =   630
   End
   Begin VB.Label lStatus3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   480
      TabIndex        =   1
      Top             =   1515
      Width           =   630
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   300
      Width           =   525
   End
End
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Line1 As String
Dim m_Line2 As String
Dim m_Line3 As String

Dim BorderRect As RECT
Dim StatusRect As RECT

Public Property Get Line1() As String
  Line1 = m_Line1
End Property
Public Property Get Line2() As String
  Line2 = m_Line2
End Property
Public Property Get Line3() As String
  Line3 = m_Line3
End Property
Public Property Let Line1(New_Line1 As String)
  If m_Line1 = New_Line1 Then Exit Property
  m_Line1 = New_Line1
  lStatus1.Caption = m_Line1
End Property
Public Property Let Line2(New_Line2 As String)
  If m_Line2 = New_Line2 Then Exit Property
  m_Line2 = New_Line2
  lStatus2.Caption = m_Line2
End Property
Public Property Let Line3(New_Line3 As String)
  If m_Line3 = New_Line3 Then Exit Property
  m_Line3 = New_Line3
  lStatus3.Caption = m_Line3
End Property

Public Sub Refresh(Optional FullRefresh As Boolean = False)
  If FullRefresh Then ClearBackground Else ClearTextRect
  DrawText
  DrawBorder
End Sub

Private Sub ClearBackground()
  With BorderRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_BACKGROUND, BF
  End With
End Sub

Private Sub ClearTextRect()
  With StatusRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_CONTROLS, BF
  End With
End Sub

Private Sub DrawText()
  lStatus1.Caption = m_Line1
  lStatus1.Refresh
  lStatus2.Caption = m_Line2
  lStatus2.Refresh
  lStatus3.Caption = m_Line3
  lStatus3.Refresh
End Sub

Private Sub DrawBorder()
  With BorderRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_OUTLINE, B
    UserControl.Line ((.Left + Screen.TwipsPerPixelX), (.Top + Screen.TwipsPerPixelY))-((.Right - Screen.TwipsPerPixelX), (.Bottom - Screen.TwipsPerPixelY)), COLOR_BACKGROUND, B
    UserControl.Line ((.Left + (Screen.TwipsPerPixelX * 2)), (.Top + (Screen.TwipsPerPixelY * 2)))-((.Right - (Screen.TwipsPerPixelX * 2)), (.Bottom - (Screen.TwipsPerPixelY * 2))), COLOR_BACKGROUND, B
  End With
End Sub

Private Sub SetRects()
  SetBorderRect
  SetStatusRect
  BorderRect.Bottom = (StatusRect.Bottom + (Screen.TwipsPerPixelY * 3))
End Sub

Private Sub SetBorderRect()
  With BorderRect
    .Left = 0
    .Top = 0
    .Right = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
  End With
End Sub

Private Sub SetStatusRect()
  With StatusRect
    .Left = (BorderRect.Left + (Screen.TwipsPerPixelX * 10))
    .Top = ((lblTitle.Top + lblTitle.Height) + (Screen.TwipsPerPixelY * 5))
    .Right = (BorderRect.Right - (.Left - BorderRect.Left))
    lStatus1.Move .Left, .Top
    lStatus2.Move lStatus1.Left, ((lStatus1.Top + lStatus1.Height) + (Screen.TwipsPerPixelY * 3))
    lStatus3.Move lStatus2.Left, ((lStatus2.Top + lStatus2.Height) + (Screen.TwipsPerPixelY * 3))
    .Bottom = ((lStatus3.Top + lStatus3.Height) + (Screen.TwipsPerPixelY * 1))
  End With
End Sub

Private Function GetSampleText() As String
  GetSampleText = "|ABCDEFG" & vbCrLf & "|gbuydb"
End Function

Private Sub MoveObjects()
  SetBorderRect
  With BorderRect
    lblTitle.Move (.Left + (Screen.TwipsPerPixelX * 4)), (.Top + (Screen.TwipsPerPixelY * 4))
  End With
  SetStatusRect
  BorderRect.Bottom = (StatusRect.Bottom + (Screen.TwipsPerPixelY * 3))
End Sub

Private Sub UserControl_Initialize()
  'p.BackColor = COLOR_BACKGROUND
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dim v As String
  On Error GoTo PropReadError
  v = ""
  v = PropBag.ReadProperty("Line1")
  m_Line1 = v
  v = ""
  v = PropBag.ReadProperty("Line2")
  m_Line2 = v
  v = ""
  v = PropBag.ReadProperty("Line3")
  m_Line3 = v
  On Error GoTo 0
  Exit Sub
PropReadError:
  Resume Next
End Sub

Private Sub UserControl_Resize()
  MoveObjects
  UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + (BorderRect.Bottom + Screen.TwipsPerPixelY))
  Refresh True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error GoTo PropWriteError
  PropBag.WriteProperty "Line1", m_Line1
  PropBag.WriteProperty "Line2", m_Line2
  PropBag.WriteProperty "Line3", m_Line3
  On Error GoTo 0
  Exit Sub
PropWriteError:
  Resume Next
End Sub
