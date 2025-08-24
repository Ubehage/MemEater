VERSION 5.00
Begin VB.UserControl FloodBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   5415
   Begin VB.Label lblRamText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   3390
      TabIndex        =   2
      Top             =   525
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lbl100 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1860
      TabIndex        =   1
      Top             =   285
      Width           =   360
   End
   Begin VB.Label lbl0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   540
      TabIndex        =   0
      Top             =   390
      Width           =   180
   End
End
Attribute VB_Name = "FloodBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FLOOD_HEIGHT As Long = 13

Dim m_Max As Currency
Dim m_Value As Currency
Dim m_AutoRefresh As Boolean
Dim m_ShowRamNumbers As Boolean

Dim BarRect As RECT
Dim FloodRect As RECT

Dim NeedFullRedraw As Boolean

Public Property Get Max() As Currency
  Max = m_Max
End Property
Public Property Get Value() As Currency
  Value = m_Value
End Property
Public Property Get AutoRefresh() As Boolean
  AutoRefresh = m_AutoRefresh
End Property
Public Property Get ShowRamNumbers() As Boolean
  ShowRamNumbers = m_ShowRamNumbers
End Property
Public Property Let Max(New_Max As Currency)
  m_Max = New_Max
  DoAutoRefresh
End Property
Public Property Let Value(New_Value As Currency)
  m_Value = New_Value
  DoAutoRefresh
End Property
Public Property Let AutoRefresh(New_AutoRefresh As Boolean)
  m_AutoRefresh = New_AutoRefresh
  DoAutoRefresh
End Property
Public Property Let ShowRamNumbers(New_ShowRamNumbers As Boolean)
  m_ShowRamNumbers = New_ShowRamNumbers
  lblRamText.Visible = m_ShowRamNumbers
  MoveObjects
  DoAutoRefresh
End Property

Public Sub Refresh(Optional FullRedraw As Boolean = False)
  If FullRedrawNeeded(FullRedraw) Then
    SetRects
    UserControl.Cls
  End If
  Redraw
  lblRamText.Visible = m_ShowRamNumbers
End Sub

Private Function FullRedrawNeeded(FullRedrawArg As Boolean) As Boolean
  If NeedFullRedraw Then
    FullRedrawNeeded = True
    NeedFullRedraw = False
  Else
    FullRedrawNeeded = FullRedrawArg
  End If
End Function

Private Sub DoAutoRefresh()
  If m_AutoRefresh = True Then Refresh
End Sub

Private Sub Redraw()
  If m_ShowRamNumbers Then
    UpdateRamLabel
    If NeedFullRedraw Then Refresh: Exit Sub
  End If
  DrawFlood
  DrawBar
End Sub

Private Sub DrawBar()
  With BarRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_OUTLINE_LIGHT, B
  End With
End Sub

Private Sub DrawFlood()
  Dim fWidth As Long, fLeft As Long
  fWidth = CalculateFloodWidth()
  With FloodRect
    If (.Left + fWidth) > .Right Then
      fWidth = (.Right - .Left)
    End If
    If fWidth >= 15 Then
      UserControl.Line (.Left, .Top)-((.Left + fWidth), .Bottom), GetFloodColor(), BF
    Else
      fWidth = 0
    End If
    fLeft = (.Left + fWidth)
    fWidth = (.Right - fLeft)
    If fWidth >= 15 Then UserControl.Line (fLeft, .Top)-(.Right, .Bottom), COLOR_CONTROLS, BF
  End With
End Sub

Private Function CalculateFloodWidth() As Long
  If m_Max <= 0 Then CalculateFloodWidth = 0: Exit Function
  With FloodRect
    CalculateFloodWidth = ((.Right - .Left) * (m_Value / m_Max))
  End With
End Function

Private Sub UpdateRamLabel()
  Dim w As Long
  With lblRamText
    w = .Width
    .Caption = GetByteSizeString(m_Value, False) & "/" & GetByteSizeString(m_Max, True)
    If .Width <> w Then NeedFullRedraw = True
    .Left = (UserControl.ScaleWidth - (.Width + Screen.TwipsPerPixelX))
    .ForeColor = GetFloodColor()
  End With
End Sub

Private Function GetFloodColor() As Long
  If m_Max <= 0 Then GetFloodColor = COLOR_GREEN: Exit Function
  Select Case (m_Value / m_Max)
    Case Is < 0.5
      GetFloodColor = COLOR_GREEN
    Case Is < 0.8
      GetFloodColor = COLOR_YELLOW
    Case Else
      GetFloodColor = COLOR_RED
  End Select
End Function

Private Sub SetRects()
  SetBarRect
  SetFloodRect
  BarRect.Bottom = (FloodRect.Bottom + Screen.TwipsPerPixelY)
  lbl100.Left = (BarRect.Right - lbl100.Width)
End Sub

Private Sub SetBarRect()
  With BarRect
    .Left = lbl0.Left
    .Top = ((lbl0.Top + lbl0.Height) + Screen.TwipsPerPixelY)
    If m_ShowRamNumbers Then
      .Right = (lblRamText.Left - (Screen.TwipsPerPixelX * 5))
    Else
      .Right = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
    End If
  End With
End Sub

Private Sub SetFloodRect()
  With FloodRect
    .Left = (BarRect.Left + Screen.TwipsPerPixelX)
    .Top = (BarRect.Top + Screen.TwipsPerPixelY)
    .Right = (BarRect.Right - Screen.TwipsPerPixelX)
    .Bottom = (.Top + (FLOOD_HEIGHT * Screen.TwipsPerPixelY))
  End With
End Sub

Private Sub MoveObjects()
  lbl0.Move 0, 0
  UpdateRamLabel
  lblRamText.Move (UserControl.ScaleWidth - (lblRamText.Width + Screen.TwipsPerPixelX)), (UserControl.ScaleHeight - lblRamText.Height)
  lbl100.Top = lbl0.Top
  SetRects
End Sub

Private Sub UserControl_InitProperties()
  m_Max = 100
  m_Value = 0
  m_ShowRamNumbers = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dim v As String
  On Error GoTo PropReadError
  v = ""
  v = PropBag.ReadProperty("Max")
  If v = "" Then v = "100"
  m_Max = CCur(v)
  v = ""
  v = PropBag.ReadProperty("Value")
  If v = "" Then v = "0"
  m_Value = CCur(v)
  v = ""
  v = PropBag.ReadProperty("ShowRamText")
  If v = "" Then v = "True"
  m_ShowRamNumbers = GetBooleanValueFromString(v)
  On Error GoTo 0
  Exit Sub
PropReadError:
  Resume Next
End Sub

Private Sub UserControl_Resize()
  MoveObjects
  UserControl.Height = (BarRect.Bottom + Screen.TwipsPerPixelY)
  Refresh True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error GoTo PropWriteError
  PropBag.WriteProperty "Max", CStr(m_Max)
  PropBag.WriteProperty "Value", CStr(m_Value)
  PropBag.WriteProperty "ShowRamText", GetStringFromBoolean(m_ShowRamNumbers)
  On Error GoTo 0
  Exit Sub
PropWriteError:
  Resume Next
End Sub
