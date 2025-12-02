VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   BackColor       =   &H002A2A2A&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Caption As String
Dim m_Enabled As Boolean

Dim m_ScreenRect As RECT
Dim m_IsCapturing As Boolean
Dim m_Hovering As Boolean
Dim m_IsPressed As Boolean
Dim m_MouseIsDown As Boolean
Dim m_KeyIsDown As Boolean

Dim ButtonRect As RECT
Dim TextPos As POINTAPI

Public Event Click()

Public Property Get Caption() As String
  Caption = m_Caption
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Let Caption(New_Caption As String)
  m_Caption = New_Caption
  SetTextPosition
  Refresh
End Property
Public Property Let Enabled(New_Enabled As Boolean)
  If m_Enabled = New_Enabled Then Exit Property
  m_Enabled = New_Enabled
  If Not m_Enabled Then
    EndHover
  Else
    Refresh
  End If
End Property

Public Sub Refresh()
  ClearBackground
  DrawTitle
  DrawBorder
End Sub

Private Sub ClearBackground()
  With ButtonRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), GetBackColor, BF
  End With
End Sub

Private Sub DrawTitle()
  UserControl.CurrentX = TextPos.x
  UserControl.CurrentY = TextPos.Y
  UserControl.ForeColor = GetTextColor
  UserControl.Print m_Caption
End Sub

Private Sub DrawBorder()
  Dim bColor As Long
  With ButtonRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_OUTLINE, B
    bColor = GetBackColor()
    UserControl.Line ((.Left + Screen.TwipsPerPixelX), (.Top + Screen.TwipsPerPixelY))-((.Right - Screen.TwipsPerPixelX), (.Bottom - Screen.TwipsPerPixelY)), bColor, B
    UserControl.Line ((.Left + (Screen.TwipsPerPixelX * 2)), (.Top + (Screen.TwipsPerPixelY * 2)))-((.Right - (Screen.TwipsPerPixelX * 2)), (.Bottom - (Screen.TwipsPerPixelY * 2))), bColor, B
  End With
End Sub

Private Sub SetTextPosition()
  TextPos.x = (UserControl.ScaleWidth - UserControl.TextWidth(m_Caption)) \ 2
  TextPos.Y = (UserControl.ScaleHeight - UserControl.TextHeight(m_Caption)) \ 2
End Sub

Private Sub SetScreenRect()
  Dim r As RECT, p As POINTAPI
  Call GetClientRect(UserControl.hWnd, r)
  p.x = r.Left
  p.Y = r.Top
  Call ClientToScreen(UserControl.hWnd, p)
  With m_ScreenRect
    .Left = p.x
    .Top = p.Y
    .Right = (.Left + r.Right)
    .Bottom = (.Top + r.Bottom)
  End With
End Sub

Private Sub SetButtonRect()
  With ButtonRect
    .Left = 0
    .Top = 0
    .Right = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
    .Bottom = (UserControl.ScaleHeight - Screen.TwipsPerPixelY)
  End With
End Sub

Private Function GetBackColor() As Long
  If m_Enabled = False Then
    GetBackColor = COLOR_BACKGROUND_DISABLED
  ElseIf m_IsPressed Then
    GetBackColor = COLOR_BUTTON_PRESSED
  ElseIf m_Hovering = True Then
    GetBackColor = COLOR_BUTTON_HOVER
  Else
    GetBackColor = COLOR_CONTROLS
  End If
End Function

Private Function GetTextColor() As Long
  If m_Enabled Then
    GetTextColor = COLOR_TEXT
  Else
    GetTextColor = COLOR_TEXT_DISABLED
  End If
End Function

Private Sub StartHover()
  If m_Hovering = False Then
    m_Hovering = True
    Refresh
    SetScreenRect
  End If
  If m_IsCapturing = True Then Exit Sub
  Call SetCapture(UserControl.hWnd)
  m_IsCapturing = True
End Sub

Private Sub EndHover()
  If m_Hovering = True Then
    m_Hovering = False
    Refresh
  End If
  If m_IsCapturing = False Or m_MouseIsDown = True Then Exit Sub
  EndCapture
End Sub

Private Sub EndCapture()
  If m_IsCapturing Then
    Call ReleaseCapture
    m_IsCapturing = False
  End If
End Sub

Private Function IsCursorOnButton() As Boolean
  Dim p As POINTAPI, hTop As Long
  Call GetCursorPos(p)
  If IsPointInRect(m_ScreenRect, p) = False Then If m_MouseIsDown = False Then Exit Function
  hTop = WindowFromPoint(p.x, p.Y)
  If hTop <> UserControl.hWnd Then If m_MouseIsDown = False Then Exit Function
  IsCursorOnButton = True
End Function

Private Sub DoClickEvent()
  EndCapture
  If m_Enabled Then RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
  m_Caption = "Button"
  m_Enabled = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If m_Enabled = False Then Exit Sub
  If KeyCode = vbKeySpace Then
    m_IsPressed = True
    m_KeyIsDown = True
    Refresh
  End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  If m_Enabled = False Then Exit Sub
  If KeyCode = vbKeySpace Then
    If m_IsPressed Then
      m_IsPressed = False
      Refresh
      If m_KeyIsDown Then
        m_KeyIsDown = False
        DoClickEvent
      End If
    ElseIf m_KeyIsDown Then
      m_KeyIsDown = False
    End If
  End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  If Button = vbLeftButton Then
    m_MouseIsDown = True
    m_IsPressed = True
    StartHover
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  StartHover
  If m_IsCapturing = False Then Exit Sub
  If IsCursorOnButton() = False Then EndHover
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  If Button = vbLeftButton Then
    m_MouseIsDown = False
    m_IsPressed = False
    Refresh
    If IsCursorOnButton Then
      DoClickEvent
    End If
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dim v As String
  On Error GoTo PropReadError
  v = ""
  v = PropBag.ReadProperty("Caption")
  If v = "" Then v = "Button"
  m_Caption = v
  v = ""
  v = PropBag.ReadProperty("Enabled")
  If v = "" Then v = "True"
  m_Enabled = GetBooleanValueFromString(v)
  On Error GoTo 0
  Exit Sub
PropReadError:
  Resume Next
End Sub

Private Sub UserControl_Resize()
  SetScreenRect
  SetButtonRect
  SetTextPosition
  Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error GoTo PropWriteError
  PropBag.WriteProperty "Caption", m_Caption
  PropBag.WriteProperty "Enabled", GetStringFromBoolean(m_Enabled)
  On Error GoTo 0
  Exit Sub
PropWriteError:
  Resume Next
End Sub
