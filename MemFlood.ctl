VERSION 5.00
Begin VB.UserControl MemFlood 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MemEater2.FloodBar Flood1 
      Height          =   450
      Left            =   300
      TabIndex        =   0
      Top             =   960
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   794
      Max             =   "100"
      Value           =   "0"
      ShowRamText     =   "True"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Physical RAM:"
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
      Height          =   225
      Left            =   315
      TabIndex        =   1
      Top             =   285
      Width           =   1365
   End
End
Attribute VB_Name = "MemFlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Title As String
Dim m_Value As Currency
Dim m_Max As Currency
Dim m_AutoRefresh As Boolean

Dim BorderRect As RECT

Dim NeedFullRedraw As Boolean

Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Get Value() As Currency
  Value = m_Value
End Property
Public Property Get Max() As Currency
  Max = m_Max
End Property
Public Property Get AutoRefresh() As Boolean
  AutoRefresh = m_AutoRefresh
End Property
Public Property Let Title(New_Title As String)
  m_Title = New_Title
  lblTitle.Caption = m_Title
End Property
Public Property Let Value(New_Value As Currency)
  m_Value = New_Value
  DoAutoRefresh
End Property
Public Property Let Max(New_Max As Currency)
  m_Max = New_Max
  DoAutoRefresh
End Property
Public Property Let AutoRefresh(New_AutoRefresh As Boolean)
  m_AutoRefresh = New_AutoRefresh
  DoAutoRefresh
End Property

Public Sub Refresh()
  If FullRedrawNeeded Then
    UserControl.Cls
    DrawBorder
  End If
  With Flood1
    .Value = m_Value
    .Max = m_Max
    .Refresh
  End With
End Sub

Private Function FullRedrawNeeded() As Boolean
  If NeedFullRedraw Then
    FullRedrawNeeded = True
    NeedFullRedraw = False
  End If
End Function

Private Sub DoAutoRefresh()
  If m_AutoRefresh Then Refresh
End Sub

Private Sub DrawBorder()
  With BorderRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_OUTLINE, B
  End With
End Sub

Private Sub SetBorderRect()
  With BorderRect
    .Left = 0
    .Top = 0
    .Right = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
  End With
End Sub

Private Sub SetProperties()
  lblTitle.Caption = m_Title
  With Flood1
    .Value = m_Value
    .Max = m_Max
  End With
End Sub

Private Sub MoveObjects()
  SetBorderRect
  lblTitle.Move (BorderRect.Left + (Screen.TwipsPerPixelX * 3)), (BorderRect.Top + (Screen.TwipsPerPixelY * 3))
  Flood1.Move lblTitle.Left, ((lblTitle.Top + lblTitle.Height) + (Screen.TwipsPerPixelY * 2)), (BorderRect.Right - (lblTitle.Left * 2))
  BorderRect.Bottom = ((Flood1.Top + Flood1.Height) + (lblTitle.Top - BorderRect.Top))
End Sub

Private Sub UserControl_InitProperties()
  m_Value = 0
  m_Max = 100
  m_Title = "Title"
  m_AutoRefresh = True
  SetProperties
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dim v As String
  On Error GoTo PropReadError
  v = ""
  v = PropBag.ReadProperty("Value")
  If v = "" Then v = "0"
  m_Value = CCur(v)
  v = ""
  v = PropBag.ReadProperty("Max")
  If v = "" Then v = "100"
  m_Max = CCur(v)
  v = ""
  v = PropBag.ReadProperty("Title")
  If v = "" Then v = "Title"
  m_Title = v
  v = ""
  v = PropBag.ReadProperty("AutoRefresh")
  If v = "" Then v = "True"
  m_AutoRefresh = GetBooleanValueFromString(v)
  On Error GoTo 0
  SetProperties
  Exit Sub
PropReadError:
  Resume Next
End Sub

Private Sub UserControl_Resize()
  MoveObjects
  UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + (BorderRect.Bottom + Screen.TwipsPerPixelY))
  NeedFullRedraw = True
  DoAutoRefresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error GoTo PropWriteError
  PropBag.WriteProperty "Value", CStr(m_Value)
  PropBag.WriteProperty "Max", CStr(m_Max)
  PropBag.WriteProperty "Title", CStr(m_Title)
  PropBag.WriteProperty "AutoRefresh", GetStringFromBoolean(m_AutoRefresh)
  On Error GoTo 0
  Exit Sub
PropWriteError:
  Resume Next
End Sub
