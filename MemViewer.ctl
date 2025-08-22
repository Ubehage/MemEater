VERSION 5.00
Begin VB.UserControl MemViewer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   ScaleHeight     =   4875
   ScaleWidth      =   8985
   Begin MemEater2.MemFlood flVirtual 
      Height          =   810
      Left            =   555
      TabIndex        =   2
      Top             =   2400
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   1429
      Value           =   "0"
      Max             =   "100"
      Title           =   "Committed (RAM and Pagefile)"
      AutoRefresh     =   "True"
   End
   Begin MemEater2.MemFlood flRam 
      Height          =   810
      Left            =   480
      TabIndex        =   1
      Top             =   1125
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   1429
      Value           =   "0"
      Max             =   "100"
      Title           =   "Physical RAM"
      AutoRefresh     =   "True"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Memory Load"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   435
      TabIndex        =   0
      Top             =   465
      Width           =   1230
   End
End
Attribute VB_Name = "MemViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Title As String

Dim BorderRect As RECT

Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Let Title(New_Title As String)
  m_Title = New_Title
  lblTitle.Caption = m_Title
End Property

Public Sub Refresh()
  Redraw
  UpdateMemoryBars
End Sub

Public Sub UpdateMemoryBars()
  Dim mInfo As INTERNAL_MEM_STATUS
  Call GetMemoryInfo(mInfo)
  With mInfo
    With .PhysicalMemory
      flRam.Max = .Max
      flRam.Value = .Used
    End With
    With .VirtualMemory
      flVirtual.Max = .Max
      flVirtual.Value = .Used
    End With
  End With
  flRam.Refresh
  flVirtual.Refresh
End Sub

Private Sub Redraw()
  UserControl.Cls
  DrawBorder
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
    .Right = (UserControl.ScaleWidth - (Screen.TwipsPerPixelX + (.Left * 2)))
  End With
End Sub

Private Sub MoveObjects()
  SetBorderRect
  lblTitle.Move (BorderRect.Left + (Screen.TwipsPerPixelX * 3)), (BorderRect.Top + (Screen.TwipsPerPixelY * 3))
  flRam.Move lblTitle.Left, ((lblTitle.Top + lblTitle.Height) + (Screen.TwipsPerPixelX * 2)), (BorderRect.Right - (lblTitle.Left * 2))
  flVirtual.Move flRam.Left, ((flRam.Top + flRam.Height) + (Screen.TwipsPerPixelY * 2)), flRam.Width
  BorderRect.Bottom = ((flVirtual.Top + flVirtual.Height) + (Screen.TwipsPerPixelY * 3))
End Sub

Private Sub UserControl_Resize()
  MoveObjects
  UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + (BorderRect.Bottom + Screen.TwipsPerPixelY))
  Refresh
End Sub
