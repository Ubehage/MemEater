VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14310
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MemEater2.Button cmdCustomGB 
      Height          =   495
      Left            =   7785
      TabIndex        =   6
      Top             =   4185
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   873
      Caption         =   "Consume Multiple GB..."
      Enabled         =   "True"
   End
   Begin MemEater2.Button cmdReleaseGB 
      Height          =   495
      Left            =   3855
      TabIndex        =   5
      Top             =   4170
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   873
      Caption         =   "Release 1GB"
      Enabled         =   "True"
   End
   Begin MemEater2.Button cmdReleaseAll 
      Height          =   525
      Left            =   105
      TabIndex        =   4
      Top             =   4200
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   926
      Caption         =   "Release all consumed memory"
      Enabled         =   "True"
   End
   Begin MemEater2.Status Status1 
      Height          =   870
      Left            =   585
      TabIndex        =   3
      Top             =   240
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   2117
      Line1           =   ""
      Line2           =   ""
      Line3           =   ""
   End
   Begin MemEater2.Button cmdGB 
      Height          =   450
      Left            =   5820
      TabIndex        =   2
      Top             =   4200
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   794
      Caption         =   "Consume 1GB"
      Enabled         =   "True"
   End
   Begin MemEater2.Button cmdConsume 
      Height          =   555
      Left            =   11160
      TabIndex        =   1
      Top             =   4095
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   979
      Caption         =   "Consume all RAM now"
      Enabled         =   "True"
   End
   Begin MemEater2.MemViewer MemViewer1 
      Height          =   2040
      Left            =   135
      TabIndex        =   0
      Top             =   1875
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   3598
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BUTTON_HEIGHT As Integer = (15 * 28) '28 pixels
Private Const BUTTON_SPACING As Integer = (15 * 10) '10 pixels

Dim MinimumFormSize As POINTAPI

Dim WithEvents DisplayTimer As MemTimer
Attribute DisplayTimer.VB_VarHelpID = -1

Dim TickCounter As Integer

Dim IsWorking As Boolean

Friend Sub SetForm()
  WindowOnTop Me.hWnd, True
  Me.BackColor = COLOR_BACKGROUND
  Me.Caption = APP_NAME
  Me.Show
  MoveObjects
  CheckClientButtons
End Sub

Friend Sub SetDisplayTimer()
  Set DisplayTimer = New MemTimer
  DisplayTimer.Interval = TIMER_INTERVAL_DISPLAY
  DisplayTimer.Enabled = True
End Sub

Private Sub KillDisplayTimer()
  If Not DisplayTimer Is Nothing Then
    DisplayTimer.Enabled = False
    Set DisplayTimer = Nothing
  End If
End Sub

Private Sub MoveObjects()
  Status1.Move (Screen.TwipsPerPixelX * 3), (Screen.TwipsPerPixelY * 3)
  Status1.Width = (Me.ScaleWidth - (Status1.Left * 2))
  MemViewer1.Move Status1.Left, ((Status1.Top + Status1.Height) + (Screen.TwipsPerPixelY * 3)), Status1.Width
  cmdConsume.Height = BUTTON_HEIGHT
  cmdGB.Height = cmdConsume.Height
  cmdReleaseGB.Height = cmdGB.Height
  cmdCustomGB.Height = cmdReleaseGB.Height
  cmdReleaseAll.Height = cmdReleaseGB.Height
  cmdReleaseAll.Top = ((MemViewer1.Top + MemViewer1.Height) + (Screen.TwipsPerPixelY * 10))
  cmdReleaseAll.Left = Status1.Left
  With MinimumFormSize
    If .Y = 0 Then .Y = ((cmdReleaseAll.Top + cmdReleaseAll.Height) + Status1.Top)
    If .x = 0 Then .x = GetMinimumButtonWidth()
  End With
  cmdConsume.Top = cmdReleaseAll.Top
  cmdConsume.Left = (Me.ScaleWidth - (cmdConsume.Width + cmdReleaseAll.Left))
  cmdReleaseGB.Top = cmdConsume.Top
  cmdGB.Top = cmdReleaseGB.Top
  cmdCustomGB.Top = cmdGB.Top
  Dim t As RECT
  With t
    .Left = (cmdReleaseAll.Left + cmdReleaseAll.Width)
    .Right = cmdConsume.Left
    
    Dim spacing As Long
    spacing = (((.Right - .Left) - (cmdGB.Width + cmdReleaseGB.Width + cmdCustomGB.Width)) / 4)
    If spacing < BUTTON_SPACING Then spacing = BUTTON_SPACING
    cmdGB.Left = ((cmdReleaseGB.Left + cmdReleaseGB.Width) + spacing)
    'cmdGB.Left = (cmdConsume.Left - (cmdGB.Width + spacing))
    cmdCustomGB.Left = ((cmdGB.Left + cmdGB.Width) + spacing)
    cmdReleaseGB.Left = (.Left + spacing)
  End With
End Sub

Private Function GetMinimumButtonWidth() As Long
  Dim r As Long, s As Long
  s = BUTTON_SPACING
  r = (cmdConsume.Width + s)
  r = (r + (cmdGB.Width + s))
  r = (r + (cmdReleaseGB.Width + s))
  r = (r + (cmdCustomGB.Width + s))
  GetMinimumButtonWidth = (r + cmdReleaseAll.Width)
End Function

Private Sub EnableButtons(DoEnable As Boolean)
  cmdConsume.Enabled = DoEnable
  cmdCustomGB.Enabled = DoEnable
  cmdGB.Enabled = DoEnable
  cmdReleaseAll.Enabled = DoEnable
  cmdReleaseGB.Enabled = DoEnable
End Sub

Private Sub CheckClientButtons()
  cmdConsume.Enabled = IIf(ActiveClients >= MaxMemoryGB, False, True)
  cmdReleaseGB.Enabled = IIf(ActiveClients > 0, True, False)
  cmdReleaseAll.Enabled = IIf(ActiveClients > 0, True, False)
End Sub

Private Sub ConsumeNumberOfGB(NumGB As Long)
  If NumGB < 1 Then Exit Sub
  Dim i As Long, c As Long, l As Collection
  IsWorking = True
  
  'Disable all buttons. We don't want the user to mess things up.
  EnableButtons False
  
  Set l = New Collection
  Status1.Line3 = "Preparing child processes"
  For i = 1 To NumGB
    Status1.Line3 = Status1.Line3 & "."
    DoEvents
    
    'Check if user closed the window
    If ExitNow Then GoTo ExitConsume
    
    c = LaunchNewClient()
    If c = 0 Then Exit For
    l.Add c
  Next
  Do Until l.Count = 0
    ClientConsumeMemory l.Item(1), SIZE_GIGA
    l.Remove 1
  Loop
ExitConsume:
  Set l = Nothing
  Status1.Line3 = ""
  
  'enable buttons again
  EnableButtons True
  
  IsWorking = False
End Sub

Private Sub cmdConsume_Click()
  If ActiveClients >= MaxMemoryGB Then Exit Sub
  ConsumeNumberOfGB (MaxMemoryGB - ActiveClients)
End Sub

Private Sub cmdCustomGB_Click()
  Dim n As String
  n = InputBox("How many GB do you wish to consume?", "Consume a number of GB...", CStr(MaxMemoryGB))
  If IsNumeric(n) Then ConsumeNumberOfGB CLng(n)
End Sub

Private Sub cmdGB_Click()
  Dim nIndex As Long
  nIndex = LaunchNewClient()
  If nIndex = 0 Then GoTo GBError
  ClientConsumeMemory nIndex, SIZE_GIGA
  Exit Sub
GBError:
  MsgBox "Could not complete the task!" & vbCrLf & "Unknown error", vbOKOnly Or vbCritical, APP_NAME
End Sub

Private Sub cmdReleaseAll_Click()
  If Not DisplayTimer Is Nothing Then DisplayTimer.Enabled = False
  ReleaseAllClients
  If Not DisplayTimer Is Nothing Then DisplayTimer.Enabled = True
End Sub

Private Sub cmdReleaseGB_Click()
  CloseOneClient flFirst
End Sub

Private Sub DisplayTimer_Timer()
  DisplayTimer.Enabled = False
  If ExitNow Then Unload Me: Exit Sub
  MemViewer1.Refresh
  If TickCounter = 1 Then
    CheckActiveProcesses
    If Not IsWorking Then
      CheckClientButtons
    End If
    TickCounter = 0
  Else
    TickCounter = 1
  End If
  Status1.Line1 = "Active child processes: " & CStr(ActiveClients)
  DisplayTimer.Enabled = True
End Sub

Private Sub Form_Resize()
  MoveObjects
  With MinimumFormSize
    If Me.ScaleWidth < .x Then Me.Width = ((Me.Width - Me.ScaleWidth) + .x): Exit Sub
    If Me.ScaleHeight <> .Y Then Me.Height = ((Me.Height - Me.ScaleHeight) + .Y): Exit Sub
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If IsWorking Then
    Cancel = 1
    ExitNow = True
  Else
    KillDisplayTimer
    If ActiveClients > 0 Then ReleaseAllClients
    WindowOnTop Me.hWnd, False
    UnloadAll
  End If
End Sub
