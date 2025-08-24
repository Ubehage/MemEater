VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MemEater2.Button cmdRelease 
      Height          =   525
      Left            =   6450
      TabIndex        =   4
      Top             =   4275
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
      Height          =   435
      Left            =   3615
      TabIndex        =   2
      Top             =   4155
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   767
      Caption         =   "Consume 1GB"
      Enabled         =   "True"
   End
   Begin MemEater2.Button cmdConsume 
      Height          =   555
      Left            =   315
      TabIndex        =   1
      Top             =   4170
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

Private Const BUTTON_HEIGHT As Integer = 28
Private Const BUTTON_SPACING As Integer = 10

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
  cmdConsume.Height = (Screen.TwipsPerPixelY * BUTTON_HEIGHT)
  cmdGB.Height = cmdConsume.Height
  cmdRelease.Height = cmdGB.Height
  cmdRelease.Top = ((MemViewer1.Top + MemViewer1.Height) + (Screen.TwipsPerPixelY * 10))
  cmdRelease.Left = Status1.Left
  With MinimumFormSize
    If .Y = 0 Then .Y = ((cmdRelease.Top + cmdRelease.Height) + Status1.Top)
    If .x = 0 Then .x = GetMinimumButtonWidth()
  End With
  cmdConsume.Top = cmdRelease.Top
  cmdConsume.Left = (Me.ScaleWidth - (cmdConsume.Width + cmdRelease.Left))
  cmdGB.Top = cmdConsume.Top
  Dim t As RECT
  With t
    .Left = ((cmdRelease.Left + cmdRelease.Width) + (Screen.TwipsPerPixelX * BUTTON_SPACING))
    .Right = (cmdConsume.Left - (Screen.TwipsPerPixelX * BUTTON_SPACING))
    cmdGB.Left = (.Left + ((.Right - .Left) - cmdGB.Width) \ 2)
  End With
End Sub

Private Function GetMinimumButtonWidth() As Long
  Dim r As Long, s As Long
  s = (Screen.TwipsPerPixelX * BUTTON_SPACING)
  r = (cmdConsume.Width + s)
  r = (r + (cmdGB.Width + s))
  GetMinimumButtonWidth = (r + cmdRelease.Width)
End Function

Private Sub EnableButtons(DoEnable As Boolean)
  cmdConsume.Enabled = DoEnable
  cmdGB.Enabled = DoEnable
  cmdRelease.Enabled = DoEnable
End Sub

Private Sub cmdConsume_Click()
  Dim i As Long, c As Long, l As Collection
  If ActiveClients >= MaxMemoryGB Then Exit Sub
  IsWorking = True
  'Disable all buttons. We don't want the user to mess things up.
  EnableButtons False
  Set l = New Collection
  Status1.Line3 = "Preparing child processes"
  i = ActiveClients
  Do Until i = MaxMemoryGB
    Status1.Line3 = Status1.Line3 & "."
    DoEvents
    
    'Check if user wanted to close the window
    If ExitNow Then GoTo ExitConsume
    c = LaunchNewClient
    If c = 0 Then Exit Do
    i = (i + 1)
    l.Add c
  Loop
  Do Until l.Count = 0
    ClientConsumeMemory l.Item(1), SIZE_GIGA
    l.Remove 1
  Loop
ExitConsume:
  Set l = Nothing
  Status1.Line3 = ""
  
  'Enable buttons again.
  EnableButtons True
  IsWorking = False
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

Private Sub cmdRelease_Click()
  If Not DisplayTimer Is Nothing Then DisplayTimer.Enabled = False
  ReleaseAllClients
  If Not DisplayTimer Is Nothing Then DisplayTimer.Enabled = True
End Sub

Private Sub DisplayTimer_Timer()
  DisplayTimer.Enabled = False
  If ExitNow Then Unload Me: Exit Sub
  MemViewer1.Refresh
  If TickCounter = 1 Then
    CheckActiveProcesses
    If Not IsWorking Then cmdConsume.Enabled = IIf(ActiveClients >= MaxMemoryGB, False, True)
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
    WindowOnTop Me.hWnd, False
    If ActiveClients > 0 Then
      Select Case MsgBox("There are active memory processes. Close them and free memory?", vbYesNoCancel Or vbQuestion Or vbMsgBoxSetForeground Or vbDefaultButton2, APP_NAME)
        Case vbYes
          ReleaseAllClients
        Case vbCancel
          Cancel = 1
          ExitNow = False
          SetDisplayTimer
          WindowOnTop Me.hWnd, True
          Exit Sub
      End Select
    End If
    UnloadAll
  End If
End Sub
