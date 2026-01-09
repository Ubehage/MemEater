VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MemEater2.SeperatorLine sepLineCmd 
      Height          =   30
      Left            =   375
      TabIndex        =   8
      Top             =   4860
      Width           =   8265
      _ExtentX        =   2990
      _ExtentY        =   212
   End
   Begin MemEater2.Button cmdCustomRelease 
      Height          =   420
      Left            =   6210
      TabIndex        =   7
      Top             =   4230
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   741
      Caption         =   "Release Multiple GB..."
      Enabled         =   "True"
   End
   Begin MemEater2.Button cmdCustomGB 
      Height          =   420
      Left            =   5790
      TabIndex        =   6
      Top             =   5340
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   741
      Caption         =   "Consume Multiple GB..."
      Enabled         =   "True"
   End
   Begin MemEater2.Button cmdReleaseGB 
      Height          =   420
      Left            =   3855
      TabIndex        =   5
      Top             =   4170
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   741
      Caption         =   "Release 1GB"
      Enabled         =   "True"
   End
   Begin MemEater2.Button cmdReleaseAll 
      Height          =   420
      Left            =   105
      TabIndex        =   4
      Top             =   4200
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   741
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
      Height          =   420
      Left            =   3600
      TabIndex        =   2
      Top             =   5280
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   741
      Caption         =   "Consume 1GB"
      Enabled         =   "True"
   End
   Begin MemEater2.Button cmdConsume 
      Height          =   420
      Left            =   180
      TabIndex        =   1
      Top             =   5190
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   741
      Caption         =   "Consume all RAM now"
      Enabled         =   "True"
   End
   Begin MemEater2.MemViewer MemViewer1 
      Height          =   2040
      Left            =   135
      TabIndex        =   0
      Top             =   1875
      Width           =   8805
      _ExtentX        =   15531
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
  ResizeButtons
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
  MoveButtons
End Sub

Private Sub MoveButtons()
  cmdConsume.Left = Status1.Left
  cmdConsume.Top = ((MemViewer1.Top + MemViewer1.Height) + (Screen.TwipsPerPixelY * 10))
  sepLineCmd.Width = (Me.ScaleWidth \ 2)
  sepLineCmd.Move (Me.ScaleWidth - sepLineCmd.Width) \ 2, ((cmdConsume.Top + cmdConsume.Height) + (BUTTON_SPACING / 2))
  cmdReleaseAll.Left = cmdConsume.Left
  cmdReleaseAll.Top = ((sepLineCmd.Top + sepLineCmd.Height) + (BUTTON_SPACING / 2))
  With MinimumFormSize
    If .y = 0 Then .y = ((cmdReleaseAll.Top + cmdReleaseAll.Height) + Status1.Top)
    If .x = 0 Then .x = GetMinimumButtonWidth
  End With
  cmdGB.Top = cmdConsume.Top
  cmdReleaseGB.Top = cmdReleaseAll.Top
  cmdCustomGB.Top = cmdGB.Top
  cmdCustomRelease.Top = cmdReleaseGB.Top
  cmdCustomGB.Left = (Me.ScaleWidth - (cmdCustomGB.Width + cmdReleaseAll.Left))
  cmdCustomRelease.Left = cmdCustomGB.Left
  Dim t As RECT
  With t
    .Left = (cmdReleaseAll.Left + cmdReleaseAll.Width)
    .Right = cmdCustomRelease.Left
    cmdReleaseGB.Left = (.Left + (((.Right - .Left) - cmdReleaseGB.Width) \ 2))
    cmdGB.Left = cmdReleaseGB.Left
  End With
End Sub

Private Sub ResizeButtons()
  Dim w As Long
  w = GetWidestButtonWidth()
  With cmdReleaseAll
    .Width = w
    .Height = BUTTON_HEIGHT
  End With
  With cmdReleaseGB
    .Width = w
    .Height = BUTTON_HEIGHT
  End With
  With cmdCustomRelease
    .Width = w
    .Height = BUTTON_HEIGHT
  End With
  With cmdConsume
    .Width = w
    .Height = BUTTON_HEIGHT
  End With
  With cmdGB
    .Width = w
    .Height = BUTTON_HEIGHT
  End With
  With cmdCustomGB
    .Width = w
    .Height = BUTTON_HEIGHT
  End With
End Sub

Private Function GetWidestButtonWidth() As Long
  Dim r As Long
  r = ReturnLargestValue(cmdReleaseAll.Width, r)
  r = ReturnLargestValue(cmdReleaseGB.Width, r)
  r = ReturnLargestValue(cmdCustomRelease.Width, r)
  r = ReturnLargestValue(cmdConsume.Width, r)
  r = ReturnLargestValue(cmdGB.Width, r)
  r = ReturnLargestValue(cmdCustomGB.Width, r)
  GetWidestButtonWidth = r
End Function

Private Function ReturnLargestValue(Value1 As Long, Value2 As Long) As Long
  ReturnLargestValue = IIf(Value1 > Value2, Value1, Value2)
End Function

Private Function GetMinimumButtonWidth() As Long
  Dim r As Long, s As Long
  s = BUTTON_SPACING
  GetMinimumButtonWidth = (cmdReleaseAll.Width + s) + (cmdReleaseGB.Width + s) + cmdCustomRelease.Width
  Exit Function
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
  cmdReleaseAll.Enabled = cmdReleaseGB.Enabled
  cmdCustomRelease.Enabled = cmdReleaseAll.Enabled
End Sub

Private Sub ConsumeNumberOfGB(NumGB As Long)
  If NumGB < 1 Then Exit Sub
  Dim i As Long, c As Long, l As Collection
  IsWorking = True
  
  'Disable all buttons. We don't want the user to mess things up.
  EnableButtons False
  
  Status1.Line3 = "Preparing child processes"
  Set l = New Collection
  For i = 1 To NumGB
    Status1.Line3 = Status1.Line3 & "."
    DoEvents
    
    'Check if user closed the window
    If ExitNow Then Exit For
    
    c = LaunchNewClient()
    If c = 0 Then Exit For
    l.Add c
  Next
  If Not ExitNow Then
    Do Until l.Count = 0
      ClientConsumeMemory l.Item(1), SIZE_GIGA
      l.Remove 1
    Loop
  End If
  Set l = Nothing
  Status1.Line3 = ""
  
  'enable buttons again
  If ExitNow = False Then EnableButtons True
  
  IsWorking = False
End Sub

Private Sub cmdConsume_Click()
  If ActiveClients >= MaxMemoryGB Then Exit Sub
  ConsumeNumberOfGB (MaxMemoryGB - ActiveClients)
End Sub

Private Sub cmdCustomGB_Click()
  Dim n As String, c As Long
  n = InputBox("How many GB do you wish to consume?", "Consume a number of GB...", CStr(MaxMemoryGB))
  If IsNumeric(n) Then
    c = CLng(n)
    If c <= 0 Then Exit Sub
    ConsumeNumberOfGB CLng(c)
  End If
End Sub

Private Sub cmdCustomRelease_Click()
  Dim n As String, c As Long
  n = InputBox("How many GB do you wish to release?" & vbCrLf & "Maximum: " & CStr(ActiveClients), "Release a number of GB...", CStr(ActiveClients))
  If IsNumeric(n) Then
    c = CLng(n)
    If c <= 0 Then Exit Sub
    If c > ActiveClients Then c = ActiveClients
    CloseNumberOfClients c, flFirst
  End If
End Sub

Private Sub cmdGB_Click()
  Dim c As Long
  c = LaunchNewClient()
  If c = 0 Then
    MsgBox "Could not complete the task!" & vbCrLf & "Unknown error", vbOKOnly Or vbCritical, APP_NAME
  Else
    ClientConsumeMemory c, SIZE_GIGA
  End If
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

Private Sub Form_Click()
  sepLineCmd.Refresh
End Sub

Private Sub Form_Resize()
  MoveObjects
  With MinimumFormSize
    If Me.ScaleWidth < .x Then Me.Width = ((Me.Width - Me.ScaleWidth) + .x): Exit Sub
    If Me.ScaleHeight <> .y Then Me.Height = ((Me.Height - Me.ScaleHeight) + .y): Exit Sub
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If IsWorking Then
    Cancel = 1
    ExitNow = True
  Else
    KillDisplayTimer
    'If ActiveClients > 0 Then ReleaseAllClients
    WindowOnTop Me.hWnd, False
    UnloadAll
  End If
End Sub
