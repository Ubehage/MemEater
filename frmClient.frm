VERSION 5.00
Begin VB.Form frmClient 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MEM_ITEM_SIZE As Long = SIZE_MEGA * 32
Private Const MEM_STEP As Long = 4096

Private Type Mem_Data
  Size As Long
  Data() As Byte
End Type

Dim MemItems As Long
Dim MemItem() As Mem_Data
Dim MemSize As Long

Dim MemStepIndex As Long

Dim TimerCounter As Long

Dim WithEvents ConsumeTimer As MemTimer
Attribute ConsumeTimer.VB_VarHelpID = -1
Dim WithEvents ClientTimer As MemTimer
Attribute ClientTimer.VB_VarHelpID = -1

Dim IsConsuming As Boolean
Dim hTimerState As Boolean

Friend Sub ConsumeMemory(BytesToConsume As Long)
  MemSize = BytesToConsume
  IsConsuming = True
  Set ConsumeTimer = New MemTimer
  ConsumeTimer.Interval = 1
  ConsumeTimer.Enabled = True
End Sub

Friend Sub ReleaseMemory()
  If MemItems = 0 Then Exit Sub
  Do Until MemItems = 0
    Erase MemItem(MemItems).Data
    MemItems = (MemItems - 1)
  Loop
  Erase MemItem
End Sub

Friend Sub SetClientTimer()
  KillClientTimer
  Set ClientTimer = New MemTimer
  ClientTimer.Interval = TIMER_INTERVAL_CLIENT
  ClientTimer.Enabled = True
End Sub

Friend Sub KillClientTimer()
  If Not ClientTimer Is Nothing Then
    ClientTimer.Enabled = False
    Set ClientTimer = Nothing
  End If
End Sub

Friend Sub ToggleHibernate(IsHibernating As Boolean)
  If IsConsuming Then
    If Not ConsumeTimer Is Nothing Then ConsumeTimer.Enabled = Not IsHibernating
  End If
  If Not ClientTimer Is Nothing Then
    If IsHibernating = True Then
      hTimerState = ClientTimer.Enabled
      If hTimerState = True Then ClientTimer.Enabled = False
    Else
      ClientTimer.Enabled = hTimerState
    End If
  End If
End Sub

Private Sub KillConsumeTimer()
  If Not ConsumeTimer Is Nothing Then
    ConsumeTimer.Enabled = False
    Set ConsumeTimer = Nothing
  End If
  IsConsuming = False
End Sub

Private Sub GoThroughMemory()
  Dim i As Long, j As Long
  If MemStepIndex = MEM_STEP Then
    MemStepIndex = 1
  Else
    MemStepIndex = (MemStepIndex + 1)
  End If
  For i = 1 To MemItems
    With MemItem(i)
      For j = MemStepIndex To .Size Step MEM_STEP
        .Data(j) = .Data(j) Xor &HFF
      Next
    End With
    If ExitNow Then Exit For
  Next
End Sub

Private Sub CheckAppMessages()
  Dim updMem As Boolean, c As Long, v As Long
  c = 0
  v = 0
  If ReadFromSharedMemory(False, True) = True Then
    With SharedMemory.Instances(SharedMemOffset)
      Select Case .ClientData.mData1
        Case MEMMSG_CONSUME
          If .ClientData.mData2 > 0 Then
            c = MEMMSG_CONSUME
            v = .ClientData.mData2
            .ClientData.mData1 = MEMMSG_SUCCESS
          Else
            .ClientData.mData1 = MEMMSG_ERROR
          End If
          .ClientData.mData2 = 0
          updMem = True
        Case MEMMSG_RELEASE
          If MemItems > 0 Then
            c = MEMMSG_RELEASE
            .ClientData.mData1 = MEMMSG_SUCCESS
          Else
            .ClientData.mData1 = MEMMSG_ERROR
          End If
          updMem = True
        Case MEMMSG_EXIT
          ExitNow = True
          .ClientData.mData1 = MEMMSG_SUCCESS
          'updMem = True
      End Select
    End With
    If updMem Then Call WriteToSharedMemory(False, True)
    Select Case c
      Case MEMMSG_CONSUME
        ConsumeMemory v
      Case MEMMSG_RELEASE
        ReleaseMemory
    End Select
  End If
End Sub

Private Sub CheckMainProcess()
  Call ReadFromSharedMemory(False, False, True, 0)
  If IsProcessAlive(SharedMemory.Instances(0).AppData.mData2) = False Then ExitNow = True
End Sub

Private Sub UnloadClient()
  KillConsumeTimer
  ReleaseMemory
  Call UnhookClient(Me.hWnd)
  UnloadAll
End Sub

Private Sub ClientTimer_Timer()
  ClientTimer.Enabled = False
  CheckMainProcess
  If ExitNow Then Unload Me Else CheckAppMessages
  If MemItems > 0 Then
    TimerCounter = (TimerCounter + 1)
    If TimerCounter >= TIMER_COUNT_MEMSTEP Then
      If IsConsuming = False Then GoThroughMemory
      TimerCounter = 0
    End If
  End If
  If Not ClientTimer Is Nothing Then ClientTimer.Enabled = True
End Sub

Private Sub ConsumeTimer_Timer()
  ConsumeTimer.Enabled = False
  If MemSize > 0 Then
    If (MemItems Mod 10) = 0 Then
      ReDim Preserve MemItem(1 To (MemItems + 10)) As Mem_Data
    End If
    MemItems = (MemItems + 1)
    With MemItem(MemItems)
      .Size = IIf(MemSize >= MEM_ITEM_SIZE, MEM_ITEM_SIZE, MemSize)
      MemSize = (MemSize - .Size)
      ReDim .Data(1 To .Size) As Byte
    End With
  End If
  If (MemSize = 0 Or ExitNow = True) Then
    KillConsumeTimer
    If ExitNow = True Then Unload Me
  Else
    ConsumeTimer.Enabled = True
  End If
End Sub

Private Sub Form_Load()
  If IsRunningInIDE = False Then Call HookClientWindowProc(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  KillClientTimer
  UnloadClient
End Sub
