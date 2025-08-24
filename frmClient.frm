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

Dim MemStepIndex As Long

Dim WithEvents ClientTimer As MemTimer
Attribute ClientTimer.VB_VarHelpID = -1

Friend Sub ConsumeMemory(BytesToConsume As Long)
  Dim cB As Long
  cB = BytesToConsume
  Do While cB > 0
    If (MemItems Mod 10) = 0 Then
      ReDim Preserve MemItem(1 To (MemItems + 10)) As Mem_Data
    End If
    MemItems = (MemItems + 1)
    With MemItem(MemItems)
      .Size = IIf(cB >= MEM_ITEM_SIZE, MEM_ITEM_SIZE, cB)
      cB = (cB - .Size)
      ReDim .Data(1 To .Size) As Byte
    End With
    If ExitNow Then Exit Do
  Loop
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
      Select Case .ClienData.mData1
        Case MEMMSG_CONSUME
          If .ClienData.mData2 > 0 Then
            c = MEMMSG_CONSUME
            v = .ClienData.mData2
            .ClienData.mData1 = MEMMSG_SUCCESS
          Else
            .ClienData.mData1 = MEMMSG_ERROR
          End If
          .ClienData.mData2 = 0
          updMem = True
        Case MEMMSG_RELEASE
          If MemItems > 0 Then
            c = MEMMSG_RELEASE
            .ClienData.mData1 = MEMMSG_SUCCESS
          Else
            .ClienData.mData1 = MEMMSG_ERROR
          End If
          updMem = True
        Case MEMMSG_EXIT
          ExitNow = True
          .ClienData.mData1 = MEMMSG_SUCCESS
          updMem = True
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

Private Sub ClientTimer_Timer()
  ClientTimer.Enabled = False
  If ExitNow Then Unload Me
  CheckAppMessages
  If MemItems > 0 Then GoThroughMemory
  If Not ClientTimer Is Nothing Then ClientTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  KillClientTimer
  ReleaseMemory
  SharedMemory.Instances(SharedMemOffset).AppData.mData2 = 0
  Call WriteToSharedMemory
  UnloadAll
End Sub
