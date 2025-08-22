Attribute VB_Name = "modSharedMemory"
Option Explicit

Global Const MEMMSG_CONSUME As Long = &H1
Global Const MEMMSG_RELEASE As Long = &H2
Global Const MEMMSG_EXIT As Long = &H3
Global Const MEMMSG_SUCCESS As Long = &HFFFF
Global Const MEMMSG_ERROR As Long = &HFFAA

Private Const PAGE_READWRITE As Long = &H4&
Private Const FILE_MAP_ALL_ACCESS As Long = &HF001F

Private Const SHAREDMEM_SIZE As Long = SIZE_KILO * 16
Global Const SHAREDMEM_NAME = "Local\UbeMemEater"
Private Const SHAREDMEM_DATASIZE As Long = 16
Private Const SHAREDMEM_HALFSIZE As Long = SHAREDMEM_DATASIZE / 2

Public Type SHAREDMEM_DATA
  mData1 As Long
  mData2 As Long
End Type
Public Type SHAREDMEM_ITEM
  ClienData As SHAREDMEM_DATA
  AppData As SHAREDMEM_DATA
End Type
Public Type SHARED_MEMORY_LAYOUT
  Instances(0 To 1023) As SHAREDMEM_ITEM
End Type

Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Declare Sub CopyMemoryByVal Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Any, ByVal Length As Long)
    
Dim SharedMemHandle As Long
Dim SharedMemBase As Long

Global SharedMemOffset As Long
Global SharedMemory As SHARED_MEMORY_LAYOUT

Global ActiveClients As Long

Public Function OpenSharedMemory() As Boolean
  Dim e As Boolean
  SharedMemHandle = CreateFileMapping(INVALID_HANDLE_VALUE, 0, PAGE_READWRITE, 0, SHAREDMEM_SIZE, SHAREDMEM_NAME)
  If SharedMemHandle = 0 Then Exit Function
  If Err.LastDllError() = ERROR_ALREADY_EXISTS Then e = True
  'If GetLastError() = ERROR_ALREADY_EXISTS Then e = True
  SharedMemBase = MapViewOfFile(SharedMemHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0)
  If SharedMemBase = 0 Then
    Call CloseSharedMemory
    Exit Function
  End If
  If e = False Then
    ClearSharedMemory
  Else
    Call ReadFromSharedMemory(True)
  End If
  OpenSharedMemory = True
End Function

Public Function CloseSharedMemory() As Boolean
  Dim r As Boolean
  If SharedMemBase <> 0 Then
    Call UnmapViewOfFile(SharedMemBase)
    SharedMemBase = 0
    r = True
  End If
  If SharedMemHandle <> 0 Then
    Call CloseHandle(SharedMemHandle)
    SharedMemHandle = 0
    r = True
  End If
  CloseSharedMemory = r
End Function

Public Function WriteToSharedMemory(Optional WriteAllData As Boolean = False, Optional WriteOnlyClientData As Boolean = False, Optional WriteOnlyAppData As Boolean = False, Optional bOffset As Long = -1) As Boolean
  If SharedMemBase = 0 Then Exit Function
  If WriteAllData = True Then
    CopyMemoryByVal SharedMemBase, SharedMemory, LenB(SharedMemory)
  Else
    Dim mAddr As Long, mOff As Long
    mOff = IIf(bOffset = -1, SharedMemOffset, bOffset)
    If (mOff < LBound(SharedMemory.Instances) Or mOff > UBound(SharedMemory.Instances)) Then Exit Function
    mAddr = (SharedMemBase + (mOff * SHAREDMEM_DATASIZE))
    If WriteOnlyClientData = True Then
      CopyMemoryByVal mAddr, SharedMemory.Instances(mOff).ClienData, LenB(SharedMemory.Instances(mOff).ClienData)
    ElseIf WriteOnlyAppData = True Then
      mAddr = (mAddr + SHAREDMEM_HALFSIZE)
      CopyMemoryByVal mAddr, SharedMemory.Instances(mOff).AppData, LenB(SharedMemory.Instances(mOff).AppData)
    Else
      CopyMemoryByVal mAddr, SharedMemory.Instances(mOff), LenB(SharedMemory.Instances(mOff))
    End If
  End If
  WriteToSharedMemory = True
End Function

Public Function ReadFromSharedMemory(Optional ReadAllData As Boolean = False, Optional ReadOnlyClientData As Boolean = False, Optional ReadOnlyAppData As Boolean = False, Optional bOffset As Long = -1) As Boolean
  If SharedMemBase = 0 Then Exit Function
  If ReadAllData = True Then
    CopyMemory SharedMemory, SharedMemBase, LenB(SharedMemory)
  Else
    Dim mAddr As Long, mOff As Long
    mOff = IIf(bOffset = -1, SharedMemOffset, bOffset)
    If (mOff < LBound(SharedMemory.Instances) Or mOff > UBound(SharedMemory.Instances)) Then Exit Function
    mAddr = (SharedMemBase + (mOff * SHAREDMEM_DATASIZE))
    If ReadOnlyClientData = True Then
      CopyMemory SharedMemory.Instances(mOff).ClienData, mAddr, LenB(SharedMemory.Instances(mOff).ClienData)
    ElseIf ReadOnlyAppData = True Then
      mAddr = (mOff + SHAREDMEM_HALFSIZE)
      CopyMemory SharedMemory.Instances(mOff).AppData, mAddr, LenB(SharedMemory.Instances(mOff).AppData)
    Else
      CopyMemory SharedMemory.Instances(mOff), mAddr, LenB(SharedMemory.Instances(mOff))
    End If
  End If
  ReadFromSharedMemory = True
End Function

Private Sub ClearSharedMemory()
  ZeroMemory SharedMemory, LenB(SharedMemory)
  Call WriteToSharedMemory(True)
End Sub

Public Sub ClearMemoryData(MemData As SHAREDMEM_DATA)
  With MemData
    .mData1 = 0
    .mData2 = 0
  End With
End Sub

Public Function GetNextAvailableOffset() As Long
  Dim i As Long
  For i = 1 To UBound(SharedMemory.Instances)
    If SharedMemory.Instances(i).AppData.mData2 = 0 Then
      GetNextAvailableOffset = i
      SharedMemory.Instances(i).AppData.mData2 = -1
      WriteToSharedMemory False, False, True, i
      Exit For
    End If
  Next
End Function

Public Sub CheckActiveProcesses()
  Dim i As Long, c As Long
  Call ReadFromSharedMemory(True)
  c = 0
  With SharedMemory
    For i = 1 To UBound(.Instances)
      If .Instances(i).AppData.mData2 <> 0 Then
        If .Instances(i).AppData.mData2 = -1 Then
          'Ignore - do nothing...
        Else
          If IsProcessAlive(.Instances(i).AppData.mData2) = False Then
            .Instances(i).AppData.mData2 = 0
            GoSub UpdateMem
          Else
            c = (c + 1)
          End If
        End If
      End If
    Next
  End With
  ActiveClients = c
  Exit Sub
UpdateMem:
  Call WriteToSharedMemory(False, False, True, i)
  Return
End Sub

Public Function CountActiveProcesses() As Long
  Dim i As Long, r As Long
  With SharedMemory
    For i = 1 To UBound(.Instances)
      Select Case .Instances(i).AppData.mData2
        Case Is <> 0, Not -1
          r = (r + 1)
      End Select
    Next
  End With
  CountActiveProcesses = r
End Function

Public Sub ReleaseAllClients()
  Dim i As Long
  With SharedMemory
    For i = 1 To UBound(.Instances)
      .Instances(i).ClienData.mData1 = MEMMSG_EXIT
      'Call WriteToSharedMemory(False, True, False, i)
    Next
  End With
  Call WriteToSharedMemory(True)
End Sub

Public Sub ClientConsumeMemory(cIndex As Long, BytesToConsume As Long)
  If (cIndex <= 0 Or cIndex >= UBound(SharedMemory.Instances)) Then Exit Sub
  With SharedMemory.Instances(cIndex).ClienData
    .mData1 = MEMMSG_CONSUME
    .mData2 = BytesToConsume
  End With
  Call WriteToSharedMemory(False, True, False, cIndex)
End Sub
