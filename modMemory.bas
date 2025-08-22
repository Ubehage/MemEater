Attribute VB_Name = "modMemory"
Option Explicit

Private Type ULARGE_INTEGER '64bit data structure
  LowPart As Long
  HighPart As Long
End Type

Private Type MEMORY_STATUS_EX
  dwLength As Long 'Size of the data structure
  dwMemoryLoad As Long 'Percentage of used memory
  ullTotalPhys As ULARGE_INTEGER 'Total amount of physical memory in bytes
  ullAvailPhys As ULARGE_INTEGER 'Available physical memory in bytes
  ullTotalPageFile As ULARGE_INTEGER 'Total amount of pagefile-memory in bytes
  ullAvailPageFile As ULARGE_INTEGER 'Available pagefile-memory in bytes
  ullTotalVirtual As ULARGE_INTEGER 'Total amount of virtual memory in bytes
  ullAvailVirtual As ULARGE_INTEGER 'Available virtual memory in bytes
  ullAvailExtendedVirtual As ULARGE_INTEGER 'Available extended virtual memory in bytes
End Type

Private Type PERFORMANCE_INFORMATION
  cB As Long
  commitTotal As Long
  commitLimit As Long
  CommitPeak As Long
  PhysicalTotal As Long
  PhysicalAvailable As Long
  SystemCache As Long
  KernelTotal As Long
  KernelPaged As Long
  KernelNonPaged As Long
  pageSize As Long
  HandleCount As Long
  ProcessCount As Long
  ThreadCount As Long
End Type

Public Type Max_Used
  Max As Currency
  Used As Currency
End Type

Public Type INTERNAL_MEM_STATUS
  PhysicalMemory As Max_Used
  VirtualMemory As Max_Used
End Type

Private Declare Function GetPerformanceInfo Lib "psapi.dll" (ByRef pPerformanceInformation As PERFORMANCE_INFORMATION, ByVal cB As Long) As Long
Private Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (lpBuffer As MEMORY_STATUS_EX) As Long

Public Sub GetMemoryInfo(TargetInfo As INTERNAL_MEM_STATUS)
  Dim MemTemp As Currency
  Dim MemInfo As MEMORY_STATUS_EX
  MemInfo.dwLength = Len(MemInfo)
  Call GlobalMemoryStatusEx(MemInfo)
  With TargetInfo
    With .PhysicalMemory
      .Max = ULargeIToCurrency(MemInfo.ullTotalPhys)
      .Used = (.Max - ULargeIToCurrency(MemInfo.ullAvailPhys))
    End With
    Call GetPageFileInfo(.VirtualMemory)
    'With .VirtualMemory
    '  .Max = ULargeIToCurrency(MemInfo.ullTotalPageFile)
    '  .Used = (.Max - ULargeIToCurrency(MemInfo.ullAvailPageFile))
    'End With
  End With
End Sub

Private Function ULargeIToCurrency(ByRef u As ULARGE_INTEGER) As Currency
  Dim hi As Double, lo As Double
  lo = (u.LowPart And &H7FFFFFFF) + IIf(u.LowPart < 0, 2147483648#, 0)
  hi = (u.HighPart And &H7FFFFFFF) + IIf(u.HighPart < 0, 2147483648#, 0)
  ULargeIToCurrency = CCur((hi * 4294967296#) + lo)
End Function

Private Sub GetPageFileInfo(MaxUsed As Max_Used)
  Dim pI As PERFORMANCE_INFORMATION
  pI.cB = Len(pI)
  If GetPerformanceInfo(pI, pI.cB) <> 0 Then
    Dim physTotal As Currency, physAvail As Currency, commitUsed As Currency, commitLimit As Currency, pageSize As Currency
    pageSize = pI.pageSize
    MaxUsed.Max = PagesToBytes(pI.commitLimit, pageSize)
    MaxUsed.Used = PagesToBytes(pI.commitTotal, pageSize)
    'physTotal = PagesToBytes(pI.PhysicalTotal, pageSize)
    'physAvail = PagesToBytes(pI.PhysicalAvailable, pageSize)
    'commitUsed = PagesToBytes(pI.commitTotal, pageSize)
    'commitLimit = PagesToBytes(pI.commitLimit, pageSize)
    'MaxUsed.Max = (commitLimit - physTotal)
    'MaxUsed.Used = (commitUsed - (physTotal - physAvail))
  End If
End Sub

Private Function PagesToBytes(ByVal pages As Long, ByVal pageSize As Long) As Currency
  PagesToBytes = CCur(CDbl(pages) * CDbl(pageSize))
End Function
