Attribute VB_Name = "modSystem"
Option Explicit

Global Const INVALID_HANDLE_VALUE As Long = -1&

Private Const HWND_NOTOPMOST  As Long = -2
Private Const HWND_TOPMOST  As Long = -1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Private Const SWP_SETWINDOWPOS  As Long = SWP_NOSIZE Or SWP_NOMOVE

Private Const ICC_LISTVIEW_CLASSES  As Long = &H1
Private Const ICC_TREEVIEW_CLASSES  As Long = &H2
Private Const ICC_BAR_CLASSES  As Long = &H4
Private Const ICC_TAB_CLASSES  As Long = &H8
Private Const ICC_UPDOWN_CLASS  As Long = &H10
Private Const ICC_PROGRESS_CLASS  As Long = &H20
Private Const ICC_HOTKEY_CLASS  As Long = &H40
Private Const ICC_ANIMATE_CLASS  As Long = &H80
Private Const ICC_WIN95_CLASSES  As Long = &HFF
Private Const ICC_DATE_CLASSES  As Long = &H100
Private Const ICC_USEREX_CLASSES  As Long = &H200
Private Const ICC_COOL_CLASSES  As Long = &H400
Private Const ICC_INTERNET_CLASSES  As Long = &H800
Private Const ICC_PAGESCROLLER_CLASS  As Long = 1000
Private Const ICC_NATIVEFNTCTL_CLASS  As Long = 2000
Private Const ICC_STANDARD_CLASSES  As Long = 4000
Private Const ICC_LINK_CLASS  As Long = 8000

Private Const SYNCHRONIZE As Long = &H100000
Private Const WAIT_OBJECT_0 As Long = 0&
Private Const WAIT_TIMEOUT As Long = &H102&

Global Const ERROR_ALREADY_EXISTS As Long = 183&

Public Enum COMMONCONTROLS_CLASSES
  ccListView_Classes = ICC_LISTVIEW_CLASSES
  ccTreeView_Classes = ICC_TREEVIEW_CLASSES
  ccToolBar_Classes = ICC_BAR_CLASSES
  ccTab_Classes = ICC_TAB_CLASSES
  ccUpDown_Classes = ICC_UPDOWN_CLASS
  ccProgress_Class = ICC_PROGRESS_CLASS
  ccHotkey_Class = ICC_HOTKEY_CLASS
  ccAnimate_Class = ICC_ANIMATE_CLASS
  ccWin95_Classes = ICC_WIN95_CLASSES
  ccCalendar_Classes = ICC_DATE_CLASSES
  ccComboEx_Classes = ICC_USEREX_CLASSES
  ccCoolBar_Classes = ICC_COOL_CLASSES
  ccInternet_Classes = ICC_INTERNET_CLASSES
  ccPageScroller_Class = ICC_PAGESCROLLER_CLASS
  ccNativeFont_Class = ICC_NATIVEFNTCTL_CLASS
  ccStandard_Classes = ICC_STANDARD_CLASSES
  ccLink_Class = ICC_LINK_CLASS
  ccAll_Classes = ccListView_Classes Or ccTreeView_Classes Or ccToolBar_Classes Or ccTab_Classes Or ccUpDown_Classes Or ccProgress_Class Or ccHotkey_Class Or ccAnimate_Class Or ccWin95_Classes Or ccCalendar_Classes Or ccComboEx_Classes Or ccCoolBar_Classes Or ccInternet_Classes Or ccPageScroller_Class Or ccNativeFont_Class Or ccStandard_Classes Or ccLink_Class
 End Enum

Private Type tagINITCOMMONCONTROLSEX
  dwSize As Long
  dwICC As Long
End Type

Public Type POINTAPI
  x As Long
  Y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, Param As Any) As Long
Private Declare Sub InitCommonControls9x Lib "comctl32" Alias "InitCommonControls" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByVal lpSource As Any, ByVal cbCopy As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function GetMypId() As Long
  GetMypId = GetCurrentProcessId()
End Function

Public Function IsProcessAlive(ProcessId As Long) As Boolean
  Dim hProc As Long, r As Long
  If ProcessId = 0 Then Exit Function
  hProc = OpenProcess(SYNCHRONIZE, 0, ProcessId)
  If hProc = 0 Then Exit Function
  r = WaitForSingleObject(hProc, 0)
  Call CloseHandle(hProc)
  IsProcessAlive = (r = WAIT_TIMEOUT)
End Function

Public Sub WindowOnTop(hWnd As Long, OnTop As Boolean)
  Dim wFlags As Long
  If OnTop Then
    wFlags = HWND_TOPMOST
  Else
    wFlags = HWND_NOTOPMOST
  End If
  SetWindowPos hWnd, wFlags, 0&, 0&, 0&, 0&, SWP_SETWINDOWPOS
End Sub

Public Function InitCommonControls(Optional ccFlags As COMMONCONTROLS_CLASSES = ccAll_Classes) As Boolean
  Dim icc As tagINITCOMMONCONTROLSEX
  On Error GoTo OldCC
  With icc
    .dwSize = Len(icc)
    .dwICC = ccFlags
  End With
  InitCommonControls = InitCommonControlsEx(icc)
ExitNow:
  On Error GoTo 0
  Exit Function
OldCC:
  InitCommonControls9x
  Resume ExitNow
End Function

Public Function IsPointInRect(pRect As RECT, pPoint As POINTAPI) As Boolean
  With pRect
    If (pPoint.x >= .Left And pPoint.x <= .Right) Then
      If (pPoint.Y >= .Top And pPoint.Y <= .Bottom) Then IsPointInRect = True
    End If
  End With
End Function
