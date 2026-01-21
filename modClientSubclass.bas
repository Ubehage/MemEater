Attribute VB_Name = "modClientSubclass"
Option Explicit

Private Const GWL_WNDPROC = -4
Private Const WM_POWERBROADCAST = &H218 'Power state is changing
Private Const PBT_APMSUSPEND As Long = &H4 'Windows is hibernating
Private Const PBT_APMRESUMEAUTOMATIC = &H12 'Windows has resumed from suspended state.

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim OldWndProc As Long

Public Sub HookClientWindowProc(hWnd As Long)
  OldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ClientWndProc)
End Sub

Public Sub UnhookClient(hWnd As Long)
  If OldWndProc <> 0 Then
    Call SetWindowLong(hWnd, GWL_WNDPROC, OldWndProc)
    OldWndProc = 0
  End If
End Sub

Public Function ClientWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If uMsg = WM_POWERBROADCAST Then
    If wParam = PBT_APMSUSPEND Then
      frmClient.ToggleHibernate True
    ElseIf wParam = PBT_APMRESUMEAUTOMATIC Then
      frmClient.ToggleHibernate False
    End If
  End If
  ClientWndProc = CallWindowProc(OldWndProc, hWnd, uMsg, wParam, lParam)
End Function
