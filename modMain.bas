Attribute VB_Name = "modMain"
Option Explicit

Private Const CMD_SEP_COMMAND As String = ";"
Private Const CMD_SEP_VALUE As String = ":"
Private Const CMD_RUN As String = "gbkgb"
Private Const CMD_CONSUME As String = "kijed"
Private Const CMD_RESERVE As String = "edxcdtj"
Private Const CMD_MEMOFFSET As String = "bhirn"

Global Const APP_NAME As String = "Ubehage's MemEater v2"

Global Const TIMER_INTERVAL_DISPLAY = 300
Global Const TIMER_INTERVAL_CLIENT = 500

Global Const FONT_MAIN As String = "Segoe UI"
Global Const FONT_SECONDARY As String = "Consolas"
Global Const FONTSIZE_MAIN As Integer = 11
Global Const FONTSIZE_SECONDARY As Integer = 9

Global Const COLOR_BACKGROUND As Long = 2105376
Global Const COLOR_CONTROLS As Long = 2763306
Global Const COLOR_BUTTON_HOVER As Long = 3684408
Global Const COLOR_BUTTON_PRESSED As Long = 3289650
Global Const COLOR_BACKGROUND_DISABLED As Long = 5263440
Global Const COLOR_TEXT_DISABLED As Long = 7895160
Global Const COLOR_TEXT As Long = 14737632
Global Const COLOR_OUTLINE As Long = 3815994
Global Const COLOR_OUTLINE_LIGHT As Long = 7368816
Global Const COLOR_GREEN As Long = 5023791
Global Const COLOR_YELLOW As Long = 4965861
Global Const COLOR_RED As Long = 4539862

Global Const SIZE_KILO As Long = 1024
Global Const SIZE_MEGA As Long = SIZE_KILO * SIZE_KILO
Global Const SIZE_GIGA As Long = SIZE_MEGA * SIZE_KILO

Global MaxMemoryGB As Long

Global ExitNow As Boolean

Dim RunNow As Boolean

Sub Main()
  Dim s As Integer
  Call InitCommonControls
  s = Start
  If s = 0 Then Exit Sub
  SharedMemory.Instances(SharedMemOffset).AppData.mData2 = GetMypId
  Call WriteToSharedMemory(False, False, True)
  Select Case s
    Case 1
      LoadMainForm
    Case 2
      StartClient
  End Select
End Sub

Private Function Start() As Integer
  SplitCommandLine Command
  
  '#Debug
  'SplitCommandLine GetCommandLineParameters(1)
  
  If RunNow Then
    If SharedMemOffset = 0 Then
      Call MsgBox("Error: Missing shared memory!" & vbCrLf & "This program cannot continue.", vbOKOnly Or vbCritical, APP_NAME)
      Exit Function
    End If
    If OpenSharedMemory() = False Then
      Call MsgBox("There was a problem reading shared memory." & vbCrLf & "This program cannot continue.", vbOKOnly Or vbCritical, APP_NAME)
      UnloadAll
      Exit Function
    End If
    Start = 2
  Else
    If SharedMemOffset <> 0 Then
      Call MsgBox("Error in command-line!" & vbCrLf & vbCrLf & "This program cannot continue.", vbOKOnly Or vbCritical, APP_NAME)
      Exit Function
    End If
    If OpenSharedMemory() = False Then
      Call MsgBox("There was a problem reading shared memory." & vbCrLf & "This program cannot continue.", vbOKOnly Or vbCritical, APP_NAME)
      Exit Function
    End If
    If CheckPrevInstance = False Then
      MsgBox "You may only run one instance of this program!", vbOKOnly Or vbInformation, APP_NAME
      Exit Function
    End If
    Start = 1
  End If
End Function

Private Function CheckPrevInstance() As Boolean
  If App.PrevInstance = True Then
    With SharedMemory.Instances(0).AppData
      If .mData2 <> 0 Then
        If IsProcessAlive(.mData2) Then Exit Function
      End If
    End With
  End If
  CheckPrevInstance = True
End Function

Private Sub LoadMainForm()
  SetMaxMemory
  Load frmMain
  frmMain.SetForm
  frmMain.SetDisplayTimer
End Sub

Private Sub StartClient()
  Load frmClient
  frmClient.SetClientTimer
End Sub

Public Sub UnloadAll()
  CloseSharedMemory
End Sub

Private Sub SetMaxMemory()
  Dim m As INTERNAL_MEM_STATUS, x As Long
  Call GetMemoryInfo(m)
  MaxMemoryGB = Int(m.PhysicalMemory.Max / SIZE_GIGA)
End Sub

Public Function GetByteSizeString(Bytes As Currency, Optional IncludePostfix As Boolean = True) As String
  Dim bV As Currency, bN As String, r As String
  bV = Bytes
  If bV >= 1024 Then
    bV = (bV / 1024)
    If bV >= 1024 Then
      bV = (bV / 1024)
      If bV >= 1024 Then
        bV = (bV / 1024)
        If bV >= 1024 Then
          bV = (bV / 1024)
          bN = "TB"
        Else
          bN = "GB"
        End If
      Else
        bN = "MB"
      End If
    Else
      bN = "KB"
    End If
  Else
    bN = "B"
  End If
  If bV = 1 Then
    bN = Left$(bN, (Len(bN) - 1))
  End If
  r = RoundByteSizeToString(bV)
  If IncludePostfix = True Then r = r & bN
  GetByteSizeString = r
End Function

Private Function RoundByteSizeToString(ByteValue As Currency) As String
  Dim i As Long
  Dim bV As String
  bV = CStr(ByteValue)
  i = InStr(bV, ",")
  If i > 0 Then
    RoundByteSizeToString = Left$(bV, (i + 1))
  Else
    RoundByteSizeToString = bV
  End If
End Function

Public Function GetBooleanValueFromString(BoolString As String) As Boolean
  Select Case LCase$(Trim$(BoolString))
    Case "true"
      GetBooleanValueFromString = True
    Case Else
      GetBooleanValueFromString = False
  End Select
End Function

Public Function GetStringFromBoolean(BoolValue As Boolean) As String
  Select Case BoolValue
    Case True
      GetStringFromBoolean = "True"
    Case Else
      GetStringFromBoolean = "False"
  End Select
End Function

Private Function GetAppFile() As String
  GetAppFile = FixPath(App.Path) & "\" & App.EXEName & ".exe"
End Function

Private Function FixPath(Path As String) As String
  Dim i As Long
  i = Len(Path)
  While Mid$(Path, i, 1) = "\"
    i = (i - 1)
  Wend
  FixPath = Left$(Path, i)
End Function

Public Function LaunchNewClient() As Long
  Dim i As Long
  i = GetNextAvailableOffset()
  If i = 0 Then Exit Function
  If LaunchClientWithIndex(i) = True Then LaunchNewClient = i
End Function

Public Function LaunchClientWithIndex(NewIndex As Long) As Boolean
  If NewIndex <= 0 Then Exit Function
  LaunchClientWithIndex = RunCommandLine(GetNewCommandLine(NewIndex))
End Function

Public Function GetNextCommandLine() As String
  GetNextCommandLine = GetNewCommandLine(GetNextAvailableOffset())
End Function

Public Function GetNewCommandLine(iOffset As Long) As String
  GetNewCommandLine = """" & GetAppFile & """ " & GetCommandLineParameters(iOffset)
End Function

Private Function GetCommandLineParameters(iOffset As Long) As String
  Dim r As String
  r = CMD_RUN
  If iOffset > 0 Then r = r & CMD_SEP_COMMAND & CMD_MEMOFFSET & CMD_SEP_VALUE & CStr(iOffset)
  GetCommandLineParameters = r
End Function

Private Function RunCommandLine(CommandLine As String) As Boolean
  On Error GoTo ShellError
  RunCommandLine = (Shell(CommandLine, vbNormalFocus) <> 0)
ShellExit:
  On Error GoTo 0
  Exit Function
ShellError:
  Resume ShellExit
End Function

Private Sub SplitCommandLine(CommandLine As String)
  Dim i As Integer, cArr() As String, c As String, v As String, vArr() As String
  cArr() = Split(CommandLine, CMD_SEP_COMMAND)
  For i = LBound(cArr) To UBound(cArr)
    c = cArr(i)
    vArr = Split(c, CMD_SEP_VALUE)
    If UBound(vArr) = 1 Then
      c = vArr(0)
      v = vArr(1)
    End If
    Select Case c
      Case CMD_RUN
        RunNow = True
      Case CMD_CONSUME
        'do something...
      Case CMD_RESERVE
        'do something...
      Case CMD_MEMOFFSET
        If Not (v = "" Or v = "0") Then SharedMemOffset = CLng(v)
    End Select
  Next
End Sub
