Attribute VB_Name = "Module1"
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Const SYNCHRONIZE = &H100000

Function SHELLGETTEXT(PROGRAM As String, Optional SHOCMD As Long = vbMinimizedNoFocus) As String
Dim SFILE As String
Dim HFILE As String
Dim ILENGTH As Long
Dim PID As Long
Dim HPROCESS As Long

SFILE = Space(1024)
ILENGTH = GetTempFileName(Environ("TEMP"), "OUT", 0, SFILE)
SFILE = Left(SFILE, ILENGTH)
PID = Shell(Environ("COMSPEC") & " /C" & PROGRAM & ">" & SFILE, SHOCMD)
HPROCESS = OpenProcess(SYNCHRONIZE, True, PID)
WaitForSingleObject HPROCESS, -1
CloseHandle HPROCESS

HFILE = FreeFile
Open SFILE For Binary As #HFILE
SHELLGETTEXT = Input$(LOF(HFILE), HFILE)
Close #HFILE
Kill SFILE
End Function

