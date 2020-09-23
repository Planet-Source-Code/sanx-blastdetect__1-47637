Attribute VB_Name = "modMain"
'Private constants
Const HSHELL_ACTIVATESHELLWINDOW = 3
Const HSHELL_WINDOWCREATED = 1
Const HSHELL_WINDOWDESTROYED = 2
Const HSHELL_WINDOWACTIVATED = 4
Const HSHELL_GETMINRECT = 5
Const HSHELL_REDRAW = 6
Const HSHELL_TASKMAN = 7
Const HSHELL_LANGUAGE = 8
Const HSHELL_ACCESSIBILITYSTATE = 11
Const LOCALE_SENGLANGUAGE As Long = &H1001

'Public Constants
Public Const GWL_WNDPROC = (-4)
Public Const RSH_DEREGISTER = 0
Public Const RSH_REGISTER = 1
Public Const RSH_REGISTER_PROGMAN = 2
Public Const RSH_REGISTER_TASKMAN = 3
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_TERMINATE = &H1&
Public Const PROCESS_CREATE_THREAD = &H2&
Public Const PROCESS_VM_OPERATION = &H8&
Public Const PROCESS_VM_READ = &H10&
Public Const PROCESS_VM_WRITE = &H206
Public Const PROCESS_DUP_HANDLE = &H40&
Public Const PROCESS_CREATE_PROCESS = &H80&
Public Const PROCESS_SET_QUOTA = &H100&
Public Const PROCESS_SET_INFORMATION = &H200&
Public Const PROCESS_QUERY_INFORMATION = &H400&
Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

'Type Definitions
Type ProcessEntry
    dwSize As Long
    peUsage As Long
    peProcessID As Long
    peDefaultHeapID As Long
    peModuleID As Long
    peThreads As Long
    peParentProcessID As Long
    pePriority As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

'Local Variables
Dim hnd                             As Long
Dim lRet                            As Long
Dim lExitCode                       As Long
Dim lPriority                       As Long
Dim exePriority                     As Long

'Public Variables
Public OldProc                      As Long
Public uRegMsg                      As Long

'API Declarations
Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwIdProc As Long) As Long
Declare Function Process32First Lib "kernel32" (ByVal hndl As Long, ByRef pstru As ProcessEntry) As Boolean
Declare Function Process32Next Lib "kernel32" (ByVal hndl As Long, ByRef pstru As ProcessEntry) As Boolean
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hnd As Long) As Boolean
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Function ListTasks(strExeName As String) As Integer

Dim iIdx            As Integer
Dim bRet            As Boolean
Dim lSnapShot       As Long
Dim tmpPE           As ProcessEntry

Dim intProcesses    As Integer
Dim intThreads      As Integer

Dim tmpProcName     As String
Dim tmpPriority     As String

   
    lSnapShot = CreateToolhelp32Snapshot(&H2, 0)
    tmpPE.dwSize = Len(tmpPE)
    bRet = Process32First(lSnapShot, tmpPE)
    ListTasks = -1
    
    Do Until bRet = False
        
        tmpProcName = LCase(Mid(tmpPE.szExeFile, _
                            InStrRev(tmpPE.szExeFile, "\", Len(tmpPE.szExeFile)) + 1, _
                            Len(tmpPE.szExeFile) - InStrRev(tmpPE.szExeFile, "\", 1)))
        tmpProcName = Left(tmpProcName, InStr(1, tmpProcName, Chr(0)) - 1)
        

        If tmpProcName = strExeName Then
            ListTasks = tmpPE.peProcessID
            Exit Function
        End If
        
        intProcesses = intProcesses + 1
        intThreads = intThreads + tmpPE.peThreads

        bRet = Process32Next(lSnapShot, tmpPE)
        DoEvents

    Loop
    
    bRet = CloseHandle(lSnapShot)

End Function

Public Sub UpdateInfo(strInfo As String)

frmMain.txtInfo.Text = frmMain.txtInfo.Text + strInfo + vbCrLf
frmMain.txtInfo.SelStart = Len(frmMain.txtInfo.Text)

End Sub

Public Function KillProcess(pid As Integer) As Boolean

Dim retCode As Integer, proHandle As Long

proHandle = OpenProcess(PROCESS_TERMINATE, False, pid)
Call TerminateProcess(proHandle, retCode)
Call CloseHandle(proHandle)

End Function

