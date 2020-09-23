VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Win32.Blaster Detection Utility"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      BackColor       =   &H8000000F&
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblHyperlink 
      Alignment       =   2  'Center
      Caption         =   "http://www.sanx.org/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      MousePointer    =   4  'Icon
      TabIndex        =   2
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************
' **            Win32.Blaster Detection Utility                     **
' **                Copyright 2003, www.sanx.org                    **
' ********************************************************************

Private Sub Form_Load()

DoEvents
CheckSystem

End Sub

Private Sub CheckSystem()

Dim wshMain As WshShell
Dim fsoMain As FileSystemObject
Dim regMain
Dim filExists As Boolean
Dim count As Integer, pid As Integer, retCode
Dim strExeName As String, strRegName As String

Set wshMain = New WshShell
Set fsoMain = New FileSystemObject
strExeName = "msblast.exe"
strRegName = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\windows auto update"

filExists = fsoMain.FileExists(fsoMain.GetSpecialFolder(1) & "\" & strExeName)

lblMain.Caption = "Checking system..."
DoEvents

UpdateInfo "Looking for worm executable..."
If filExists = True Then
    UpdateInfo "Virus found!"
    DoEvents
    UpdateInfo "Looking for active process..."
    pid = ListTasks(strExeName)
    DoEvents
    Do
        If pid > -1 Then
            UpdateInfo "Active process found. Terminating..."
            retCode = KillProcess(pid)
        Else
            UpdateInfo "Process not found. Deleting virus executable..."
        End If
        pid = ListTasks(strExeName)
    Loop Until pid < 0
    If DeleteFile(fsoMain.GetSpecialFolder(1) & "\" & strExeName) = True Then
        UpdateInfo "Virus executable deleted."
    Else
        UpdateInfo "Unable to delete executable. Exiting program"
        lblMain.ForeColor = &HFF&
        lblMain.Caption = "ERROR!"
        Exit Sub
    End If
Else
    UpdateInfo "Virus executable not found."
End If

Select Case CheckReg(strRegName)
    Case 1
        UpdateInfo "Registry key deleted."
    Case 0
        UpdateInfo "Virus auto-loader not found..."
    Case -1
        Exit Sub
End Select

UpdateInfo "Finished..."
lblMain.ForeColor = &HC000&
lblMain.Caption = "ALL CLEAR!"


End Sub

Function DeleteFile(strExeName As String) As Boolean

On Error GoTo ErrHandle

Dim fsoMain As FileSystemObject
Dim objFile As File
Set fsoMain = New FileSystemObject
Set objFile = fsoMain.GetFile(strExeName)

objFile.Delete True

DeleteFile = True

Exit Function

ErrHandle:
DeleteFile = False

End Function

Function CheckReg(strRegValue As String) As Integer

On Error GoTo ErrHandle

Dim wshMain As WshShell
Set wshMain = New WshShell

UpdateInfo "Checking registry..."
wshMain.RegRead strRegValue

On Error GoTo NoDelete

wshMain.RegDelete strRegValue
CheckReg = 1

Exit Function

ErrHandle:
CheckReg = 0

Exit Function

NoDelete:
UpdateInfo "Error deleting registry key."
lblMain.ForeColor = &HFF&
lblMain.Caption = "ERROR!"
CheckReg = -1

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

End

End Sub

Private Sub lblHyperlink_Click()

ShellExecute 0&, "Open", "http:\\www.sanx.org\", "", vbNullString, SW_SHOWNORMAL

End Sub
