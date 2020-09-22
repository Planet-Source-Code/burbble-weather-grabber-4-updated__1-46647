Attribute VB_Name = "modSysTray"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    sTip As String * 64
End Type

Public SysTrayIcon As NOTIFYICONDATA

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const SW_RESTORE = 9
Public Const SW_MINIMIZE = 6
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Function SetIcon(IconPath As String)
frmIcon.Icon = LoadPicture(IconPath)
SysTrayIcon.hIcon = frmIcon.Icon
Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
End Function

Public Function SetIconToolTip(Text As String)
SysTrayIcon.sTip = Text & vbNullChar
'Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
End Function
