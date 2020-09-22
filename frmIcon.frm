VERSION 5.00
Begin VB.Form frmIcon 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   2565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "00°"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "This stuff is for creating the system tray icons."
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblIcon 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuShow 
         Caption         =   "Show Weather Grabber"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide Weather Grabber"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit Weather Grabber"
      End
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
SavePicture frmIcon.Icon, App.Path & "\Icon.ico"
SysTrayIcon.cbSize = Len(SysTrayIcon)
SysTrayIcon.hwnd = frmIcon.hwnd
SysTrayIcon.uID = vbNull
SysTrayIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
SysTrayIcon.uCallbackMessage = WM_MOUSEMOVE
SysTrayIcon.hIcon = frmIcon.Icon
SysTrayIcon.sTip = "Weather Grabber" & vbNullChar
Shell_NotifyIcon NIM_ADD, SysTrayIcon
If HideMe = True Then
Load frmMain
frmMain.Visible = False
Else
Load frmMain
frmMain.Visible = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim IconEvent As Long

IconEvent = X / Screen.TwipsPerPixelX

Select Case IconEvent
Case WM_RBUTTONUP
If frmControls.inetData.StillExecuting = True Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub
PopupMenu mnuMain
Case WM_LBUTTONUP
frmMain.Visible = True
If frmControls.inetData.StillExecuting = True Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub
If RegIcon = "1" Then
CheckNow = True
Else
CheckNow = False
End If
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, SysTrayIcon
Set frmIcon = Nothing
Set frmMain = Nothing
Set frmRadar = Nothing
Set frmPreferences = Nothing
Set frmControls = Nothing
Set frmDetails = Nothing
Set frmAlerts = Nothing
Set frmAlertWindow = Nothing
Set frmIP = Nothing
End
End Sub

Private Sub Command2_Click()

'Code for generating icons...
'------
'For i = -20 To 110
'DoEvents
'Text1.Text = i '& "°"
'picIcon.Cls
'lblIcon.Caption = Text1.Text
'picIcon.CurrentX = (picIcon.Width / 2) - (lblIcon.Width / 2)
'picIcon.CurrentY = 45
'picIcon.FontSize = lblIcon.FontSize
'picIcon.FontBold = lblIcon.FontBold
'picIcon.Font = lblIcon.Font
'picIcon.Print Text1.Text
'SavePicture picIcon.Image, App.Path & "\Icon " & i & ".bmp"
'Next i
'------

'I used a freeware program called IconShop to
'convert the BMPs to ICOs.

'Code for testing icons...
'------
'For i = -20 To 110
'Text1.Text = i
'frmIcon.Icon = LoadPicture(App.Path & "\Icons\ICO\Icon " & Text1.Text & ".ico")
'SysTrayIcon.hIcon = frmIcon.Icon
'Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
'Next i
'------
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHide_Click()
If frmControls.inetData.StillExecuting = True Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub

If frmRadar.Visible = True Then
MsgBox "Please close the Radar window.", vbExclamation
Exit Sub
End If
If frmPreferences.Visible = True Then
MsgBox "Please close the Preferences window.", vbExclamation
Exit Sub
End If
If frmForecast.Visible = True Then
MsgBox "Please close the Forecast window.", vbExclamation
Exit Sub
End If
frmMain.Visible = False
End Sub

Private Sub mnuShow_Click()
frmMain.Visible = True
End Sub
