VERSION 5.00
Begin VB.Form frmPreferences 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferences"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox fileSound 
      Height          =   510
      Left            =   2400
      Pattern         =   "*.wav"
      TabIndex        =   29
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CheckBox chkAlert 
      Caption         =   "Show New Alert Window"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CheckBox chkProxy 
      Caption         =   "Use Proxy Server"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdUnits 
      Caption         =   "Change Units"
      Height          =   375
      Left            =   2775
      TabIndex        =   26
      Top             =   3045
      Width           =   1200
   End
   Begin VB.CheckBox chkConnect 
      Caption         =   "Connect to the internet when program starts."
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   3855
   End
   Begin VB.FileListBox fileImages 
      Height          =   300
      Left            =   1440
      Pattern         =   "*.gif"
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkIcon 
      Caption         =   "Refresh on system tray icon click."
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtLineColor 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtFontColor 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4200
      Width           =   2775
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   120
      Picture         =   "frmPreferences.frx":0000
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   16
      Top             =   3480
      Width           =   960
   End
   Begin VB.CommandButton cmdClearCache 
      Caption         =   "Clear Image Cache"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox chkNoCache 
      Caption         =   "Disable image caching"
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CheckBox chkInternet 
      Caption         =   "Only refresh data if connected to the internet."
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox chkRefresh 
      Caption         =   "Auto refresh data"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox lstBackgrounds 
      Height          =   1110
      ItemData        =   "frmPreferences.frx":533E
      Left            =   1680
      List            =   "frmPreferences.frx":5340
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.CheckBox chkLoad 
      Caption         =   "Load data when program starts."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   1
      Left            =   120
      Top             =   480
   End
   Begin VB.VScrollBar vscrollRefresh 
      Height          =   255
      Left            =   510
      Max             =   1
      Min             =   60
      TabIndex        =   6
      Top             =   1095
      Value           =   60
      Width           =   255
   End
   Begin VB.FileListBox fileBackgrounds 
      Height          =   300
      Left            =   2760
      Pattern         =   "*.gif"
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRefresh 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtZipCode 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.OptionButton optFontColor 
      Caption         =   "Font Color:"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   3960
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton optLineColor 
      Caption         =   "Line Color:"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "New Alert Sound:"
      Height          =   495
      Left            =   2400
      TabIndex        =   30
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblBackgrounds 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "minutes."
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1125
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Choose Background:"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Refresh data every:"
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurX As Integer
Dim CurY As Integer
Dim Proxy As String
Dim ProxyUser As String
Dim ProxyPassword As String
Dim ProxyPort As String

Private Sub chkProxy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkProxy.Value = 1 Then

Proxy = InputBox("Enter the address of the proxy server.", , Proxy)

ProxyPort = InputBox("Enter the port for the proxy server (optional).", , ProxyPort)

If Proxy = "" Then
chkProxy.Value = 0
Exit Sub
End If

ProxyUser = InputBox("Enter the user name you want to use (optional).", , ProxyUser)

ProxyPassword = InputBox("Enter the password you want to use (optional).", , ProxyPassword)

End If
End Sub

Private Sub cmdCancel_Click()
JustCanceled = True
Unload Me
End Sub

Private Sub cmdClearCache_Click()
For t = 0 To fileImages.ListCount - 1
If Left(fileImages.List(t), 3) = "52_" Or Left(fileImages.List(t), 3) = "31_" Then Kill App.Path & "\" & fileImages.List(t)
Next t
End Sub

Private Sub cmdOK_Click()
JustCanceled = False
If txtZipCode.Text = RegZip Then
ChangedZip = False
Else
ChangedZip = True
End If
SaveSetting "WeatherGrabber", "Settings", "ZipCode", txtZipCode.Text
SaveSetting "WeatherGrabber", "Settings", "Background", lstBackgrounds.Text
SaveSetting "WeatherGrabber", "Settings", "Refresh", txtRefresh.Text
SaveSetting "WeatherGrabber", "Settings", "LoadAtStart", chkLoad.Value
SaveSetting "WeatherGrabber", "Settings", "EverRefresh", chkRefresh.Value
SaveSetting "WeatherGrabber", "Settings", "Internet", chkInternet.Value
SaveSetting "WeatherGrabber", "Settings", "NoCache", chkNoCache.Value
SaveSetting "WeatherGrabber", "Settings", "FontColor", txtFontColor.Text
SaveSetting "WeatherGrabber", "Settings", "LineColor", txtLineColor.Text
SaveSetting "WeatherGrabber", "Settings", "IconClicked", chkIcon.Value
SaveSetting "WeatherGrabber", "Settings", "Connect", chkConnect.Value
SaveSetting "WeatherGrabber", "Settings", "UseProxy", chkProxy.Value
SaveSetting "WeatherGrabber", "Settings", "Proxy", Proxy
SaveSetting "WeatherGrabber", "Settings", "ProxyPort", ProxyPort
SaveSetting "WeatherGrabber", "Settings", "ProxyUser", ProxyUser
SaveSetting "WeatherGrabber", "Settings", "ProxyPassword", ProxyPassword
SaveSetting "WeatherGrabber", "Settings", "AlertWindow", chkAlert.Value
SaveSetting "WeatherGrabber", "Settings", "AlertSound", fileSound.List(fileSound.ListIndex)

RegZip = txtZipCode.Text
RegRefresh = txtRefresh.Text
RegBackground = lstBackgrounds.Text
RegLoad = chkLoad.Value
RegEverRefresh = chkRefresh.Value
RegInternet = chkInternet.Value
RegNoCache = chkNoCache.Value
RegFontColor = txtFontColor.Text
RegLineColor = txtLineColor.Text
RegIcon = chkIcon.Value
RegConnect = chkConnect.Value
RegUseProxy = chkProxy.Value
RegProxy = Proxy
RegProxyPort = ProxyPort
RegProxyUser = ProxyUser
RegProxyPassword = ProxyPassword
RegAlertWindow = chkAlert.Value
RegSound = fileSound.List(fileSound.ListIndex)

Unload Me
End Sub

Private Sub cmdDefaults_Click()
vscrollRefresh.Value = 30
chkNoCache.Value = 0
chkLoad.Value = 1
chkRefresh.Value = 1
chkInternet.Value = 1
chkIcon.Value = 0
chkConnect.Value = 0
txtFontColor.Text = "0"
txtLineColor.Text = "0"
chkProxy.Value = 0
Proxy = ""
ProxyUser = ""
ProxyPassword = ""
chkAlert.Value = 1

For d = 0 To lstBackgrounds.ListCount - 1
If lstBackgrounds.List(d) = "Radiating Blue" Then
lstBackgrounds.ListIndex = d
Exit For
End If
Next d

For d = 0 To fileSound.ListCount - 1
If fileSound.List(d) = "Alert.wav" Then
fileSound.ListIndex = d
Exit For
End If
Next d

End Sub

Private Sub cmdUnits_Click()
If MsgBox("A new window containing a Weather.com page will open. Scroll down to the bottom of the page and click 'English Units' or 'Metric Units'. Then close the window.", vbInformation + vbOKCancel) = vbCancel Then Exit Sub
RegZip = ""
Shell "explorer.exe " & """" & "http://www.weather.com/weather/local/" & ZipCode & """", vbMaximizedFocus
End Sub

Private Sub Form_Load()
fileBackgrounds.Path = App.Path & "\Backgrounds"
fileSound.Path = App.Path & "\Sounds"
fileImages.Path = App.Path
fileImages.Pattern = "31_*.gif;52_*.gif"
fileImages.Refresh

For d = 0 To fileSound.ListCount - 1
If fileSound.List(d) = RegSound Then
fileSound.ListIndex = d
Exit For
End If
Next d

For c = 0 To fileBackgrounds.ListCount - 1
lstBackgrounds.AddItem Mid(fileBackgrounds.List(c), 1, Len(fileBackgrounds.List(c)) - 4)
Next c

For d = 0 To lstBackgrounds.ListCount - 1
If lstBackgrounds.List(d) = RegBackground Then
lstBackgrounds.ListIndex = d
Exit For
End If
Next d

If lstBackgrounds.ListCount = 1 Then
lblBackgrounds.Caption = "1 Background Available"
Else
lblBackgrounds.Caption = lstBackgrounds.ListCount & " Backgrounds Available"
End If

txtZipCode.Text = ZipCode
vscrollRefresh.Value = RegRefresh

If RegLoad = "1" Then
chkLoad.Value = 1
Else
chkLoad.Value = 0
End If

If RegEverRefresh = "1" Then
chkRefresh.Value = 1
Else
chkRefresh.Value = 0
End If

If RegInternet = "1" Then
chkInternet.Value = 1
Else
chkInternet.Value = 0
End If

If RegNoCache = "1" Then
chkNoCache.Value = 1
Else
chkNoCache.Value = 0
End If

If RegIcon = "1" Then
chkIcon.Value = 1
Else
chkIcon.Value = 0
End If

If RegConnect = "1" Then
chkConnect.Value = 1
Else
chkConnect.Value = 0
End If

txtFontColor.Text = RegFontColor
txtLineColor.Text = RegLineColor

If RegUseProxy = "1" Then
chkProxy.Value = 1
Else
chkProxy.Value = 0
End If

If RegAlertWindow = "1" Then
chkAlert.Value = 1
Else
chkAlert.Value = 0
End If

Proxy = RegProxy
ProxyUser = RegProxyUser
ProxyPassword = RegProxyPassword
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmPreferences = Nothing
End Sub

Private Sub picColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If optFontColor.Value = True Then
txtFontColor.Text = picColors.Point(CurX, CurY)
ElseIf optLineColor.Value = True Then
txtLineColor.Text = picColors.Point(CurX, CurY)
End If
End Sub

Private Sub picColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = X
CurY = Y
End Sub

Private Sub tmrRefresh_Timer()
txtRefresh.Text = vscrollRefresh.Value
End Sub
