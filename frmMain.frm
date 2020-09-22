VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weather Grabber"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdShutDown 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shut Down"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlerts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alerts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Timer tmrCheck 
      Interval        =   1
      Left            =   5880
      Top             =   1680
   End
   Begin VB.Timer tmrLoad 
      Interval        =   1
      Left            =   5880
      Top             =   1200
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   1000
      Left            =   5880
      Top             =   720
   End
   Begin VB.CommandButton cmdForecast 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forecast"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreferences 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preferences"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox picCur1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   870
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   240
      Width           =   780
   End
   Begin VB.FileListBox fileImages 
      Height          =   315
      Left            =   4080
      Pattern         =   "*.gif"
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Refresh Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Line Line14 
      X1              =   160
      X2              =   440
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Label lblWeatherDotCom 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to visit Weather Grabber's source of information, Weather.com."
      Height          =   495
      Left            =   2520
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Line Line13 
      X1              =   160
      X2              =   440
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   4485
      Width           =   3975
   End
   Begin VB.Image imgRadar 
      Height          =   2655
      Left            =   2520
      MouseIcon       =   "frmMain.frx":0614
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3975
   End
   Begin VB.Line Line12 
      X1              =   440
      X2              =   440
      Y1              =   8
      Y2              =   336
   End
   Begin VB.Line Line11 
      X1              =   160
      X2              =   440
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line10 
      X1              =   8
      X2              =   160
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line Line9 
      X1              =   8
      X2              =   160
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line8 
      X1              =   8
      X2              =   160
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Line Line7 
      X1              =   8
      X2              =   160
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line Line6 
      X1              =   8
      X2              =   160
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Line Line5 
      X1              =   8
      X2              =   160
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Line Line4 
      X1              =   8
      X2              =   440
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line3 
      X1              =   8
      X2              =   440
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Line Line2 
      X1              =   160
      X2              =   160
      Y1              =   8
      Y2              =   336
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   8
      Y1              =   8
      Y2              =   336
   End
   Begin VB.Label lblPressure 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblVisibility 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblWind 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblDew 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblUV 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lblHumidity 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblFeelsLike 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblCur1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpData1 As String
Dim tmpData2 As String
Dim tmpData3 As String
Dim tmpData4 As String
Dim tmpData5 As String
Dim tmpData6 As String
Dim tmpData7 As String
Dim tmpData8 As String
Dim tmpData9 As String
Dim tmpData10 As String

Dim tmpInt1 As Long
Dim tmpInt2 As Long
Dim tmpInt3 As Long
Dim tmpInt4 As Long
Dim tmpInt5 As Long
Dim tmpInt6 As Long
Dim tmpInt7 As Long
Dim tmpInt8 As Long
Dim tmpInt9 As Long
Dim tmpInt10 As Long

Dim ExitNow As Boolean

Dim CurMinute As String
Dim CurSecond As String

Dim LastEnd As Long
Dim LastCaption As String

Private Sub ResetVariables()
tmpData1 = ""
tmpData2 = ""
tmpData3 = ""
tmpData4 = ""
tmpData5 = ""
tmpData6 = ""
tmpData7 = ""
tmpData8 = ""
tmpData9 = ""
tmpData10 = ""
tmpInt1 = 0
tmpInt2 = 0
tmpInt3 = 0
tmpInt4 = 0
tmpInt5 = 0
tmpInt6 = 0
tmpInt7 = 0
tmpInt8 = 0
tmpInt9 = 0
tmpInt10 = 0
End Sub

Private Function CheckImageCached(ImageFileName As String)
fileImages.Path = App.Path
fileImages.Refresh
For a = 0 To fileImages.ListCount - 1
If fileImages.List(a) = ImageFileName Then
CheckImageCached = True
Exit Function
End If
Next a
CheckImageCached = False
Exit Function
End Function

Private Sub CacheImage(URL2 As String)
Dim URL As String
Dim b() As Byte
Dim Filename As String
Dim Size As String

URL = URL2
tmpInt1 = InStr(1, URL, "cons/")
If tmpInt1 = 0 Then tmpInt1 = InStr(1, URL, "rape/")
tmpInt2 = InStr(tmpInt1 + 5, URL, "/")
tmpData1 = Mid(URL, tmpInt1 + 5, tmpInt2 - (tmpInt1 + 5))
Size = tmpData1

tmpInt1 = InStr(tmpInt2 + 1, URL, ".gif")
tmpData2 = Mid(URL, tmpInt2 + 1, tmpInt1 - (tmpInt2 + 1))
Filename = tmpData2

b() = frmControls.inetData.OpenURL(URL, icByteArray)
Open App.Path & "\" & Size & "_" & Filename & ".gif" For Binary Access Write As #1
Put #1, , b()
Close #1
frmControls.inetData.Cancel
Unload frmControls
Set frmControls = Nothing
End Sub

Private Sub GrabData(ZipCode As String)
On Error GoTo ErrHand
Loop1:
DoEvents
CurrentData = frmControls.inetData.OpenURL("http://www.weather.com/weather/local/" & ZipCode)
If CurrentData = "" Then GoTo Loop1
frmControls.inetData.Cancel
Unload frmControls
Set frmControls = Nothing

Exit Sub
ErrHand:
Err.Clear
GoTo Loop1
End Sub

Private Sub GrabText(Find1 As String, Find2 As String)
tmpInt1 = InStr(LastEnd, CurrentData, Find1)

If tmpInt1 = 0 Then
MsgBox "The Weather data could not be read.", vbCritical
tmpInt1 = 1
End If

tmpInt2 = InStr(tmpInt1 + Len(Find1), CurrentData, Find2)
tmpData1 = Mid(CurrentData, tmpInt1 + Len(Find1), tmpInt2 - (tmpInt1 + Len(Find1)))
LastEnd = (tmpInt1 + Len(Find1)) + (tmpInt2 - (tmpInt1 + Len(Find1))) + 1
End Sub

Private Sub GrabMainIcon()
GrabText "<TD CLASS=obsInfo1 VALIGN=TOP ALIGN=CENTER WIDTH=50%>", "<IMG"
GrabText "SRC=", " WIDTH="

tmpInt1 = InStr(1, tmpData1, "/52/")
tmpInt2 = InStr(tmpInt1, tmpData1, ".gif")
tmpData2 = Mid(tmpData1, tmpInt1 + 4, tmpInt2 - (tmpInt1 + 4))
tmpData3 = "52_" & tmpData2 & ".gif"
If CheckImageCached(tmpData3) = False Then CacheImage tmpData1
picCur1.Picture = LoadPicture(App.Path & "\" & tmpData3)
If RegNoCache = "1" Then Kill App.Path & "\" & tmpData3
ResetVariables
End Sub

Private Sub GrabCity()
GrabText "Forecast for ", "</B>"
frmMain.Caption = "Weather Grabber - " & tmpData1
City = tmpData1
SetIconToolTip frmMain.Caption
ResetVariables
End Sub

Private Sub GrabMainWeather()
GrabText "<TD VALIGN=TOP ALIGN=CENTER CLASS=obsInfo2><B CLASS=obsTextA>", "</B></TD>"
lblCur1.Caption = tmpData1
ResetVariables
End Sub

Private Sub GrabTemperature()
GrabText "<B CLASS=obsTempTextA>", "</B>"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblTemp.Caption = tmpData1
tmpData1 = Replace(tmpData1, "°", "")
tmpData1 = Replace(tmpData1, "F", "")
tmpData1 = Replace(tmpData1, "C", "")
tmpData1 = Replace(tmpData1, " ", "")

If tmpData1 = "N/A" Then
SysTrayIcon.hIcon = frmMain.Icon
Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
ResetVariables
Exit Sub
End If

If Int(tmpData1) <= 110 And Int(tmpData1) >= -20 Then
SetIcon App.Path & "\Icons\Icon " & tmpData1 & ".ico"
Else
SysTrayIcon.hIcon = frmMain.Icon
Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
End If

ResetVariables
End Sub

Private Sub GrabFeelsLike()
GrabText "<TD VALIGN=TOP ALIGN=CENTER CLASS=obsInfo2> <B CLASS=obsTextA>Feels Like<BR>", "</B></TD>"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblFeelsLike.Caption = "Feels Like " & tmpData1
ResetVariables
End Sub

Private Sub GrabHumidity()
GrabText "<TD ALIGN=LEFT VALIGN=TOP CLASS=obsInfo2>", "</TD>"
tmpData1 = Replace(tmpData1, "&nbsp;", " ")
lblHumidity.Caption = "Humidity: " & tmpData1
ResetVariables
End Sub

Private Sub GrabUVIndex()
GrabText "<TD ALIGN=LEFT VALIGN=TOP CLASS=obsInfo2>", "</TD>"
tmpData1 = Replace(tmpData1, "&nbsp;", " ")
lblUV.Caption = "UV Index: " & tmpData1
ResetVariables
End Sub

Private Sub GrabDewPoint()
GrabText "<TD ALIGN=LEFT VALIGN=TOP CLASS=obsInfo2>", "</TD>"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblDew.Caption = "Dew Point: " & tmpData1
End Sub

Private Sub GrabVisibility()
GrabText "<TD ALIGN=LEFT VALIGN=TOP CLASS=obsInfo2>", "</TD>"
tmpData1 = Replace(tmpData1, "&nbsp;", " ")
lblVisibility.Caption = "Visibility: " & tmpData1
End Sub

Private Sub GrabPressure()
GrabText "<TD ALIGN=LEFT VALIGN=TOP CLASS=obsInfo2>", "</TD>"
tmpData1 = Replace(tmpData1, "&nbsp;", " ")
lblPressure.Caption = "Pressure: " & tmpData1
End Sub

Private Sub GrabWind()
GrabText "<TD ALIGN=LEFT VALIGN=TOP CLASS=obsInfo2>", "</TD>"
tmpData1 = Replace(tmpData1, "&nbsp;", " ")
lblWind.Caption = "Wind: " & tmpData1
End Sub

Private Sub GrabCaption()
GrabText ">As reported at</A>", "</TD>"
tmpData1 = Replace(tmpData1, "&nbsp;", " ")
tmpData1 = Replace(tmpData1, "</A>", "")
lblCaption.Caption = "As reported at " & tmpData1
End Sub

Private Sub GrabRadarText(Find1 As String, Find2 As String)
tmpInt1 = InStr(1, RadarData, Find1)

If tmpInt1 = 0 Then
MsgBox "The Radar data could not be read.", vbCritical
tmpInt1 = 1
End If

tmpInt2 = InStr(tmpInt1, RadarData, Find2)
tmpData1 = Mid(RadarData, tmpInt1 + Len(Find1), tmpInt2 - (tmpInt1 + Len(Find1)))
End Sub

Private Sub GrabRadarPicture(URL As String)
Dim b() As Byte
b() = frmControls.inetRadar.OpenURL(URL, icByteArray)
Open App.Path & "\RadarLarge.gif" For Binary Access Write As #1
Put #1, , b()
Close #1
frmControls.inetRadar.Cancel
Unload frmControls
Set frmControls = Nothing
imgRadar.Picture = LoadPicture(App.Path & "\RadarLarge.gif")
Kill App.Path & "\RadarLarge.gif"
End Sub

Private Sub GrabRadar()
On Error GoTo ErrHand

Dim LoopNum As Integer

If LastCaption = lblCaption.Caption Then
imgRadar.Picture = LoadPicture(App.Path & "\TempRadar.jpg")
Kill App.Path & "\TempRadar.jpg"
GoTo NoRadar
End If

frmMain.Caption = "Weather Grabber - " & City & " - Radar Loading... 0%"

Loop3:
LoopNum = 3
DoEvents
RadarData = frmControls.inetRadar.OpenURL("http://www.weather.com/weather/map/" & ZipCode) '& "?from=LAPmaps")
If RadarData = "" Then GoTo Loop3
GrabRadarText "if (isMinNS4) var mapNURL = " & """", """" & ";"

frmMain.Caption = "Weather Grabber - " & City & " - Radar Loading... 33%"

Loop4:
LoopNum = 4
DoEvents
RadarData = frmControls.inetRadar.OpenURL("http://www.weather.com" & tmpData1)
If RadarData = "" Then GoTo Loop4
GrabRadarText "<IMG NAME=" & """" & "mapImg" & """" & " SRC=" & """", """" & " WIDTH=600 HEIGHT=405 BORDER=0"

frmMain.Caption = "Weather Grabber - " & City & " - Radar Loading... 66%"

GrabRadarPicture tmpData1

frmMain.Caption = "Weather Grabber - " & City

ResetVariables

LoopNum = 0

LastCaption = lblCaption.Caption
Kill App.Path & "\TempRadar.jpg"

NoRadar:
Exit Sub
ErrHand:
Err.Clear
If LoopNum = 3 Then
GoTo Loop3
ElseIf LoopNum = 4 Then
GoTo Loop4
Else
Exit Sub
End If
End Sub

Private Sub GrabWeather()
GrabTemperature
GrabMainWeather
GrabFeelsLike
GrabUVIndex
GrabDewPoint
GrabHumidity
GrabVisibility
GrabPressure
GrabWind
GrabCaption
LastEnd = 1
GrabMainIcon
GrabAlerts
End Sub

Private Sub cmdAlerts_Click()
If frmControls.inetData.StillExecuting = True Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub
frmAlerts.Show vbModal
End Sub

Private Sub cmdForecast_Click()
If frmControls.inetData.StillExecuting = True Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub
frmForecast.Show vbModal
End Sub

Private Sub cmdPreferences_Click()
If frmControls.inetData.StillExecuting = True Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub
frmPreferences.Show vbModal

If JustCanceled = True Then Exit Sub
SetupForm

Unload frmIP
If frmIP.wsIP.LocalIP = "0.0.0.0" Or frmIP.wsIP.LocalIP = "127.0.0.1" Then Exit Sub

If ChangedZip = False Then Exit Sub
tmrLoad.Enabled = True
End Sub

Private Sub RefreshNow(Connect As Boolean, Radar As Boolean)
'On Error GoTo ErrHand
Dim tmpCaption As String
Dim tmpNothing As String

If frmControls.inetData.StillExecuting = True Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub
LastEnd = 1
If Connect = True Then GoTo Continue3

If RegInternet = "1" Then
Unload frmIP
If frmIP.wsIP.LocalIP = "0.0.0.0" Or frmIP.wsIP.LocalIP = "127.0.0.1" Then Exit Sub
End If

Continue3:
If frmIP.wsIP.LocalIP = "0.0.0.0" Or frmIP.wsIP.LocalIP = "127.0.0.1" Then
tmpNothing = frmControls.inetData.OpenURL("http://www.weather.com")
Do
DoEvents
Unload frmIP
If frmIP.wsIP.LocalIP = "0.0.0.0" Or frmIP.wsIP.LocalIP = "127.0.0.1" Then
Else
Exit Do
End If
Loop
tmpNothing = ""
End If
CurSecond = 0
CurMinute = 0
tmpCaption = frmMain.Caption
frmMain.Caption = frmMain.Caption & " - Loading..."

If Radar = True Then
If imgRadar.Picture = LoadPicture() Then GoTo ContinueLoading
SavePicture imgRadar.Picture, App.Path & "\TempRadar.jpg"
imgRadar.Picture = LoadPicture()
End If

ContinueLoading:

GrabData ZipCode
If InStr(1, CurrentData, "Sorry, the page you requested was not found on weather.com.") Then
MsgBox "Invalid Zip Code!", vbCritical
frmMain.Caption = tmpCaption
tmpCaption = ""
Exit Sub
End If
GrabCity
GrabWeather
If Radar = True Then GrabRadar
tmpCaption = ""

Exit Sub
ErrHand:
Err.Clear
GoTo Continue3
End Sub

Private Sub cmdRefresh_Click()
Unload frmIP
If frmIP.wsIP.LocalIP = "0.0.0.0" Or frmIP.wsIP.LocalIP = "127.0.0.1" Then Exit Sub
RefreshNow False, True
End Sub

Private Sub SetupForm()
Dim Input1 As String

CurSecond = 0
CurMinute = 0

RegFirst = GetSetting("WeatherGrabber", "Settings", "FirstRun")
If RegFirst = "" Then
Input1 = InputBox("This is your first time running Weather Grabber. Please enter your zip code below.", , "Zip Code")
If Input1 = "" Then
ExitNow = True
Unload frmMain
End If
SaveSetting "WeatherGrabber", "Settings", "ZipCode", Input1
SaveSetting "WeatherGrabber", "Settings", "FirstRun", "1"
SaveSetting "WeatherGrabber", "Settings", "Background", "Radiating Blue"
SaveSetting "WeatherGrabber", "Settings", "Refresh", "30"
SaveSetting "WeatherGrabber", "Settings", "LoadAtStart", "1"
SaveSetting "WeatherGrabber", "Settings", "EverRefresh", "1"
SaveSetting "WeatherGrabber", "Settings", "CheckInternet", "1"
SaveSetting "WeatherGrabber", "Settings", "NoCache", "0"
SaveSetting "WeatherGrabber", "Settings", "FontColor", vbBlack
SaveSetting "WeatherGrabber", "Settings", "LineColor", vbBlack
SaveSetting "WeatherGrabber", "Settings", "IconClicked", "0"
SaveSetting "WeatherGrabber", "Settings", "Connect", "0"
SaveSetting "WeatherGrabber", "Settings", "UseProxy", "0"
SaveSetting "WeatherGrabber", "Settings", "Proxy", ""
SaveSetting "WeatherGrabber", "Settings", "ProxyPort", "80"
SaveSetting "WeatherGrabber", "Settings", "ProxyUser", ""
SaveSetting "WeatherGrabber", "Settings", "ProxyPassword", ""
SaveSetting "WeatherGrabber", "Settings", "AlertWindow", "1"
SaveSetting "WeatherGrabber", "Settings", "AlertSound", "Alert.wav"
End If

RegZip = GetSetting("WeatherGrabber", "Settings", "ZipCode")
ZipCode = RegZip

RegBackground = GetSetting("WeatherGrabber", "Settings", "Background")
frmMain.Picture = LoadPicture(App.Path & "\Backgrounds\" & RegBackground & ".gif")

RegRefresh = GetSetting("WeatherGrabber", "Settings", "Refresh")

RegEverRefresh = GetSetting("WeatherGrabber", "Settings", "EverRefresh")

RegLoad = GetSetting("WeatherGrabber", "Settings", "LoadAtStart")

RegInternet = GetSetting("WeatherGrabber", "Settings", "CheckInternet")

RegNoCache = GetSetting("WeatherGrabber", "Settings", "NoCache")

RegFontColor = GetSetting("WeatherGrabber", "Settings", "FontColor")
lblCaption.ForeColor = RegFontColor
lblCur1.ForeColor = RegFontColor
lblDew.ForeColor = RegFontColor
lblFeelsLike.ForeColor = RegFontColor
lblHumidity.ForeColor = RegFontColor
lblPressure.ForeColor = RegFontColor
lblTemp.ForeColor = RegFontColor
lblUV.ForeColor = RegFontColor
lblVisibility.ForeColor = RegFontColor
lblWeatherDotCom.ForeColor = RegFontColor
lblWind.ForeColor = RegFontColor

RegLineColor = GetSetting("WeatherGrabber", "Settings", "LineColor")
Line1.BorderColor = RegLineColor
Line2.BorderColor = RegLineColor
Line3.BorderColor = RegLineColor
Line4.BorderColor = RegLineColor
Line5.BorderColor = RegLineColor
Line6.BorderColor = RegLineColor
Line7.BorderColor = RegLineColor
Line8.BorderColor = RegLineColor
Line9.BorderColor = RegLineColor
Line10.BorderColor = RegLineColor
Line11.BorderColor = RegLineColor
Line12.BorderColor = RegLineColor
Line13.BorderColor = RegLineColor
Line14.BorderColor = RegLineColor

RegIcon = GetSetting("WeatherGrabber", "Settings", "IconClicked")
RegConnect = GetSetting("WeatherGrabber", "Settings", "Connect")
RegUseProxy = GetSetting("WeatherGrabber", "Settings", "UseProxy")
RegProxy = GetSetting("WeatherGrabber", "Settings", "Proxy")
RegProxyPort = GetSetting("WeatherGrabber", "Settings", "ProxyPort")
RegProxyUser = GetSetting("WeatherGrabber", "Settings", "ProxyUser")
RegProxyPassword = GetSetting("WeatherGrabber", "Settings", "ProxyPassword")

If RegUseProxy = "1" Then
frmControls.inetData.Proxy = RegProxy
If RegProxyPort <> "" Then frmControls.inetData.RemotePort = RegProxyPort
If RegProxyUser <> "" Then frmControls.inetData.UserName = RegProxyUser
If RegProxyPassword <> "" Then frmControls.inetData.Password = RegProxyPassword

frmControls.inetRadar.Proxy = RegProxy
If RegProxyPort <> "" Then frmControls.inetRadar.RemotePort = RegProxyPort
If RegProxyUser <> "" Then frmControls.inetRadar.UserName = RegProxyUser
If RegProxyPassword <> "" Then frmControls.inetRadar.Password = RegProxyPassword
End If

RegAlertWindow = GetSetting("WeatherGrabber", "Settings", "AlertWindow")
RegSound = GetSetting("WeatherGrabber", "Settings", "AlertSound")

End Sub

Private Sub cmdShutDown_Click()
If frmControls.inetData.StillExecuting = True Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub
If MsgBox("Are you sure you want to exit Weather Grabber?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
ExitNow = True
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 97 Then
MsgBox """" & tmpData1 & """"
KeyAscii = 0
Exit Sub
End If
If KeyAscii = 98 Then
MsgBox CurMinute & ":" & CurSecond & vbNewLine & RegRefresh
KeyAscii = 0
Exit Sub
End If
End Sub

Private Sub Form_Load()
SetupForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ExitNow = True Then
Unload frmIcon
Set frmMain = Nothing
Set frmRadar = Nothing
Set frmPreferences = Nothing
Set frmControls = Nothing
Set frmDetails = Nothing
Set frmAlerts = Nothing
Set frmAlertWindow = Nothing
Set frmIP = Nothing
End
Else
Cancel = 1
frmMain.Visible = False
End If
End Sub

Private Sub imgRadar_Click()
If imgRadar.Picture = LoadPicture() Then Exit Sub
If frmControls.inetRadar.StillExecuting = True Then Exit Sub
If frmControls.inetData.StillExecuting = True Then Exit Sub
frmRadar.Show vbModal
End Sub

Private Sub lblWeatherDotCom_Click()
Shell "explorer.exe " & """" & "http://www.weather.com" & """", vbMaximizedFocus
End Sub

Private Sub tmrCheck_Timer()
If CheckNow = True Then
RefreshNow False, False
CheckNow = False
End If
End Sub

Private Sub tmrLoad_Timer()
If RegConnect = "1" Then
If HideMe = True Then frmMain.Visible = False
RefreshNow True, True
tmrLoad.Enabled = False
Exit Sub
End If

If RegLoad = "1" Then
If HideMe = True Then frmMain.Visible = False
RefreshNow False, True
tmrLoad.Enabled = False
Exit Sub
End If
End Sub

Private Sub tmrRefresh_Timer()
If RegEverRefresh = "1" Then
RefreshRate = RegRefresh
If CurMinute >= RefreshRate Then
cmdRefresh_Click
CurMinute = 0
CurSecond = 0
Exit Sub
End If

If CurSecond = 59 Then
CurMinute = CurMinute + 1
CurSecond = 0
Exit Sub
End If

CurSecond = CurSecond + 1
End If
End Sub
