VERSION 5.00
Begin VB.Form frmRadar 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Radar - Left-click to close - Right-click to animate"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10800
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
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   840
   End
   Begin VB.Image imgFrame 
      Height          =   375
      Index           =   6
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgFrame 
      Height          =   375
      Index           =   5
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgFrame 
      Height          =   375
      Index           =   4
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgFrame 
      Height          =   375
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgFrame 
      Height          =   375
      Index           =   2
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgFrame 
      Height          =   375
      Index           =   1
      Left            =   840
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmRadar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpData1 As String
Dim tmpInt1 As Long
Dim tmpInt2 As Long
Dim Animated As Boolean
Dim CurFrame As Integer
Dim tmpTime As String

'The following function was written by 'egkenny'.

Private Function getTime()
Day_1_1_1970 = Format(CDate("1,1,1970"), "#######.#")
Day_Today = Format(Date, "#######.#")
Days_Elapsed = Day_Today - Day_1_1_1970
Time_Zone = 5
Hrs_Elapsed = Days_Elapsed * 24 + Time_Zone
Sec_Elapsed = Hrs_Elapsed * 3600
Sec_Today = Timer
Msec_Elapsed = (Sec_Elapsed + Sec_Today) * 1000
getTime = Msec_Elapsed
End Function

Private Sub GrabRadarFrame(URL As String, Frame As Integer)
Dim b() As Byte
b() = frmControls.inetRadar.OpenURL(URL, icByteArray)
Open App.Path & "\RadarFrame.jpg" For Binary Access Write As #1
Put #1, , b()
Close #1
imgFrame(Frame).Picture = LoadPicture(App.Path & "\RadarFrame.jpg")
Kill App.Path & "\RadarFrame.jpg"
End Sub

Private Sub GrabAnimateText(Find1 As String, Find2 As String)
tmpInt1 = InStr(1, AnimateData, Find1)

If tmpInt1 = 0 Then
MsgBox "The Animation data could not be read.", vbCritical
tmpInt1 = 1
End If

tmpInt2 = InStr(tmpInt1, AnimateData, Find2)
tmpData1 = Mid(AnimateData, tmpInt1 + Len(Find1), tmpInt2 - (tmpInt1 + Len(Find1)))
End Sub

Private Sub AnimateRadar()
Dim tmpString As String
Dim tmpInteger As Integer

frmRadar.Caption = "Animating Radar - 12%"
AnimateData = frmControls.inetRadar.OpenURL("http://www.weather.com/weather/map/" & ZipCode & "?name=index_large_animated&day=1")
GrabAnimateText "if (isMinNS4) var mapNURL = " & """", """" & ";"

frmRadar.Caption = "Animating Radar - 24%"
AnimateData = frmControls.inetRadar.OpenURL("http://www.weather.com" & tmpData1)
GrabAnimateText "var thisMap = ['", "'];"

tmpTime = getTime
If InStr(1, tmpTime, ".") <> 0 Then tmpTime = Left(tmpTime, Len(tmpTime) - (Len(tmpTime) - InStr(1, tmpTime, ".")) - 1)

For f = 1 To 6
tmpString = f
tmpInteger = f
GrabRadarFrame "http://image.weather.com" & tmpData1 & tmpString & "L.jpg" & "?" & tmpTime, tmpInteger
'MsgBox "http://image.weather.com" & tmpData1 & tmpString & "L.jpg" & "?" & tmpTime
If f = 1 Then frmRadar.Caption = "Animating Radar - 36%"
If f = 2 Then frmRadar.Caption = "Animating Radar - 48%"
If f = 3 Then frmRadar.Caption = "Animating Radar - 60%"
If f = 4 Then frmRadar.Caption = "Animating Radar - 72%"
If f = 5 Then frmRadar.Caption = "Animating Radar - 84%"
If f = 6 Then frmRadar.Caption = "Animating Radar - 96%"
DoEvents
Next f

frmRadar.Caption = "Radar - Click to Close"

tmpString = ""
tmpData1 = ""
tmpInt1 = 0
tmpInt2 = 0
End Sub

Private Sub Form_Load()
frmRadar.Picture = frmMain.imgRadar.Picture
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmControls.inetRadar.StillExecuting = True Then Exit Sub

If Animated = True Or Button = vbLeftButton Then
Unload Me
Exit Sub
End If
If Button = vbRightButton Then
Unload frmIP
If frmIP.wsIP.LocalIP = "0.0.0.0" Or frmIP.wsIP.LocalIP = "127.0.0.1" Then Exit Sub
Animated = True
frmRadar.Picture = LoadPicture()
CurFrame = 1
tmrAnimate.Enabled = True

AnimateRadar
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If frmControls.inetRadar.StillExecuting = True Then
Cancel = 1
Exit Sub
End If
frmControls.inetRadar.Cancel
Unload frmControls
Set frmControls = Nothing
For a = 1 To 6
imgFrame(a).Picture = LoadPicture()
Next a
Set frmRadar = Nothing
End Sub

Private Sub tmrAnimate_Timer()
If CurFrame = 7 Then
If imgFrame(6).Picture = LoadPicture() Then GoTo Continue2
frmRadar.PaintPicture imgFrame(6).Picture, 0, 0
frmRadar.ForeColor = vbWhite
frmRadar.FontSize = 10
frmRadar.CurrentX = 6
frmRadar.CurrentY = 387
frmRadar.Print "6"
Continue2:
CurFrame = 1
DoEvents
Exit Sub
End If
If imgFrame(CurFrame).Picture = LoadPicture() Then GoTo Continue1
frmRadar.PaintPicture imgFrame(CurFrame).Picture, 0, 0
frmRadar.ForeColor = vbWhite
frmRadar.FontSize = 10
frmRadar.CurrentX = 2
frmRadar.CurrentY = 387
frmRadar.Print CurFrame
Continue1:
CurFrame = CurFrame + 1
DoEvents
End Sub
