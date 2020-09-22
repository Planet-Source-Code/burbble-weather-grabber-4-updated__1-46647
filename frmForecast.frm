VERSION 5.00
Begin VB.Form frmForecast 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "10-Day Forecast"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6120
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
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   3120
   End
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   2880
      Top             =   360
   End
   Begin VB.FileListBox fileImages 
      Height          =   300
      Left            =   1680
      Pattern         =   "*.gif"
      TabIndex        =   43
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblDetailed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click here for the Detailed Forcast"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   1635
   End
   Begin VB.Line Line13 
      X1              =   340
      X2              =   340
      Y1              =   32
      Y2              =   440
   End
   Begin VB.Line Line12 
      X1              =   252
      X2              =   252
      Y1              =   32
      Y2              =   440
   End
   Begin VB.Line Line11 
      X1              =   108
      X2              =   108
      Y1              =   0
      Y2              =   440
   End
   Begin VB.Line Line10 
      X1              =   0
      X2              =   408
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line9 
      X1              =   0
      X2              =   408
      Y1              =   396
      Y2              =   396
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   408
      Y1              =   356
      Y2              =   356
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   408
      Y1              =   316
      Y2              =   316
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   408
      Y1              =   276
      Y2              =   276
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   408
      Y1              =   236
      Y2              =   236
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   408
      Y1              =   196
      Y2              =   196
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   408
      Y1              =   156
      Y2              =   156
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   408
      Y1              =   116
      Y2              =   116
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   408
      Y1              =   76
      Y2              =   76
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   5160
      TabIndex        =   42
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   3840
      TabIndex        =   41
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   1680
      TabIndex        =   40
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   39
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   9
      Left            =   1080
      Top             =   6000
      Width           =   465
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   5160
      TabIndex        =   38
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   37
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   1680
      TabIndex        =   36
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   35
      Top             =   5400
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   8
      Left            =   1080
      Top             =   5400
      Width           =   465
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   34
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   33
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   32
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   31
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   7
      Left            =   1080
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   30
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   29
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   28
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   4200
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   6
      Left            =   1080
      Top             =   4200
      Width           =   465
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   26
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   25
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   24
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   5
      Left            =   1080
      Top             =   3600
      Width           =   465
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   21
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   4
      Left            =   1080
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   18
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   16
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   3
      Left            =   1080
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   2
      Left            =   1080
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   1
      Left            =   1080
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Precip. %"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Daytime High / Overnight Low"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Forecast"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblPrecip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblForecast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   0
      Left            =   1080
      Top             =   600
      Width           =   465
   End
End
Attribute VB_Name = "frmForecast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpData1 As String
Dim tmpData2 As String
Dim tmpData3 As String
Dim tmpInt1 As Long
Dim tmpInt2 As Long
Dim LastEnd As Long

Dim DataStart(9) As Long
Dim OnlyLow As Boolean

Dim UnloadNow As Boolean

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
tmpInt1 = InStr(1, URL, "page/")
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

Private Sub GrabText(Start As Long, Find1 As String, Find2 As String)
tmpInt1 = InStr(Start, CurrentData, Find1)

If tmpInt1 = 0 Then
tmpInt1 = InStr(1, CurrentData, Find1)
MsgBox "The Weather data could not be read.", vbCritical
'tmpInt1 = 1
End If

tmpInt2 = InStr(tmpInt1 + Len(Find1), CurrentData, Find2)
tmpData1 = Mid(CurrentData, tmpInt1 + Len(Find1), tmpInt2 - (tmpInt1 + Len(Find1)))
LastEnd = (tmpInt1 + Len(Find1)) + (tmpInt2 - (tmpInt1 + Len(Find1))) + 1
End Sub

Private Sub ResetVariables()
tmpData1 = ""
tmpData2 = ""
tmpData3 = ""
tmpInt1 = 0
tmpInt2 = 0
End Sub

Private Sub Form_Load()
If UnloadNow = True Then Exit Sub
If CurrentData = "" Then
tmrMain.Enabled = False
Exit Sub
End If

CurrentData = Replace(CurrentData, "CLASS=f2a", "")

tmpInt1 = InStr(1, CurrentData, "<!-- begin loop -->")
If tmpInt1 = 0 Then
UnloadNow = True
tmrUnload.Enabled = True
Exit Sub
Else
UnloadNow = False
End If

For g = 0 To 9
If g = 0 Then
DataStart(g) = InStr(tmpInt1, CurrentData, "<TR>")
Else
DataStart(g) = InStr(DataStart(g - 1) + 4, CurrentData, "<TR>")
End If
Next g
tmpInt1 = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If frmControls.inetRadar.StillExecuting = True Then
Cancel = 1
Exit Sub
End If

If frmControls.inetData.StillExecuting = True Then
Cancel = 1
Exit Sub
End If

Set frmForecast = Nothing
End Sub

Private Sub lblDetailed_Click()
If frmControls.inetData.StillExecuting = True Or frmControls.inetRadar.StillExecuting = True Then
MsgBox "Please wait a few moments and try again.", vbCritical
Exit Sub
End If
Unload frmIP
If frmIP.wsIP.LocalIP = "0.0.0.0" Or frmIP.wsIP.LocalIP = "127.0.0.1" Then
MsgBox "You must be connected to the internet to use this feature.", vbCritical
Exit Sub
End If
frmDetails.Show vbModal
End Sub

Private Sub tmrMain_Timer()

'Thanks to Michael Rilling for supplying
'the code to make this work with cities
'other than those in the US.

If UnloadNow = True Then Exit Sub
Dim tStr As String

If Left(ZipCode, 1) >= 0 And Left(ZipCode, 1) <= 9 Then
  tStr = "/weather/detail/"
Else
  tStr = "/outlook/travel/detail/"
End If

GrabText DataStart(0), "<A HREF=" & tStr & ZipCode & ">", "</A><BR>"

lblDay(0).Caption = tmpData1
If tmpData1 = "Tonight" Then
OnlyLow = True
Else
OnlyLow = False
End If

GrabText DataStart(0), "</A><BR> ", "</TD>"
lblDay(0).Caption = lblDay(0).Caption & vbNewLine & tmpData1

GrabText LastEnd, "<TD WIDTH=" & """" & "35%" & """" & ">", "</TD>"
lblForecast(0).Caption = tmpData1

If OnlyLow = True Then
GrabText LastEnd, "<TD WIDTH=" & """" & "25%" & """" & " ALIGN=" & """" & "CENTER" & """" & "><B>", "</B>"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblTemp(0).Caption = tmpData1
ElseIf OnlyLow = False Then
GrabText LastEnd, "<TD WIDTH=" & """" & "25%" & """" & " ALIGN=" & """" & "CENTER" & """" & "><B>", "/"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblTemp(0).Caption = tmpData1
GrabText LastEnd - 1, "/", "</B>"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblTemp(0).Caption = lblTemp(0).Caption & "/" & tmpData1
End If

GrabText LastEnd, "<TD WIDTH=" & """" & "15%" & """" & " ALIGN=" & """" & "CENTER" & """" & ">", "</TD>"
tmpData1 = Replace(tmpData1, " ", "")
lblPrecip(0).Caption = tmpData1

For h = 1 To 9
LastEnd = DataStart(h)
GrabText LastEnd, "?dayNum=" & h & ">", "</A><BR>"
lblDay(h).Caption = tmpData1
If tmpData1 = "Tonight" Then
OnlyLow = True
Else
OnlyLow = False
End If

GrabText DataStart(h), "</A><BR> ", "</TD>"
lblDay(h).Caption = lblDay(h).Caption & vbNewLine & tmpData1

GrabText LastEnd, "WIDTH=31 HEIGHT=31 BORDER=0></T", ">"
GrabText LastEnd, "<TD >", "</TD>"
lblForecast(h).Caption = tmpData1

If OnlyLow = True Then
GrabText LastEnd, "<B>", "</B>"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblTemp(h).Caption = tmpData1
ElseIf OnlyLow = False Then
GrabText LastEnd, "<B>", "/"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblTemp(h).Caption = tmpData1
GrabText LastEnd - 1, "/", "</B>"
tmpData1 = Replace(tmpData1, "&deg;", "°")
lblTemp(h).Caption = lblTemp(h).Caption & "/" & tmpData1
End If

GrabText LastEnd, "<TD  ALIGN=" & """" & "CENTER" & """" & ">", "</TD>"
tmpData1 = Replace(tmpData1, " ", "")
lblPrecip(h).Caption = tmpData1
Next h

Unload frmIP
If frmIP.wsIP.LocalIP = "0.0.0.0" Or frmIP.wsIP.LocalIP = "127.0.0.1" Then
tmrMain.Enabled = False
Exit Sub
End If

For i = 0 To 9
GrabText DataStart(i), "<IMG SRC=" & """", """" & " WIDTH="
tmpInt1 = InStr(1, tmpData1, "/31/")
tmpInt2 = InStr(tmpInt1, tmpData1, ".gif")
tmpData2 = Mid(tmpData1, tmpInt1 + 4, tmpInt2 - (tmpInt1 + 4))
tmpData3 = "31_" & tmpData2 & ".gif"
If CheckImageCached(tmpData3) = False Then CacheImage tmpData1
imgIcon(i).Picture = LoadPicture(App.Path & "\" & tmpData3)
If RegNoCache = "1" Then Kill App.Path & "\" & tmpData3
ResetVariables
Next i

tmrMain.Enabled = False
End Sub

Private Sub tmrUnload_Timer()
MsgBox "The Forecast could not be read.", vbCritical
'Unload frmForecast
tmrUnload.Enabled = False
End Sub
