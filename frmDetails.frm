VERSION 5.00
Begin VB.Form frmDetails 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detailed Forecast"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11835
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
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   789
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConvert2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Convert"
      Height          =   255
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtF2 
      Height          =   255
      Left            =   8640
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   21
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtC2 
      Height          =   255
      Left            =   7680
      MaxLength       =   3
      TabIndex        =   18
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton cmdConvert1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Convert"
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtC1 
      Height          =   255
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   15
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtF1 
      Height          =   255
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   12
      Top             =   4680
      Width           =   495
   End
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.Line Line7 
      X1              =   394
      X2              =   394
      Y1              =   0
      Y2              =   293
   End
   Begin VB.Label lblF2 
      BackStyle       =   0  'Transparent
      Caption         =   "°F:"
      Height          =   495
      Left            =   8640
      TabIndex        =   22
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblTo2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   255
      Left            =   8280
      TabIndex        =   20
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblC2 
      BackStyle       =   0  'Transparent
      Caption         =   "°C:"
      Height          =   495
      Left            =   7680
      TabIndex        =   19
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblC1 
      BackStyle       =   0  'Transparent
      Caption         =   "°C:"
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblTo1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblF1 
      BackStyle       =   0  'Transparent
      Caption         =   "°F:"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   4440
      Width           =   495
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   788
      Y1              =   292
      Y2              =   292
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   788
      Y1              =   244
      Y2              =   244
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   788
      Y1              =   196
      Y2              =   196
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   788
      Y1              =   148
      Y2              =   148
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   788
      Y1              =   100
      Y2              =   100
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   788
      Y1              =   52
      Y2              =   52
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   11
      Left            =   5955
      TabIndex        =   11
      Top             =   3720
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   10
      Left            =   5955
      TabIndex        =   10
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   9
      Left            =   5955
      TabIndex        =   9
      Top             =   2280
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   8
      Left            =   5955
      TabIndex        =   8
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   7
      Left            =   5955
      TabIndex        =   7
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   6
      Left            =   5955
      TabIndex        =   6
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   5
      Left            =   15
      TabIndex        =   5
      Top             =   3720
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   4
      Left            =   15
      TabIndex        =   4
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   3
      Left            =   15
      TabIndex        =   3
      Top             =   2280
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   2
      Left            =   15
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   15
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpData1 As String
Dim tmpInt1 As Long
Dim tmpInt2 As Long
Dim DataStart(11) As Long

Private Sub GrabText(Start As Long, Find1 As String, Find2 As String)
tmpInt1 = InStr(Start, ForecastData, Find1)

If tmpInt1 = 0 Then
MsgBox "The Forecast data could not be read.", vbCritical
tmpInt1 = 1
End If

tmpInt2 = InStr(tmpInt1, ForecastData, Find2)
tmpData1 = Mid(ForecastData, tmpInt1 + Len(Find1), tmpInt2 - (tmpInt1 + Len(Find1)))
End Sub

Private Sub ResetVariables()
tmpData1 = ""
tmpInt1 = 0
tmpInt2 = 0
End Sub

Private Sub cmdConvert1_Click()
On Error GoTo ErrHand
If txtF1.Text = "" Then Exit Sub
txtC1.Text = (5 / 9) * (Int(txtF1.Text) - 32)
Exit Sub
ErrHand:
MsgBox "Invalid temperature!", vbCritical
txtC1.Text = ""
txtF1.Text = ""
Err.Clear
Exit Sub
End Sub

Private Sub cmdConvert2_Click()
On Error GoTo ErrHand
If txtC2.Text = "" Then Exit Sub
txtF2.Text = (9 / 5) * Int(txtC2.Text) + 32
Exit Sub
ErrHand:
MsgBox "Invalid temperature!", vbCritical
txtC2.Text = ""
txtF2.Text = ""
Err.Clear
Exit Sub
End Sub

Private Sub tmrMain_Timer()
frmDetails.Caption = "Detailed Forecast - Loading..."
Loop2:
DoEvents
ForecastData = frmControls.inetData.OpenURL("http://www.weather.com/weather/narrative/" & ZipCode)
If ForecastData = "" Then GoTo Loop2
GrabText 1, "<!-- begin loop -->", "<!-- end loop -->"
ForecastData = tmpData1

For j = 0 To 11
If j = 0 Then
DataStart(j) = InStr(1, ForecastData, "<!-- insert day/date/time of day --")
Else
DataStart(j) = InStr(DataStart(j - 1) + 35, ForecastData, "<!-- insert day/date/time of day --")
End If
Next j

For k = 0 To 11
If DataStart(k) = 0 Then
lblInfo(k).Caption = ""
GoTo NextK
End If
GrabText DataStart(k), ">", "         </TD>"
tmpData1 = Replace(tmpData1, "</B>&nbsp;", " ")
lblInfo(k).Caption = Replace(tmpData1, Chr(9), "")
NextK:
Next k

frmDetails.Caption = "Detailed Forecast"
tmrMain.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If frmControls.inetData.StillExecuting = True Then
Cancel = 1
Exit Sub
End If
Set frmDetails = Nothing
End Sub
