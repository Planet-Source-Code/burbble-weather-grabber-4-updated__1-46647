VERSION 5.00
Begin VB.Form frmAlertWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Alert!"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1500
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
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSound 
      Interval        =   1
      Left            =   120
      Top             =   480
   End
   Begin VB.Timer tmrMain 
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
   Begin VB.OLE oleSound 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblAlert 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A new weather alert has been issued in your area. Click here to open the Alerts window."
      Height          =   1215
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAlertWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmAlertWindow = Nothing
End Sub

Private Sub lblAlert_Click()
frmAlerts.Show
Unload frmAlertWindow
End Sub

Private Sub tmrMain_Timer()
Unload frmAlertWindow
tmrMain.Enabled = False
End Sub

Private Sub tmrSound_Timer()
If AlertPlaySound = False Then
tmrSound.Enabled = False
Exit Sub
End If
'oleSound.CreateEmbed App.Path & "\Sounds\Alert.wav"
'oleSound.DoVerb
PlayWav App.Path & "\Sounds\" & RegSound
AlertPlaySound = False
tmrSound.Enabled = False
End Sub
