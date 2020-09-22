VERSION 5.00
Begin VB.Form frmAlerts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alerts"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8280
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
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   3600
      Top             =   600
   End
   Begin VB.ListBox lstAlerts 
      Height          =   1110
      ItemData        =   "frmAlerts.frx":0000
      Left            =   120
      List            =   "frmAlerts.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alerts in your area (double-click on an alert to read the details):"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmAlerts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpData1 As String
Dim tmpInt1 As Long
Dim tmpInt2 As Long
Dim DataStart1() As Long
Dim DataStart2() As Long
Dim UnloadNow As Boolean

Private Sub GrabText(Start As Long, Find1 As String, Find2 As String)
tmpInt1 = InStr(Start, AlertData, Find1)

If tmpInt1 = 0 Then
MsgBox "There are no weather alerts in your area.", vbCritical
tmpData1 = ""
UnloadNow = True
Exit Sub
Else
tmpInt2 = InStr(tmpInt1 + Len(Find1), AlertData, Find2)
tmpData1 = Mid(AlertData, tmpInt1 + Len(Find1), tmpInt2 - (tmpInt1 + Len(Find1)))
End If
End Sub

Private Sub ResetVariables()
tmpData1 = ""
tmpInt1 = 0
tmpInt2 = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmAlerts = Nothing
End Sub

Private Sub lstAlerts_DblClick()
If lstAlerts.ListIndex = -1 Then Exit Sub
Shell "explorer.exe " & """" & AlertURL(lstAlerts.ListIndex + 1) & """", vbMaximizedFocus
End Sub

Private Sub tmrMain_Timer()
If UnloadNow = True Then Exit Sub
If CurrentData = "" Then
MsgBox "You must be connected to the internet to use this feature.", vbCritical
Exit Sub
End If

lstAlerts.Clear
AlertData = CurrentData
UnloadNow = False
GrabText 1, "var marqueecontents=", "var rawAlertMessageLength = "
If UnloadNow = True Then Unload Me
If tmpData1 = "" Then Exit Sub
AlertData = tmpData1
AlertPreviousCount = AlertCount
AlertCount = 0
For l = 1 To Len(AlertData)
If Mid(AlertData, l, 1) = "[" Then AlertCount = AlertCount + 1
Next l

ReDim DataStart1(1 To AlertCount)
ReDim DataStart2(1 To AlertCount)
ReDim AlertURL(1 To AlertCount)
ReDim AlertCaption(1 To AlertCount)

For p = 1 To AlertCount
If p = 1 Then
DataStart1(p) = InStr(1, AlertData, "parent.mapWindowOpen")
DataStart2(p) = InStr(DataStart1(p), AlertData, ">")
Else
DataStart1(p) = InStr(DataStart1(p - 1) + 9, AlertData, "parent.mapWindowOpen")
DataStart2(p) = InStr(DataStart1(p), AlertData, ">")
End If
Next p

For q = 1 To AlertCount
GrabText DataStart1(q), "('", "'"
AlertURL(q) = "http://www.weather.com" & tmpData1
GrabText DataStart2(q) - 1, ">", "</A>"
tmpData1 = Replace(tmpData1, "… [More Details]", "")
tmpData1 = Replace(tmpData1, ".. [More Details]", "")
tmpData1 = Replace(tmpData1, "  ", "")
AlertCaption(q) = tmpData1
lstAlerts.AddItem AlertCaption(q)
Next q

ResetVariables
tmrMain.Enabled = False
End Sub
