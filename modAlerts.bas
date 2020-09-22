Attribute VB_Name = "modAlerts"
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Dim tmpData1 As String
Dim tmpInt1 As Long
Dim tmpInt2 As Long
Dim DataStart1() As Long
Dim DataStart2() As Long

Public Function PlayWav(Path As String)
sndPlaySound Path, &H1
End Function

Private Sub GrabText(Start As Long, Find1 As String, Find2 As String)
tmpInt1 = InStr(Start, AlertData, Find1)

If tmpInt1 = 0 Then
tmpData1 = ""
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

Public Sub GrabAlerts()
AlertData = CurrentData
GrabText 1, "var marqueecontents=", "var rawAlertMessageLength = "
If tmpData1 = "" Then Exit Sub
AlertData = tmpData1
AlertPreviousCount = AlertCount
AlertCount = 0
For o = 1 To Len(AlertData)
If Mid(AlertData, o, 1) = "[" Then AlertCount = AlertCount + 1
Next o

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
Next q

ResetVariables
If RegAlertWindow = "0" Then Exit Sub

If AlertCount > AlertPreviousCount Then
AlertPlaySound = True
frmAlertWindow.Show
Exit Sub
End If

End Sub

