Attribute VB_Name = "modMain"
Sub Main()
If Left(Command, 8) = "-startup" Then
HideMe = True
Load frmIcon
Else
HideMe = False
Load frmIcon
End If
End Sub
