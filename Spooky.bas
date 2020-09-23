Attribute VB_Name = "Module1"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Sub LoadAVI(MMCObj As Object, ContObj As Object, strAVI As String)
    MMCObj.hWndDisplay = ContObj.hWnd
    MMCObj.Command = "Close"
    MMCObj.DeviceType = "AVIVideo"
    MMCObj.FileName = strAVI
    MMCObj.Command = "Open"
End Sub


Sub MMCommand(MMCObj As Object, strCmd As String)

    MMCObj.Command = strCmd
    
End Sub

Sub PlayAVIFrom(MMCObj As Object, Optional PlayFrom As Long)

If IsMissing(PlayFrom) Then PlayFrom = 1

MMCObj.From = PlayFrom
MMCObj.Command = "Play"

End Sub

