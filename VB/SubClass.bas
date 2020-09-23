Attribute VB_Name = "SubClass"
Option Explicit

Public colTrackMouse As New Collection


Public Function procTrackMouse(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Dim tmItem As New CTrackMouse
Set tmItem = colTrackMouse.Item("TM" & hWnd)
If Not (tmItem Is Nothing) Then procTrackMouse = tmItem.MessageReceived(wMsg, wParam, lParam)
End Function



