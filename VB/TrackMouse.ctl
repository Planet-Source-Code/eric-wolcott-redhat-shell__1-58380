VERSION 5.00
Begin VB.UserControl TrackMouse 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image Image1 
      Height          =   420
      Left            =   75
      Picture         =   "TrackMouse.ctx":0000
      Top             =   90
      Width           =   420
   End
End
Attribute VB_Name = "TrackMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents tmLink As CTrackMouse
Attribute tmLink.VB_VarHelpID = -1
Public Event MouseOver()
Public Event MouseOut()
Public Event MouseLeftDown()
Public Event MouseLeftUp()

Public CurrentState As Integer
'7 = Watching
'8 = Closed

Private Sub tmLink_MouseOut()
    RaiseEvent MouseOut
End Sub

Private Sub tmLink_MouseOver()
    RaiseEvent MouseOver
End Sub

Private Sub tmLink_MouseLeftDown()
    RaiseEvent MouseLeftDown
End Sub

Private Sub tmLink_MouseLeftUp()
    RaiseEvent MouseLeftUp
End Sub

Public Function Watch(obj As Object)
Set tmLink = New CTrackMouse
Set tmLink.TrackObject = obj
CurrentState = 7
End Function

Public Function CloseWatch()
If CurrentState = 7 Then
Set tmLink.TrackObject = Nothing
CurrentState = 8
End If
End Function

Private Sub UserControl_Resize()
Image1.Top = 0
Image1.Left = 0
UserControl.Height = Image1.Height
UserControl.Width = Image1.Width
End Sub
