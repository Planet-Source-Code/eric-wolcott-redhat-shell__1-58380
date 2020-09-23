VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrackMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bTracking As Boolean

Private mTrackObject As Object

Private procPrevWndFunc As Long

Private Const WM_MOUSEHOVER = &H2A1&
Private Const WM_MOUSELEAVE = &H2A3&

Private Const WM_MOUSEMOVE = &H200

Private Const WM_LBUTTONDBLCLK As Integer = &H203
Private Const WM_LBUTTONDOWN As Integer = &H201
Private Const WM_LBUTTONUP  As Integer = &H202
Private Const WM_MBUTTONDBLCLK  As Integer = &H209
Private Const WM_MBUTTONDOWN  As Integer = &H207
Private Const WM_MBUTTONUP  As Integer = &H208
Private Const WM_MOUSEACTIVATE  As Integer = &H21
Private Const WM_MOUSEFIRST  As Integer = &H200
Private Const WM_MOUSELAST  As Integer = &H209
Private Const WM_RBUTTONDBLCLK  As Integer = &H206
Private Const WM_RBUTTONDOWN  As Integer = &H204
Private Const WM_RBUTTONUP  As Integer = &H205

Private Const TME_HOVER = &H1&
Private Const TME_LEAVE = &H2&
Private Const TME_QUERY = &H40000000
Private Const TME_CANCEL = &H80000000

Private Const HOVER_DEFAULT = &HFFFFFFFF
Private Const GWL_WNDPROC = -4

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTINFO) As Long

Private Type TRACKMOUSEEVENTINFO
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Public Event MouseOver()
Public Event MouseOut()
Public Event MouseLeftDown()
Public Event MouseLeftUp()
Friend Function MessageReceived(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_MOUSELEAVE
        RaiseEvent MouseOut
    Case WM_MOUSEHOVER
        RaiseEvent MouseOver
    Case WM_MOUSEMOVE
        StartTracking
    Case WM_LBUTTONDOWN
        RaiseEvent MouseLeftDown
    Case WM_LBUTTONUP
        RaiseEvent MouseLeftUp
End Select
'StartTracking
MessageReceived = CallWindowProc(procPrevWndFunc, mTrackObject.hWnd, wMsg, wParam, lParam)
End Function

Public Function StartTracking() As Boolean
If mTrackObject Is Nothing Then
    StartTracking = False
Else
    If bTracking = True Then StopTracking
    Dim hWnd As Long
    hWnd = mTrackObject.hWnd
    colTrackMouse.Add Me, "TM" & hWnd 'so procTrackMouse knows which instance of the class to call
    procPrevWndFunc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf procTrackMouse)
    Dim tme As TRACKMOUSEEVENTINFO
    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_HOVER Or TME_LEAVE
        .dwHoverTime = 1 'HOVER_DEFAULT
        .hwndTrack = hWnd
    End With
    TrackMouseEvent tme
    bTracking = True
End If
End Function


Public Function StopTracking() As Boolean
If Not (mTrackObject Is Nothing) Then
    Dim hWnd As Long
    hWnd = mTrackObject.hWnd
    SetWindowLong hWnd, GWL_WNDPROC, procPrevWndFunc
    On Error Resume Next
    colTrackMouse.Remove "TM" & hWnd
    bTracking = False
End If
End Function
Property Get TrackObject() As Object
Set TrackObject = mTrackObject
End Property

Property Set TrackObject(obj As Object)
If obj Is Nothing Then
    StopTracking
    Set mTrackObject = Nothing
Else
    Set mTrackObject = obj
    StartTracking
End If
End Property

Private Sub Class_Terminate()
Set TrackObject = Nothing
End Sub



