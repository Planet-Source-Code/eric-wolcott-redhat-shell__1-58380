Attribute VB_Name = "HandleWindows"
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Public Const WM_CLOSE = &H10
    Public Const SW_HIDE = 0
    Public Const SW_MAXIMIZE = 3
    Public Const SW_SHOW = 5
    Public Const SW_MINIMIZE = 6
    
Private Type RECT
    Left                                              As Long
    Top                                               As Long
    Right                                             As Long
    Bottom                                            As Long
End Type
Public Const SPI_GETWORKAREA    As Integer = 48
Public Const SWP_HIDEWINDOW    As Long = &H80
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function SetActiveWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Private AkhilSt                                   As RECT

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                   ByVal hWndInsertAfter As Long, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long, _
                                                   ByVal cx As Long, _
                                                   ByVal cy As Long, _
                                                   ByVal wFlags As Long) As Long
                                                   
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2

Public Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hWnd As Long, _
  ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hWnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Public Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hWnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public LastAlpha As Long

Sub WindowHandle(win, Cas As Long)
    'Case 0 = CloseWindow
    'Case 1 = Show Win
    'Case 2 = Hide Win
    'Case 3 = Max Win
    'Case 4 = Min Win
    Select Case Cas
        Case 0:
        Dim x%
            x% = SendMessage(win, WM_CLOSE, 0, 0)
        Case 1:
            x = ShowWindow(win, SW_SHOW)
        Case 2:
            x = ShowWindow(win, SW_HIDE)
        Case 3:
            x = ShowWindow(win, SW_MAXIMIZE)
            SetWindowPos win, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
            SetActiveWindow win
            'DoEvents
            'Dim hInstance As Long
            'hInstance = GetActiveWindow()
            'MsgBox hInstance & ":" & win
        Case 4:
            x = ShowWindow(win, SW_MINIMIZE)
    End Select

End Sub

Function Window(Name As String, Cas As Long)
Dim tWnd As Long
tWnd = FindWindow(Name, vbNullString)
If tWnd <> 0 Then WindowHandle tWnd, Cas
End Function

Public Sub SetSysWorkArea(ByVal l As Integer, ByVal t As Integer, ByVal r As Integer, ByVal b As Integer)
    With AkhilSt
        .Left = l
        .Top = t
        .Right = r
        .Bottom = b
    End With 'AkhilSt
    Akhil = SystemParametersInfo(47, 0, AkhilSt, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
End Sub

Function ResetWorkArea(NewHeight)
    rtn = FindWindow("Shell_traywnd", vbNullString)
    SetWindowPos rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW
    SetSysWorkArea 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY - NewHeight
    SetWindowPos rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW
    'MsgBox (GetTaskbarHeight / Screen.TwipsPerPixelY)
End Function

Public Function GetTaskbarHeight() As Integer
    Dim rectVal As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, rectVal, 0
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

Public Function TransForm(fhWnd As Long, Alpha As Byte) As Boolean
'Set alpha between 0-255
' 0 = Invisible , 128 = 50% transparent , 255 = Opaque
    SetWindowLong fhWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes fhWnd, 0, Alpha, LWA_ALPHA
    LastAlpha = Alpha
End Function
