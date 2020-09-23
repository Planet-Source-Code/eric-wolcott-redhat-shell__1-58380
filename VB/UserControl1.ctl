VERSION 5.00
Begin VB.UserControl EnumTasks 
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   3870
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   600
      Left            =   -60
      Picture         =   "UserControl1.ctx":0312
      ScaleHeight     =   540
      ScaleWidth      =   570
      TabIndex        =   2
      Top             =   -45
      Width           =   630
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   705
      TabIndex        =   1
      Top             =   60
      Width           =   2430
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   705
      TabIndex        =   0
      Top             =   360
      Width           =   2430
   End
End
Attribute VB_Name = "EnumTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function addApp(Item As String)
    ListApp.AddItem Item
End Function

Public Function GetAppName(Index As Integer) As String
    GetAppName = ListAppName.List(Index)
End Function

Public Function GetAppHWND(Index As Integer) As String
    GetAppHWND = ListApp.List(Index)
End Function

Public Function GetWindows() As String
    EnumMe
End Function

Public Property Get RunningAppCount() As String
    RunningAppCount = ListApp.ListCount
End Property

Function GetIcon(picture As PictureBox, Index As Integer)
DrawIcon picture.hDC, GetAppHWND(Index), 0, 0
End Function

Private Function EnumMe()
    Set ListApp = List1
    Set ListAppName = List2
    fEnumWindows
End Function


Public Sub DrawIcon(hDC As Long, hwnd As Long, X As Integer, Y As Integer)
ico = GetIcons(hwnd)
DrawIconEx hDC, X, Y, ico, 16, 16, 0, 0, DI_NORMAL
End Sub

Public Function GetIcons(hwnd As Long) As Long
Call SendMessageTimeout(hwnd, WM_GETICON, 0, 0, 0, 1000, GetIcons)
If Not CBool(GetIcons) Then GetIcons = GetClassLong(hwnd, GCL_HICONSM)
If Not CBool(GetIcons) Then Call SendMessageTimeout(hwnd, WM_GETICON, 1, 0, 0, 1000, GetIcons)
If Not CBool(GetIcons) Then GetIcons = GetClassLong(hwnd, GCL_HICON)
If Not CBool(GetIcons) Then Call SendMessageTimeout(hwnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcons)
End Function

