VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TaskBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   3120
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2700
      ScaleHeight     =   18
      ScaleMode       =   2  'Point
      ScaleWidth      =   18
      TabIndex        =   1
      Top             =   735
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2265
      Top             =   1590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin Project1.TaskBarButton TaskBarButton1 
      Height          =   360
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   635
   End
   Begin Project1.EnumTasks EnumTasks1 
      Left            =   1125
      Top             =   2280
      _ExtentX        =   1138
      _ExtentY        =   1032
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3735
      Top             =   3090
   End
End
Attribute VB_Name = "TaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Const RDW_INVALIDATE = &H1
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type
Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

'15132390
Private Const Color_Cap = "53,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329"
Private Const Color_Cent_1 = "53,10066329,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,-2,10066329"
Private Const Color_Filled = "53,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,10066329,-1,-1"
Private Const Color_Cent_2 = "53,8947848,14606046,12829635,10855845,9342606,8553090,8289918,8026746,7763574,7434609,7303023,6974058,6776679,6513507,6250335,6052956,6052956,6052956,5855577,5658198,5526612,5395026,5263440,5131854,5131854,5131854,5131854,5197647,5395026,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612,5526612"
Private Color_Cent As String
Private LastButtonLeft As Integer
Private CurrentButtonOver As Integer

Public fClassList As New Collection
Public fClassButton As New Collection

Public Event ButtonClicked(Index As Integer, Button As Integer)

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Type Buttons_
Left As Integer
Top As Integer
count As Integer
MaxWidth As Integer
BarWidth As Integer
MODcount As Integer
hwnd As Variant
End Type

Public TaskBarColor

Dim Buttons As Buttons_

Dim hold_Style As Style_

Public Property Get Style() As Style_
    Style = hold_Style
    TaskBarButton1(0).Style = hold_Style
    LoadGUI
End Property

Public Property Let Style(strStyle As Style_)
    hold_Style = strStyle
    TaskBarButton1(0).Style = strStyle
    LoadGUI
End Property

Function Repaint()
    EmptyButtonBin
    LoadTasks
    Dim x As Integer
    For x = 0 To fClassList.count - 1
    RepaintIcon x
    Next
End Function


Function RipPicture(TransColor As ColorConstants) As String
        Dim i As Integer
        Dim j As Integer
        Dim Temp As String
        Temp = Temp & Picture1.ScaleHeight - 1 & ","
        Do Until i >= Picture1.ScaleWidth
            j = 0
            Do Until j >= Picture1.ScaleHeight
            DoEvents
                Dim CurrColor As Long
                CurrColor = GetPixel(Picture1.hdc, i, j)
                If CurrColor = TransColor Then CurrColor = -1
                Temp = Temp & CurrColor & ","
                j = j + 1
            Loop
            i = i + 1
        Loop
        Temp = Left(Temp, Len(Temp) - 1)
        RipPicture = Temp
End Function
Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, x As Integer, y As Integer, Optional Gray As Boolean = True, Optional Brightness As Integer) As Integer
    Dim PixCount
    Dim Colors() As String, CurrentRow, CurrentColumn, count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For count = 1 To UBound(Colors)
        If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
            If Colors(count) <> -1 Then
                If Colors(count) = -2 Then Colors(count) = TaskBarColor
                If Gray = True Then
                UserControl.Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), AdjustBrightness(Colors(count), Brightness)
                Else
                UserControl.Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), MakeGrey(Colors(count))
                End If
            End If
        CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function

Function LoadGUI()
    Select Case hold_Style
    Case Red_Hat
    Color_Cent = Color_Cent_1
    Case Longhorn
    Color_Cent = Color_Cent_2
    End Select
TaskBarColor = 15132390
LoadBmpMenuLines 1, Color_Cap, 0, 0
LoadBmpMenuLines UserControl.ScaleWidth - 2, Color_Cent, 1, 0
LoadBmpMenuLines 1, Color_Cap, UserControl.ScaleWidth - 1, 0
UserControl.Height = 54 * 15
End Function



Private Sub Timer1_Timer()
With UserControl
    LoadBmpMenuLines 99, Color_Cent, UserControl.ScaleWidth - 100, 0
    Dim htext As String
    Dim lentext As Long
    Dim vh As Integer
    Dim hRect As RECT
    htext = FormatDateTime(Date, vbLongDate) & vbNewLine & Time
    lentext = Len(htext)
    SetRect hRect, 4, 0, .ScaleWidth - 4, .ScaleHeight
    vh = DrawText(.hdc, htext, lentext, hRect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
    SetRect hRect, 4, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth - 4, (.ScaleHeight * 0.5) + (vh * 0.5)
    DrawText .hdc, htext, lentext, hRect, DT_RIGHT Or DT_WORDBREAK
    .Refresh
End With
End Sub

Private Sub Timer2_Timer()
EnumTasks1.GetWindows
If Buttons.count <> EnumTasks1.RunningAppCount Then
Buttons.count = EnumTasks1.RunningAppCount
LoadTasks
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim HoldIndex As Integer
HoldIndex = GetButtonIndex(x)
If HoldIndex <> 0 Then
If fClassList(HoldIndex).IconIndex <> 0 Then
If fClassList(HoldIndex).Enabled = True Then
LoadBmpMenuLines fClassList(HoldIndex).Right - fClassList(HoldIndex).Left - 1, Color_Cent, fClassList(HoldIndex).Left, 0
ReadIcon fClassList(HoldIndex).IconIndex, fClassList(HoldIndex).Left, 2, 2, fClassList(HoldIndex).Arrow, fClassList(HoldIndex).Enabled
End If
End If
End If
If Button = vbRightButton Then
FrmTaskBarMenu.Top = Screen.Height - UserControl.Height - FrmTaskBarMenu.Height
FrmTaskBarMenu.Left = x * Screen.TwipsPerPixelX
If FrmTaskBarMenu.Left + FrmTaskBarMenu.Width > Screen.Width Then FrmTaskBarMenu.Left = Screen.Width - FrmTaskBarMenu.Width

FrmTaskBarMenu.ShowForm
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RepaintIcon GetButtonIndex(x)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim HoldIndex As Integer
HoldIndex = GetButtonIndex(x)
If HoldIndex <> 0 Then
If fClassList(HoldIndex).IconIndex <> 0 Then
If fClassList(HoldIndex).Enabled = True Then
LoadBmpMenuLines fClassList(HoldIndex).Right - fClassList(HoldIndex).Left, Color_Cent, fClassList(HoldIndex).Left, 0
ReadIcon fClassList(HoldIndex).IconIndex, fClassList(HoldIndex).Left, 2, 2, fClassList(HoldIndex).Arrow, fClassList(HoldIndex).Enabled, 50
RaiseEvent ButtonClicked(HoldIndex, Button)
End If
End If
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
hold_Style = PropBag.ReadProperty("hold_Style", Red_Hat)
End Sub

Private Sub UserControl_Resize()
LoadGUI
End Sub

Function AddButton(Icon As IconList, Arrow As Boolean, Enabled As Boolean)
Dim x As Integer
x = LastButtonLeft + ReadIcon(Icon, LastButtonLeft, , , Arrow, Enabled) + 5
If Arrow = True Then
LoadBmpMenuLines 1, Color_ARROW, LastButtonLeft, 0
End If
AddListData LastButtonLeft, x, Icon, Arrow, Enabled
LastButtonLeft = x
CurrentButtonOver = -1
End Function


Private Function AddListData(Left As Integer, Right As Integer, Icon As IconList, Arrow As Boolean, Enabled As Boolean)
    Dim Container As New TaskBar_Container
    Container.Left = Left
    Container.Right = Right
    Container.IconIndex = Icon
    Container.Arrow = Arrow
    Container.Enabled = Enabled
    fClassList.Add Container
End Function

Function GetButtonIndex(x) As Integer
Dim i
For i = 1 To fClassList.count
If x > fClassList(i).Left And x < fClassList(i).Right Then
GetButtonIndex = i
Exit Function
End If
Next
End Function

Function ReadIcon(Index As IconList, Left As Integer, Optional OffsetX As Integer = 0, Optional OffsetY As Integer = 0, Optional Arrow As Boolean = False, Optional Enabled As Boolean = True, Optional Brightness As Integer = 0) As Integer
Select Case Index
Case 0
ReadIcon = LoadBmpMenuLines(1, Color_SPACER, Left + OffsetX, OffsetY)
Case 1
ReadIcon = LoadBmpMenuLines(1, Color_REDHAT, Left + OffsetX, OffsetY, Enabled, Brightness)
Case 2
ReadIcon = LoadBmpMenuLines(1, Color_INTERNET, Left + OffsetX, OffsetY, Enabled, Brightness)
Case 3
ReadIcon = LoadBmpMenuLines(1, Color_MAIL, Left + OffsetX, OffsetY, Enabled, Brightness)
Case 4
ReadIcon = LoadBmpMenuLines(1, Color_LETTER, Left + OffsetX, OffsetY, Enabled, Brightness)
Case 5
ReadIcon = LoadBmpMenuLines(1, Color_GRAPH, Left + OffsetX, OffsetY, Enabled, Brightness)
Case 6
ReadIcon = LoadBmpMenuLines(1, Color_SPREADSHEET, Left + OffsetX, OffsetY, Enabled, Brightness)
Case 7
ReadIcon = LoadBmpMenuLines(1, Color_PRINTER, Left + OffsetX, OffsetY, Enabled, Brightness)
End Select

If Arrow = True And Index <> Spacer Then
LoadBmpMenuLines 1, Color_ARROW, Left + OffsetX, OffsetY
End If

End Function

Function RepaintIcon(Index As Integer)
If Index <> 0 Then
If CurrentButtonOver <> Index Then
If fClassList(Index).IconIndex <> 0 Then
If fClassList(Index).Enabled = True Then
LoadBmpMenuLines fClassList(Index).Right - fClassList(Index).Left, Color_Cent, fClassList(Index).Left, 0
'LoadBmpMenuLines 1, Color_OVER, fClassList(Index).Left, 1
ReadIcon fClassList(Index).IconIndex, fClassList(Index).Left, 2, 2, fClassList(Index).Arrow, fClassList(Index).Enabled, 50
End If
End If
ResetLastButton
CurrentButtonOver = Index
End If
Else
ResetLastButton
End If
End Function

Function ResetLastButton()
If CurrentButtonOver <> -1 Then
LoadBmpMenuLines fClassList(CurrentButtonOver).Right - fClassList(CurrentButtonOver).Left, Color_Cent, fClassList(CurrentButtonOver).Left, 0
ReadIcon fClassList(CurrentButtonOver).IconIndex, fClassList(CurrentButtonOver).Left, 0, 0, fClassList(CurrentButtonOver).Arrow, fClassList(CurrentButtonOver).Enabled
End If
CurrentButtonOver = -1
End Function

Function ButtonEnable(Index As Integer, Enable As Boolean)
fClassList(Index).Enabled = Enable
RepaintIcon Index
ResetLastButton
End Function

Function LoadWindows()
EnumTasks1.GetWindows
    Dim x As Integer, y As Integer, w, z
    y = EnumTasks1.RunningAppCount
For x = 0 To y - 1
w = EnumTasks1.GetAppName(x)
z = EnumTasks1.GetAppHWND(x)
    EnumTasks1.GetIcon Picture1, x
    ImageList1.ListImages.Add , , Picture1.Image
    Picture1.Cls
    ListView1.ListItems.Add , , w, x + 1, x + 1
Next
End Function

Function LoadTasks()
Timer1.Enabled = True
EmptyButtonBin

EnumTasks1.GetWindows
Buttons.count = EnumTasks1.RunningAppCount
Buttons.Left = fClassList(fClassList.count).Right + 2
Buttons.MaxWidth = UserControl.ScaleWidth - fClassList(fClassList.count).Right - 150
Buttons.Top = 5
Buttons.BarWidth = 0
Buttons.MODcount = EnumTasks1.RunningAppCount Mod 2
Dim x As Integer

For x = 0 To EnumTasks1.RunningAppCount - 1
LoadTaskButton x
LoadEmpytButton
Next

Timer2.Enabled = True
End Function

Function LoadTaskButton(Index As Integer)
    Dim TempWidth As Integer
    If Buttons.MODcount = 0 Then
    TempWidth = EnumTasks1.RunningAppCount / 2
    Else
    TempWidth = (EnumTasks1.RunningAppCount + 1) / 2
    End If
    TaskBarButton1(Index).Width = (Buttons.MaxWidth - 30) / TempWidth
    Buttons.BarWidth = Buttons.BarWidth + TaskBarButton1(Index).Width
    If Buttons.BarWidth >= Buttons.MaxWidth Then
        Buttons.Top = Buttons.Top + TaskBarButton1(Index).Height
        Buttons.BarWidth = 0
        Buttons.Left = fClassList(fClassList.count).Right + 2
    End If
    TaskBarButton1(Index).Left = Buttons.Left
    TaskBarButton1(Index).Top = Buttons.Top
    TaskBarButton1(Index).OffSet = 25
    TaskBarButton1(Index).Visible = True
    TaskBarButton1(Index).Caption = EnumTasks1.GetAppName(Index)
    TaskBarButton1(Index).intHWND = EnumTasks1.GetAppHWND(Index)
    EnumTasks1.GetIcon Picture1, Index
    ImageList1.ListImages.Add , , Picture1.Image
    Dim XY As String
    XY = RipPicture(Picture1.BackColor)
    TaskBarButton1(Index).Image = XY
    AddListButtonData XY
    TaskBarButton1(Index).PrintIcon XY
    Picture1.Cls
    Buttons.Left = Buttons.Left + TaskBarButton1(Index).Width
    
    TaskBarButton1(Index).SubClassMe
End Function

Private Function AddListButtonData(Image As String)
    Dim Container As New TaskBarButton_Container
    Container.Image = RipPicture(Picture1.BackColor)
    DoEvents
    fClassButton.Add Container
End Function

Function LoadEmpytButton()
Load TaskBarButton1(TaskBarButton1.count)
End Function

Function EmptyButtonBin()
Dim x
TaskBarButton1(0).UnSubClassMe
For x = 1 To TaskBarButton1.count - 1
TaskBarButton1(x).UnSubClassMe
Unload TaskBarButton1(x)
Next
End Function

Public Function SubClassOff()
EmptyButtonBin
End Function


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "hold_Style", hold_Style, Red_Hat
End Sub
