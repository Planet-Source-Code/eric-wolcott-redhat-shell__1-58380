VERSION 5.00
Begin VB.UserControl TaskBarButton 
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
   Begin Project1.TrackMouse TrackMouse1 
      Left            =   3240
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   720
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "TaskBarButton"
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
Private Const Color_Left_1 = "23,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,13553358,16448250,6052956"
Private Const Color_Cent_1 = "23,6052956,13553358,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,15132390,16448250,6052956"
Private Const Color_Right_1 = "23,6052956,13553358,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,16448250,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956,6052956"

Private Const Color_Cent_2 = "26,6052956,10395294,14606046,12829635,10855845,9408399,9408399,8618883,8289918,8026746,7763574,7434609,7303023,6974058,6776679,6513507,6513507,6513507,6250335,6052956,5855577,5855577,5658198,5526612,5395026,7171437,4210752"
Private Const Color_Left_2 = "26,-1,6118749,6118749,6118749,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6184542,6250335,6381921,6052956,6052956,-1,6118749,11316396,12829635,12829635,12697793,12566206,12566206,12369084,12105912,11842997,11579825,11316396,11053224,10790052,10461087,10197915,9869206,9606035,9342606,9145227,8947847,8947847,8684932,8553090,8421504,6710886,4210752"
Private Const Color_Right_2 = "26,6118749,11645361,12829635,12829635,12697793,12566206,12566206,12369084,12105912,11842997,11579825,11316396,11053224,10790052,10461087,10197915,9869206,9606035,9342606,9145227,8947847,8947847,8684932,8553090,8421504,6447714,4210752,-1,6381921,6579300,6316128,6250335,6250335,6250335,6316128,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6381921,6316128,6250335,6052956,-1"

Private Color_Cent As String
Private Color_Left As String
Private Color_Right As String
Private Color_Height As Integer

Private Button_Caption As String
Private Button_OffSet As Integer
Private Button_hwnd As Long
Public Image As String
Private meOver As Boolean 'intHWND

Dim hold_Style As Style_

Public Property Get Style() As Style_
    Style = hold_Style
    LoadGUI
End Property

Public Property Let Style(strStyle As Style_)
    hold_Style = strStyle
    LoadGUI
End Property

Property Let intHWND(StrValue As Long)
    Button_hwnd = StrValue
End Property

Property Get intHWND() As Long
    intHWND = Button_hwnd
End Property


Property Let Caption(StrValue As String)
    Button_Caption = StrValue
    Picture1.Cls
    LoadGUI
    WriteCaption Button_Caption, Button_OffSet
End Property

Property Get Caption() As String
    Caption = Button_Caption
End Property

Property Let OffSet(StrValue As Integer)
    Button_OffSet = StrValue
End Property

Property Get OffSet() As Integer
    OffSet = Button_OffSet
End Property

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, x As Integer, y As Integer, Optional Gray As Boolean = True, Optional Brightness As Integer) As Integer
    If ColorPallet = "" Then Exit Function
    Dim PixCount
    Dim Colors() As String, CurrentRow, CurrentColumn, count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For count = 1 To UBound(Colors)
        If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
            If Colors(count) <> -1 Then
                If Gray = True Then
                Picture1.Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), AdjustBrightness(Colors(count), Brightness)
                Else
                Picture1.Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), MakeGrey(Colors(count))
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
Color_Left = Color_Left_1
Color_Right = Color_Right_1
Color_Height = 24
Case Longhorn
Color_Cent = Color_Cent_2
Color_Left = Color_Left_2
Color_Right = Color_Right_2
Color_Height = 27
End Select
Picture1.Top = 0
Picture1.Left = 0
Picture1.Width = UserControl.ScaleWidth
meOver = False
LoadBmpMenuLines 1, Color_Left, 0, 0
LoadBmpMenuLines UserControl.ScaleWidth - 4, Color_Cent, 2, 0
LoadBmpMenuLines 1, Color_Right, UserControl.ScaleWidth - 2, 0
UserControl.Height = Color_Height * 15 '24
Picture1.Height = UserControl.Height
WriteCaption Button_Caption, Button_OffSet
LoadBmpMenuLines 1, Image, 9, 3
End Function

Function WriteCaption(Caption As String, Optional Offest As Integer = 25)
With UserControl
    Dim htext As String
    Dim lentext As Long
    Dim vh As Integer
    Dim hRect As RECT
    htext = Caption
    lentext = Len(htext)
    SetRect hRect, 4, 0, .ScaleWidth - 4 - Offest, .ScaleHeight
    vh = DrawText(.hdc, htext, lentext, hRect, DT_CALCRECT Or DT_CENTER)
    SetRect hRect, 4 + Offest, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth - 4, (.ScaleHeight) + (vh)
    DrawText Picture1.hdc, htext, lentext, hRect, DT_LEFT
    .Refresh
End With
End Function

Private Sub TrackMouse1_MouseLeftDown()
meOver = True
LoadBmpMenuLines 1, Color_Left, 0, 0, , -50
LoadBmpMenuLines UserControl.ScaleWidth - 4, Color_Cent, 2, 0, , -50
LoadBmpMenuLines 1, Color_Right, UserControl.ScaleWidth - 2, 0, , -50
'UserControl.Height = 27 * 15 '24
WriteCaption Button_Caption, Button_OffSet
LoadBmpMenuLines 1, Image, 9, 3, , -50
End Sub

Private Sub TrackMouse1_MouseLeftUp()
LoadGUI
DoEvents
meOver = False
  If IsIconic(Button_hwnd) Then
    WindowHandle Button_hwnd, 3
  ElseIf IsZoomed(Button_hwnd) Then
    WindowHandle Button_hwnd, 4
  Else
    WindowHandle Button_hwnd, 4
  End If
    'MsgBox GetActiveWindow
End Sub

Private Sub TrackMouse1_MouseOut()
LoadGUI
meOver = False
End Sub

Private Sub TrackMouse1_MouseOver()
If meOver = False Then
meOver = True
LoadBmpMenuLines 1, Color_Left, 0, 0, , 50
LoadBmpMenuLines UserControl.ScaleWidth - 4, Color_Cent, 2, 0, , 50
LoadBmpMenuLines 1, Color_Right, UserControl.ScaleWidth - 2, 0, , 50
'UserControl.Height = 24 * 15
WriteCaption Button_Caption, Button_OffSet
LoadBmpMenuLines 1, Image, 9, 3, , 50
End If
End Sub

Private Sub UserControl_Resize()
LoadGUI
End Sub

Private Sub UserControl_Show()
LoadGUI
End Sub

Function PrintIcon(Icon As String)
LoadBmpMenuLines 1, Icon, 9, 3
End Function

Function SubClassMe()
TrackMouse1.Watch Picture1
End Function

Function UnSubClassMe()
TrackMouse1.CloseWatch
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
hold_Style = PropBag.ReadProperty("hold_Style", Red_Hat)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "hold_Style", hold_Style, Red_Hat
End Sub
