VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlIcon 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
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
   HasDC           =   0   'False
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "Icon.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1410
      Picture         =   "Icon.ctx":1AB2
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3870
      Top             =   2820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const RDW_INVALIDATE = &H1
Const BS_HATCHED = 2
Const HS_CROSS = 4
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long

End Type

Dim DoubleLine As Boolean

Private Hold_Picture As Picture
Private Hold_Caption As String

Property Get Caption() As String
Caption = Hold_Caption
End Property

Property Let Caption(NewValue As String)
Hold_Caption = NewValue
DoubleLine = False
If InStr(1, NewValue, vbNewLine) Then
DoubleLine = True
Else
Hold_Caption = Hold_Caption & vbNewLine
End If
PropertyChanged "Hold_Caption"
Clear
LoadGUI
End Property

Property Get Picture() As Picture
    Set Picture = Hold_Picture
    Set Picture1.Picture = Hold_Picture
End Property

Public Property Set Picture(ByVal NewValue As Picture)
    Set Hold_Picture = NewValue
    Set Picture1.Picture = NewValue
    Clear
    LoadGUI
    PropertyChanged "Hold_Picture"
End Property
Function SelectMe()
Picture1.Cls
LoadSelGUI
LoadGUI
End Function

Private Sub UserControl_DblClick()
MsgBox 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set Hold_Picture = PropBag.ReadProperty("Hold_Picture", Nothing)
Hold_Caption = PropBag.ReadProperty("Hold_Caption", "Caption")
Set Picture1.Picture = Hold_Picture
Clear
LoadGUI
End Sub

Private Sub UserControl_Show()
Clear
LoadGUI
End Sub
Function LoadSelGUI()
Dim LB As LOGBRUSH, R As RECT, Rgn As Long, RgnRect As RECT, hBrush As Long

With Picture1
    .ForeColor = &HCD663F
    If DoubleLine = True Then
    RoundRect .hdc, 0, .ScaleHeight, .ScaleWidth, .ScaleHeight - 25, 5, 5
    Else
    RoundRect .hdc, 0, .ScaleHeight - 25, .ScaleWidth, .ScaleHeight - 10, 5, 5
    End If
    LB.lbColor = .ForeColor
    LB.lbStyle = 0
    LB.lbHatch = HS_CROSS
    hBrush = CreateBrushIndirect(LB)
    If DoubleLine = True Then
    SetRect R, 1, .ScaleHeight + 1, .ScaleWidth - 1, .ScaleHeight - 24
    Else
    SetRect R, 1, .ScaleHeight - 24, .ScaleWidth - 1, .ScaleHeight - 11
    End If
    FillRect .hdc, R, hBrush
End With
End Function

Function LoadGUI()
WriteCaption Hold_Caption, 2, 1, vbBlack
WriteCaption Hold_Caption, , , vbWhite

'Set Picture1.Picture = Hold_Picture
ImageList1.ListImages.Add 1, , Picture1.Image
UserControl.Picture = ImageList1.ListImages(1).Picture
UserControl.MaskPicture = UserControl.Picture
UserControl.MaskColor = &HFF00FF
ImageList1.ListImages.Remove 1
UserControl.Width = Picture1.Width
UserControl.Height = Picture1.Height + 7 * Screen.TwipsPerPixelX
End Function



Function WriteCaption(Caption As String, Optional Offest As Integer = 0, Optional Offest2 As Integer = 0, Optional Color As ColorConstants = vbBlack)
    Dim htext As String
    Dim lentext As Long
    Dim vh As Integer
    Dim vm As Integer
    Dim hRect As RECT
With Picture1
    htext = Caption
    lentext = Len(htext)
  
    .ForeColor = Color
    SetRect hRect, 0, 0, .ScaleWidth, .ScaleHeight
    vh = DrawText(.hdc, htext, lentext, hRect, DT_CALCRECT Or DT_CENTER)
    'MsgBox (.ScaleHeight * 0.5) - (vh * 0.5)
    SetRect hRect, 0 + Offest, (.ScaleHeight - vh) + Offest2, .ScaleWidth, (.ScaleHeight) + (vh)
    vm = DrawText(.hdc, htext, lentext, hRect, DT_CENTER)
End With
End Function

Function Clear()
Picture1.Cls
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Hold_Picture", Hold_Picture, Nothing
PropBag.WriteProperty "Hold_Caption", Hold_Caption, "Caption"
End Sub
