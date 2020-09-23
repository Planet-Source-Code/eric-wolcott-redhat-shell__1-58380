VERSION 5.00
Begin VB.Form SubMenu_Unkown 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.MenuButton MenuButton1 
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   900
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Hold_Caption    =   ""
   End
End
Attribute VB_Name = "SubMenu_Unkown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ButtonIndex As Integer

Function ShowForm()
    TransForm Me.hwnd, 0
    Me.Show
Dim i As Long
    For i = 100 To 255 Step 10
        TransForm Me.hwnd, CByte(i)
        DoEvents
    Next
    TransForm Me.hwnd, 255
End Function

Private Sub MenuButton1_MouseOver()
FrmStartMenu.Button(ButtonIndex).LoadGUI_OVER
End Sub

Private Sub Form_Load()
MenuButton1.CapBottom = True
MenuButton1.CapTop = True
MenuButton1.Left = 0
MenuButton1.Top = 0
MenuButton1.Caption = "Unkown SubMenu"
MenuButton1.Icon = i_Exclaim
MenuButton1.SubClassMe
Me.Width = MenuButton1.Width
Me.Height = MenuButton1.Top + MenuButton1.Height

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmStartMenu.Button(ButtonIndex).LoadGUI_OVER
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Visible = False
FrmStartMenu.Button(ButtonIndex).LoadGUI
MenuButton1.UnSubClassMe
End Sub



