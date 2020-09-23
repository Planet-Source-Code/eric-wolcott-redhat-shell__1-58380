VERSION 5.00
Begin VB.Form FrmTaskBarMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.MenuButton Button 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   661
      Hold_Caption    =   ""
   End
End
Attribute VB_Name = "FrmTaskBarMenu"
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

Private Sub Button_MouseDown(Index As Integer)
Select Case Index
Case 0

Case 1
Form1.TaskBar1.Style = Red_Hat
Form1.TaskBar1.Repaint
Case 2
Form1.TaskBar1.Style = Longhorn
Form1.TaskBar1.Repaint
Case 3

Case 4

Case 5

End Select
Me.Hide
End Sub

Private Sub Button_MouseOver(Index As Integer)
FrmStartMenu.Button(ButtonIndex).LoadGUI_OVER
End Sub

Private Sub Form_Load()
Dim x, y
For x = 0 To 5
    If x = 0 Then
        Button(x).CapTop = True
        Button(x).CapBottom = True
        Button(x).Top = 0
        Button(x).Caption = "Select Color"
    Else
        Load Button(x)
        If x = 5 Then
        Button(x).CapBottom = True
        Button(x).CapTop = False
        Else
        Button(x).CapBottom = False
        Button(x).CapTop = False
        End If
        
        Button(x).Top = Button(x - 1).Top + Button(x - 1).Height
        Button(x).Visible = True
        Select Case x
        Case 1
        Button(x).Caption = "Red Hat Theme"
        Case 2
        Button(x).Caption = "Longhorn Theme"
        Case 3
        Button(x).Caption = "XP Blue Theme"
        Case 4
        Button(x).Caption = "XP Green Theme"
        Case 5
        Button(x).Caption = "XP Silver Theme"
        End Select
    End If
        Button(x).Left = 0
        Button(x).SubClassMe
        'Button(X).Icon = 18 - X
Next

Me.Width = Button(0).Width
Me.Height = Button(5).Top + Button(5).Height

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
FrmStartMenu.Button(ButtonIndex).LoadGUI_OVER
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Visible = False
FrmStartMenu.Button(ButtonIndex).LoadGUI
UnloadAll
End Sub

Function UnloadAll()
Dim x
For x = 0 To 5
        Button(x).UnSubClassMe
Next
End Function
