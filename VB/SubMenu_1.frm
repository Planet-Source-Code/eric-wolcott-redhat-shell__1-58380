VERSION 5.00
Begin VB.Form SubMenu_Internet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.MenuButton Button 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   345
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   661
      Hold_Caption    =   ""
   End
End
Attribute VB_Name = "SubMenu_Internet"
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

Private Sub Button_MouseOver(Index As Integer)
FrmStartMenu.Button(ButtonIndex).LoadGUI_OVER
End Sub

Private Sub Form_Load()
Dim X, Y
For X = 0 To 5
    If X = 0 Then
        Button(X).CapTop = True
        Button(X).CapBottom = True
        Button(X).Top = 0
        Button(X).Caption = "Launch Browser"
    Else
        Load Button(X)
        If X = 5 Then
        Button(X).CapBottom = True
        Button(X).CapTop = False
        Else
        Button(X).CapBottom = False
        Button(X).CapTop = False
        End If
        
        Button(X).Top = Button(X - 1).Top + Button(X - 1).Height
        Button(X).Visible = True
        Select Case X
        Case 1
        Button(X).Caption = "Network Connections"
        Case 2
        Button(X).Caption = "Setup New Connection"
        Case 3
        Button(X).Caption = "Network Information"
        Case 4
        Button(X).Caption = "Bandwidth Monitor"
        Case 5
        Button(X).Caption = "Internet Settings"
        End Select
    End If
        Button(X).Left = 0
        Button(X).SubClassMe
        'Button(X).Icon = 18 - X
Next

Me.Width = Button(0).Width
Me.Height = Button(5).Top + Button(5).Height

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmStartMenu.Button(ButtonIndex).LoadGUI_OVER
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Visible = False
FrmStartMenu.Button(ButtonIndex).LoadGUI
UnloadAll
End Sub

Function UnloadAll()
Dim X
For X = 0 To 5
        Button(X).UnSubClassMe
Next
End Function

