VERSION 5.00
Begin VB.Form FrmStartMenu 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   LinkTopic       =   "Form2"
   ScaleHeight     =   9090
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.MenuButton Button 
      Height          =   375
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   225
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   661
      Hold_Caption    =   ""
   End
End
Attribute VB_Name = "FrmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xForm As Form

Dim CurrentOver As Integer
Dim LastOver As Integer
Public ButtonOver As Boolean

Private Sub Button_MouseOver(Index As Integer)
If Index = CurrentOver Then Exit Sub
LastOver = CurrentOver
CurrentOver = Index
If xForm.hwnd = frmMenu.hwnd Then Exit Sub
'Unload xForm
FrmStartMenu.Button(LastOver).LoadGUI
xForm.Visible = False
If Button(Index).Arrow = False Then Exit Sub
ButtonOver = True
Select Case Index
Case 4
Set xForm = SubMenu_Internet
Case Else
Set xForm = SubMenu_Unkown
End Select
'xForm.Show
xForm.Top = Me.Top + Button(Index).Top - (2 * Screen.TwipsPerPixelY)
xForm.Left = Me.Left + Me.Width - 75
xForm.ShowForm
xForm.ZOrder 0
xForm.ButtonIndex = Index
SetWindowPos xForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

'Set xForm = Nothing
End Sub

Private Sub Button_MouseUp(Index As Integer)
xForm.ZOrder 0
Select Case Index
Case 17
Form1.UnloadAll
End Select

End Sub

Private Sub Form_Load()
ButtonOver = False
Set xForm = New frmMenu
Dim X, Y
For X = 0 To 17
    If X = 0 Then
        Button(X).CapTop = True
        Button(X).CapBottom = True
        Button(X).Top = 0
        Button(X).Caption = "The GIMP"
    Else
        Load Button(X)
        If X = 17 Then
        Button(X).CapBottom = True
        Button(X).CapTop = False
        ElseIf X = 15 Then
        Button(X).CapBottom = True
        Button(X).CapTop = True
        Else
        Button(X).CapBottom = False
        Button(X).CapTop = False
        End If
        If X < 11 Then
        Button(X).Arrow = True
        End If
        Button(X).Top = Button(X - 1).Top + Button(X - 1).Height
        Button(X).Visible = True
        Select Case X
        Case 1
        Button(X).Caption = "Accessories"
        Case 2
        Button(X).Caption = "Games"
        Case 3
        Button(X).Caption = "Graphics"
        Case 4
        Button(X).Caption = "Internet"
        Case 5
        Button(X).Caption = "Office"
        Case 6
        Button(X).Caption = "Preferences"
        Case 7
        Button(X).Caption = "Programming"
        Case 8
        Button(X).Caption = "Sound && Video"
        Case 9
        Button(X).Caption = "System Settings"
        Case 10
        Button(X).Caption = "System Tools"
        Case 11
        Button(X).Caption = "Control Center"
        Case 12
        Button(X).Caption = "Find Files"
        Case 13
        Button(X).Caption = "Help"
        Case 14
        Button(X).Caption = "Home"
        Case 15
        Button(X).Caption = "Run Command..."
        Case 16
        Button(X).Caption = "System Lock"
        Case 17
        Button(X).Caption = "Logout 'User'"
        End Select
    End If
        Button(X).Left = 0
        Button(X).SubClassMe
        Button(X).Icon = 18 - X
Next

Me.Width = Button(0).Width
Me.Height = Button(17).Top + Button(17).Height
End Sub

Private Sub Form_LostFocus()
Me.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnloadAll
End Sub

Function UnloadAll()
Dim X
For X = 0 To 17
        Button(X).UnSubClassMe
Next
End Function

Function MeTop()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'SetActiveWindow Me.hwnd
End Function

Function ShowForm()
    TransForm Me.hwnd, 0
    Me.Show
Dim i As Long
    For i = 0 To 255 Step 5
        TransForm Me.hwnd, CByte(i)
        DoEvents
    Next
    TransForm Me.hwnd, 255
End Function

Function HideForm()
'ButtonOver = False
xForm.Visible = False
DoEvents
Dim i As Long
    For i = 255 To 0 Step -5
        TransForm Me.hwnd, CByte(i)
        DoEvents
    Next
    TransForm Me.hwnd, 0
    Me.Hide
End Function



