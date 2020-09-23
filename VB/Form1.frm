VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Project1.ctlIcon ctlIcon1 
      Height          =   1200
      Index           =   0
      Left            =   330
      TabIndex        =   4
      Top             =   720
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   2117
      Hold_Picture    =   "Form1.frx":B1A2
      Hold_Caption    =   "test"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enable Taskbar"
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   4995
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable Taskbar"
      Height          =   285
      Left            =   1695
      TabIndex        =   2
      Top             =   5325
      Visible         =   0   'False
      Width           =   1470
   End
   Begin Project1.TaskBar TaskBar1 
      Height          =   810
      Left            =   15
      TabIndex        =   1
      Top             =   6465
      Width           =   8115
      _extentx        =   14314
      _extenty        =   1429
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   315
      Left            =   165
      TabIndex        =   0
      Top             =   180
      Width           =   750
   End
   Begin Project1.ctlIcon ctlIcon1 
      Height          =   1200
      Index           =   1
      Left            =   375
      TabIndex        =   5
      Top             =   1890
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   2117
      Hold_Picture    =   "Form1.frx":CCF4
      Hold_Caption    =   "test"
   End
   Begin Project1.ctlIcon ctlIcon1 
      Height          =   1200
      Index           =   2
      Left            =   375
      TabIndex        =   6
      Top             =   3120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   2117
      Hold_Picture    =   "Form1.frx":E906
      Hold_Caption    =   "test"
   End
   Begin Project1.ctlIcon ctlIcon1 
      Height          =   1200
      Index           =   3
      Left            =   435
      TabIndex        =   7
      Top             =   4305
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   2117
      Hold_Picture    =   "Form1.frx":10458
      Hold_Caption    =   "test"
   End
   Begin Project1.ctlIcon ctlIcon1 
      Height          =   1200
      Index           =   4
      Left            =   435
      TabIndex        =   8
      Top             =   5550
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   2117
      Hold_Picture    =   "Form1.frx":11D6A
      Hold_Caption    =   "test"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Gradient As clsGradient

Dim OldTaskBarHeight
Dim StartMenuOn As Boolean

Private Sub Command1_Click()
UnloadAll
End Sub
Function UnloadAll()
ResetWorkArea OldTaskBarHeight
TaskBar1.SubClassOff
Unload Me
FrmStartMenu.UnloadAll
Unload FrmStartMenu.xForm
Unload FrmStartMenu
Unload frmMenu
Unload SubMenu_Internet
Unload SubMenu_Unkown
Unload FrmTaskBarMenu
Unload Me
If Environment = EnvironCompiled Then End
End Function
Private Sub Command2_Click()
Dim x As Integer
For x = 1 To 7
TaskBar1.ButtonEnable x, False
Next
End Sub

Private Sub Command3_Click()
Dim x As Integer
For x = 1 To 7
TaskBar1.ButtonEnable x, True
Next
End Sub

Function LoadIcons()
ctlIcon1(0).Caption = "Zach's" & vbNewLine & "Home"
ctlIcon1(1).Caption = "Help"
ctlIcon1(2).Caption = "Internet" & vbNewLine & "Browser"
ctlIcon1(3).Caption = "Trash" & vbNewLine & "Bin"
ctlIcon1(4).Caption = "Graphics"

Dim x
For x = 0 To 4
ctlIcon1(x).Left = 330 / Screen.TwipsPerPixelX
Next
End Function

Private Sub ctlIcon1_GotFocus(Index As Integer)
ctlIcon1(Index).SelectMe
End Sub

Private Sub ctlIcon1_LostFocus(Index As Integer)
ctlIcon1(Index).Clear
ctlIcon1(Index).LoadGUI
End Sub

Private Sub Form_Load()
StartMenuOn = False
TaskBar1.Width = Screen.Width
TaskBar1.AddButton RedHat, True, True
TaskBar1.AddButton internet, False, True
TaskBar1.AddButton Mail, False, True
TaskBar1.AddButton Letter, False, True
TaskBar1.AddButton Graph, False, True
TaskBar1.AddButton SpreadSheet, False, True
TaskBar1.AddButton Printer, False, True
TaskBar1.AddButton Spacer, False, True
TaskBar1.LoadTasks
'+ Form1.Height / Screen.TwipsPerPixelY
Dim rtn
OldTaskBarHeight = GetTaskbarHeight / Screen.TwipsPerPixelY
SetSysWorkArea 0, 0, Screen.Width / Screen.TwipsPerPixelX, Int(Screen.Height / Screen.TwipsPerPixelY) - 54 'TaskBar1.Height 'TaskBar1.Height - 500 '(TaskBar1.Height / Screen.TwipsPerPixelY)
rtn = FindWindow("Shell_traywnd", vbNullString)
SetWindowPos rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW

Me.Visible = True

Window "explorer.exe", 2
Window "Shell_traywnd", 2

Load FrmStartMenu
FrmStartMenu.Left = 0 'Screen.TwipsPerPixelX * 3
FrmStartMenu.Top = Form1.Height - FrmStartMenu.ScaleHeight - (TaskBar1.Height * Screen.TwipsPerPixelY) '- (1 * Screen.TwipsPerPixelY)

LoadIcons
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
TaskBar1.ResetLastButton

If FrmStartMenu.ButtonOver = True Then
FrmStartMenu.xForm.Visible = False
FrmStartMenu.ButtonOver = False
StartMenuOn = True
End If

If StartMenuOn = True Then
StartMenuOn = False
FrmStartMenu.HideForm
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Window "explorer.exe", 1
Window "Shell_traywnd", 1
End Sub

Private Sub Form_Resize()
Set Gradient = New clsGradient
Gradient.Gradient Form1, 1, 4398867, 8143399
TaskBar1.Top = Me.ScaleHeight - TaskBar1.Height
TaskBar1.Left = 0
TaskBar1.Width = Me.ScaleWidth
End Sub

Private Sub TaskBar1_ButtonClicked(Index As Integer, Button As Integer)
Select Case Index
Case 1
'FrmStartMenu.Visible = True
StartMenuOn = True
FrmStartMenu.ShowForm
FrmStartMenu.MeTop
FrmStartMenu.ZOrder 0
End Select
End Sub
