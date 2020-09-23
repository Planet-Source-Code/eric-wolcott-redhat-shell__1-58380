VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskBar_Container"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Button_Left As Integer
Private Button_Right As Integer
Private Button_PictureIndex As IconList
Private Button_Arrow As Boolean
Private Button_Enabled As Boolean

Property Let Enabled(StrValue As Boolean)
    Button_Enabled = StrValue
End Property

Property Get Enabled() As Boolean
    Enabled = Button_Enabled
End Property

Property Let Arrow(StrValue As Boolean)
    Button_Arrow = StrValue
End Property

Property Get Arrow() As Boolean
    Arrow = Button_Arrow
End Property

Property Let IconIndex(StrValue As IconList)
    Button_PictureIndex = StrValue
End Property

Property Get IconIndex() As IconList
    IconIndex = Button_PictureIndex
End Property

Property Let Left(ByVal StrValue As Integer)
    Button_Left = StrValue
End Property

Property Get Left() As Integer
    Left = Button_Left
End Property

Property Let Right(ByVal StrValue As Integer)
    Button_Right = StrValue
End Property

Property Get Right() As Integer
    Right = Button_Right
End Property

