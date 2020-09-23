VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum eEnvironment
    EnvironIDE = 1
    EnvironCompiled = 2
End Enum


Public Property Get QueryEnvironment() As eEnvironment
    QueryEnvironment = EnvironCompiled
    Debug.Assert Not SetToIDE(QueryEnvironment)
End Property


Private Function SetToIDE(Env As eEnvironment) As Boolean
    Env = EnvironIDE
End Function

