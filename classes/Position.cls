VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim iX As Integer
Dim iY As Integer

Public Property Let X(piX As Integer)
    iX = piX
End Property

Public Property Get X() As Integer
    X = iX
End Property

Public Property Let Y(piY As Integer)
    iY = piY
End Property

Public Property Get Y() As Integer
    Y = iY
End Property

