VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oPosition As Position
Private iWidth As Integer
Private iHeight As Integer
Private sCondition As String

Public Property Set Position(poPosition As Position)
    Set oPosition = poPosition
End Property

Public Property Get Position() As Position
    Set Position = oPosition
End Property

Public Property Let Width(piWidth As Integer)
    iWidth = piWidth
End Property

Public Property Get Width() As Integer
    Width = iWidth
End Property

Public Property Let Height(piHeight As Integer)
    iHeight = piHeight
End Property

Public Property Get Height() As Integer
    Height = iHeight
End Property

Public Property Let PrintCondition(psCondition As String)
    sCondition = psCondition
End Property

Public Property Get PrintCondition() As String
    PrintCondition = sCondition
End Property


Private Sub Class_Initialize()
    Set oPosition = New Position
End Sub
