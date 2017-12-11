VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim bFileOpen As Boolean
Dim sSource As String
Dim aStr() As String

Private Sub Class_Initialize()
    bFileOpen = False
    sSource = ""
End Sub

Public Sub Init()
    bFileOpen = True
    sSource = ""
End Sub

Public Sub CloseFile()
    If bFileOpen Then
        bFileOpen = False
        sSource = ""
    End If
End Sub

Public Sub WriteLine(ParamArray vals() As Variant)
    Dim i As Integer
    
    If bFileOpen Then
        For i = LBound(vals) To UBound(vals)
            If i = UBound(vals) Then
                sSource = sSource & CStr(vals(i)) & vbCrLf
            Else
                sSource = sSource & CStr(vals(i))
            End If
        Next
    End If
End Sub

Public Function GetString() As String
    GetString = sSource
End Function

