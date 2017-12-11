VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim iFile As Integer
Dim bFileOpen As Boolean
Dim aStr() As String

Private Sub Class_Initialize()
    iFile = FreeFile
    bFileOpen = False
End Sub

Public Sub Init(psFileName As String)
    Open psFileName For Output As #iFile
    bFileOpen = True
End Sub

Public Sub CloseFile()
    If bFileOpen Then
        Close #iFile
        bFileOpen = False
    End If
End Sub

Public Sub WriteLine(ParamArray vals() As Variant)
    Dim i As Integer
    
    If bFileOpen Then
        For i = LBound(vals) To UBound(vals)
            If i = UBound(vals) Then
                Write #iFile, CStr(vals(i))
            Else
                Write #iFile, CStr(vals(i));
            End If
        Next
    End If
End Sub

Private Function ConvertToString(vals() As Variant) As String()
    Dim i As Integer
    
    ReDim aStr(UBound(vals))
    
    For i = LBound(vals) To UBound(vals)
        aStr(i) = CStr(vals(i))
    Next
    
    ConvertToString = aStr
End Function
Private Sub Class_Terminate()
    Call CloseFile
End Sub
