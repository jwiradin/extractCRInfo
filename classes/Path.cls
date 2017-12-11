VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const API_NULL As Long = 0

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" ( _
    ByVal lpModuleName As Long) As Long

Private Declare Function GetModuleBaseName Lib "psapi" _
    Alias "GetModuleBaseNameA" ( _
    ByVal hProcess As Long, _
    ByVal hModule As Long, _
    ByVal BaseName As String, _
    ByVal nSize As Long) As Long

Private Declare Function GetModuleFileNameEx Lib "psapi" _
    Alias "GetModuleFileNameExA" ( _
    ByVal hProcess As Long, _
    ByVal hModule As Long, _
    ByVal FileName As String, _
    ByVal nSize As Long) As Long

Public Function GetExtension(sFile As String) As String
    GetExtension = sFile
    If Trim$(sFile) <> "" Then
        GetExtension = Right$(Trim$(sFile), Len(Trim$(sFile)) - InStrRev(Trim$(sFile), "."))
    End If
End Function

Public Function GetFileName(sFile As String) As String
    GetFileName = sFile
    If Trim$(sFile) <> "" Then
        GetFileName = Right$(Trim$(sFile), Len(Trim$(sFile)) - InStrRev(Trim$(sFile), "\"))
    End If
End Function

Public Function GetFileNameOnly(sFile As String) As String
    GetFileNameOnly = sFile
    
    If Trim$(sFile) <> "" Then
        If InStr(sFile, ".") > 1 Then
            GetFileNameOnly = Right$(Trim$(sFile), Len(Trim$(sFile)) - InStrRev(Trim$(sFile), "\"))
            GetFileNameOnly = Left$(GetFileNameOnly, InStr(GetFileNameOnly, ".") - 1)
        End If
    End If
End Function

Public Function Combine(sParam As String, sParam1 As String, ParamArray args() As Variant) As String
    Dim sResult As String
    Dim iCnt As Integer
    
    sResult = Trim$(sParam) & IIf(Right$(Trim$(sParam), 1) = "\", "", "\") & Trim$(sParam1)
    
    For iCnt = 0 To UBound(args)
        sResult = Combine(sResult, CStr(args(iCnt)))
    Next
    
    Combine = sResult
End Function


Public Function GetExeBaseName() As String
    Dim Size As Long

    GetExeBaseName = Space$(256)
    Size = GetModuleBaseName(GetCurrentProcess(), GetModuleHandle(API_NULL), GetExeBaseName, 256)
    GetExeBaseName = Left$(GetExeBaseName, Size)
End Function

Public Function GetExeFileName() As String
    Dim Size As Long

    GetExeFileName = Space$(256)
    Size = GetModuleFileNameEx(GetCurrentProcess(), GetModuleHandle(API_NULL), GetExeFileName, 256)
    GetExeFileName = Left$(GetExeFileName, Size)
End Function





