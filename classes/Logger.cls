VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iCurrentMax As Integer
Private sPath As String
Private aPattern() As String

Private sPattern As String
Private MAXLENGTH As Long
Private MAXCOUNT As Long

Private iCounterLocation As Integer 'Counter location in file name
Private iDateLocation As Integer 'Date format location in file name
Private aValidDateFormat() As String

Private bDebug As Boolean
Private bError As Boolean
Private bWarning As Boolean

Private sMsgType As String

Private Const SEPARATOR As String = " "
Dim sEXEPath As String
Dim sEXEName As String

Public Property Get IsDebug() As Boolean
    IsDebug = bDebug
End Property

Public Property Get IsWarning() As Boolean
    IsWarning = bWarning
End Property

Public Property Get IsError() As Boolean
    IsError = bError
End Property

Friend Sub Init()

    Dim sIni As String
    Dim sTemp As String
    
    On Error GoTo Err_Handler
    GetRuntimeVariable sEXEPath, sEXEName
    
    sPattern = sEXEName & "_yyyymmdd_##"
    sPath = Path.Combine(sEXEPath, "Log\")
    bDebug = True
    
    MAXLENGTH = 1 * 1024 * 1024
    
    iCurrentMax = 1
    
    If InStr(sPattern, "##") = 0 Then
        sPattern = sPattern & "_##"
    End If
    
    aPattern = Split(sPattern, "_")
    
    'Find the location of the counter parameter and date format
    Dim iCnt As Integer
    Dim sDt As Variant
    
    For iCnt = 0 To UBound(aPattern)
        sTemp = aPattern(iCnt)
        
        If InStr(Trim$(sTemp), "##") > 0 Then
            iCounterLocation = iCnt
        Else
            If iDateLocation = 0 Then
                For Each sDt In aValidDateFormat
                    If sTemp = CStr(sDt) Then
                        iDateLocation = iCnt
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    CreatePath sPath
    EnumerateLog
    
    Exit Sub
Err_Handler:
    Err.Raise vbObjectError, "clsLog.Init", FormatError("", Err)
End Sub

Private Sub WriteError(psModule As String, poErr As ErrObject)
    WriteLine psModule & " - " & poErr.Source & ":" & CStr(poErr.Number) & "-" & poErr.Description
End Sub

Public Sub LogDebug(psMessage As String)
    If bDebug Then
        sMsgType = "[DEBUG]"
        WriteLine psMessage
    End If
End Sub

Public Sub LogWarning(psMessage As String)
    If bWarning Then
        sMsgType = "[WARNING]"
        WriteLine psMessage
    End If
End Sub

Public Sub LogError(psMessage As String)
    If bError Then
        sMsgType = "[ERROR]"
        WriteLine psMessage
    End If
End Sub

Private Sub WriteLine(psMessage As String)
    Dim iCnt As Integer
    Dim sNew As String
    Dim sPrv As String
    
    On Error GoTo Err_Handler
    
    If IsRollOver() Then
        For iCnt = iCurrentMax To 1 Step -1
            If iCnt = MAXCOUNT Then
                Kill GenerateFileName(iCnt)
            Else
                'rename me to next count
                Name GenerateFileName(iCnt) As GenerateFileName(iCnt + 1)
            End If
        Next
        
        iCurrentMax = IIf(iCurrentMax < MAXCOUNT, iCurrentMax + 1, MAXCOUNT)
    End If
        
    Open GenerateFileName(1) For Append As #1
    Print #1, FormatMessage(psMessage)
    Close #1

    Exit Sub
Err_Handler:
    Err.Raise vbObjectError, "clsLog.WriteLine", FormatError("", Err)
End Sub

Private Function FormatMessage(psMessage As String) As String
    FormatMessage = Join(Array(Format(Now, "yyyy-mm-dd hh:nn:ss"), sMsgType, "[" & sEXEName & "]", "->", psMessage), SEPARATOR)
End Function

Private Function FormatError(psModule As String, poErr As ErrObject) As String
    FormatError = poErr.Source & ":" & CStr(poErr.Number) & "-" & poErr.Description
End Function

Private Function IsRollOver() As Boolean
    IsRollOver = False
    If Dir(GenerateFileName(1)) <> "" Then
        IsRollOver = (FileLen(GenerateFileName(1)) >= MAXLENGTH)
    End If
End Function

Private Function GenerateFileName(piCnt As Integer) As String
    GenerateFileName = Replace(sPattern, aPattern(iDateLocation), Format(Now(), aPattern(iDateLocation)))
    GenerateFileName = Replace(GenerateFileName, aPattern(iCounterLocation), Format(piCnt, String$(Len(aPattern(iCounterLocation)), "0")))
    GenerateFileName = Path.Combine(sPath, GenerateFileName & ".log")
End Function

Private Sub GetRuntimeVariable(ByRef psPath As String, ByRef psExeName As String)
    If UCase(Path.GetExeBaseName) = "VB6.EXE" Then
        psExeName = App.EXEName
        psPath = App.Path
    Else
        psExeName = Path.GetExeBaseName
        psPath = Replace(Path.GetExeFileName, psExeName, "")
    End If
End Sub

Private Sub EnumerateLog()
    
    Dim sMask As String
    Dim iCnt As Integer
    Dim sFile As String
    Dim sFileName As String
    Dim aTmp() As String
    
    sMask = ""
    
    For iCnt = LBound(aPattern) To UBound(aPattern)
        Select Case True
            Case iCnt = iDateLocation
                sMask = sMask & "_" & Format(Now(), aPattern(iCnt))
            Case iCnt = iCounterLocation
                sMask = sMask & "_" & "*"
            Case Else
                sMask = sMask & IIf(sMask <> "", "_", "") & aPattern(iCnt)
        End Select
    Next
    
    sMask = sMask & ".log"
    sFile = Dir(Path.Combine(sPath, sMask))
    
    Do While sFile <> ""
        sFileName = Path.GetFileNameOnly(sFile)
        aTmp = Split(sFileName, "_")
        
        If IsNumeric(aTmp(UBound(aTmp))) Then
            iCurrentMax = IIf(iCurrentMax < CInt(aTmp(UBound(aTmp))), CInt(aTmp(UBound(aTmp))), iCurrentMax)
        End If
        
        sFile = Dir
    Loop
    
End Sub

Private Sub CreatePath(psPath As String)
    Dim aPath() As String
    Dim sPath As String
    Dim iCnt As Integer
    
    aPath = Split(psPath, "\")
    sPath = ""
    
    For iCnt = LBound(aPath) To UBound(aPath)
        If aPath(iCnt) <> "" Then
            sPath = sPath & aPath(iCnt) & "\"
            
            If Dir(sPath, vbDirectory) = "" Then
                MkDir sPath
            End If
        End If
    Next
End Sub

Private Sub Class_Initialize()
    ' Init valid date formats
    iCounterLocation = 0
    MAXCOUNT = 10
    MAXLENGTH = CLng(1024) * 1024
    
    bDebug = True
    bWarning = True
    bError = True
    
    aValidDateFormat = Split("yyyymmdd,ddmmyyyy,mmddyyyy", ",")
End Sub



