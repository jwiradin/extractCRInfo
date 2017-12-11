VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JasperGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim aGroup() As String
Dim sName As String
Dim sExpression As String
Dim sField As String
Dim bStartNewPage As Boolean

Dim oCol As Collection
Private Enum DATAPOS
    GroupName = 1
    GroupStartNewPage
    GroupExpression
    GroupBAND
End Enum

Private Sub Class_Initialize()
    aGroup = Split("<group name=',,',,><groupExpression><![CDATA[$F{,,}]]></groupExpression>,,</group>", ",")
    Set oCol = New Collection
End Sub

Public Property Let Expression(psExpression As String)
    sExpression = psExpression
End Property

Public Property Get Expression() As String
    Expression = sExpression
End Property

Public Property Let IsStartNewPage(pbValue As Boolean)
    bStartNewPage = pbValue
End Property

Public Property Get IsStartNewPage() As Boolean
    IsStartNewPage = bStartNewPage
End Property

Public Sub AddFrame(oFrame As JasperFrame)
    oCol.Add oFrame, oFrame.Name
End Sub

Public Property Let Name(psName As String)
    sName = psName
End Property

Public Property Get Name() As String
    Name = sName
End Property

Public Function GetJasperElement() As String
    Dim oFrame As JasperFrame
    Dim sresult As String
    Dim sFooter As String
    Dim sHeader As String
    Dim sDetail As String
    Dim sTemp As String
    sresult = ""

    For Each oFrame In oCol
        sTemp = oFrame.GetJasperElement()
        If sTemp <> "" Then
            Select Case oFrame.FrameType
            Case HEADER
                sHeader = sHeader & sTemp & vbCrLf
            Case FOOTER
                sFooter = sFooter & sTemp & vbCrLf
            Case DETAIL
                sDetail = sDetail & sTemp & vbCrLf
            End Select
        End If
    Next

    If sHeader <> "" Then
        sresult = sresult & "<groupHeader>" & sHeader & "</groupHeader>"
    End If
    If sFooter <> "" Then
        sresult = sresult & "<groupFooter>" & sFooter & "</groupFooter>"
    End If

    sresult = sresult & sDetail
    
    aGroup(DATAPOS.GroupName * 2 - 1) = sName
    aGroup(DATAPOS.GroupExpression * 2 - 1) = sExpression
    aGroup(DATAPOS.GroupBAND * 2 - 1) = sresult
    aGroup(DATAPOS.GroupStartNewPage * 2 - 1) = IIf(bStartNewPage, " isStartNewPage='true' ", "")
    'GetJasperElement = Replace(Join(aGroup, ""), "'", """")
    GetJasperElement = Join(aGroup, "")
End Function
