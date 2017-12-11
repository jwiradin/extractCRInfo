VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JasperFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim aFrame() As String
Dim sName As String
Dim sExpression As String
Dim oCol As Collection
Dim eFrameType As FrameType
Dim oIndex As ListView

Private Sub Class_Initialize()
    aFrame = Split("<band height='23'><printWhenExpression><![CDATA[,,]]></printWhenExpression><frame><reportElement mode='Transparent' x='0' y='0' width='800' height='16' />,,</frame></band>", ",")
    Set oCol = New Collection
    Set oIndex = New ListView
End Sub

Public Property Let FrameType(peFrameType As FrameType)
    eFrameType = peFrameType
End Property
Public Property Get FrameType() As FrameType
    FrameType = eFrameType
End Property


Public Property Let Name(psName As String)
    sName = psName
End Property
Public Property Get Name() As String
    Name = sName
End Property

Public Property Let Expression(psExpression As String)
    sExpression = psExpression
End Property
Public Property Get Expression() As String
    Expression = sExpression
End Property


Public Sub AddField(oField As iJasperField)
    oCol.Add oField, "K" & CStr(CLng((oField.ReportElement.Position.Y * CLng(1000))) + (oField.ReportElement.Position.X))
    oIndex.ListItems.Add Key:="K" & CStr(CLng((oField.ReportElement.Position.Y * CLng(1000))) + (oField.ReportElement.Position.X)), Text:=Right("0000" & CStr(oField.ReportElement.Position.X), 4)
End Sub

Public Function GetJasperElement() As String
    Dim oField As iJasperField

    Dim oListItem As ListItem
    Dim oFrame As JasperFrame
    Dim sresult As String
    
    sresult = ""
    If oCol.Count > 0 Then
        oIndex.SortKey = 0
        oIndex.SortOrder = lvwAscending
        oIndex.Sorted = True
            
        For Each oListItem In oIndex.ListItems
            sresult = sresult & oCol(oListItem.Key).JasperElement()
        Next

        
        aFrame(1) = sExpression
        aFrame(3) = sresult
        GetJasperElement = Join(aFrame, "")
    Else
        GetJasperElement = sresult
    End If
End Function
