VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JasperFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oReportElement As ReportElement
Dim sExpression As String
Dim aTemplate() As String
Implements iJasperField

Private Sub Class_Initialize()
    Set oReportElement = New ReportElement
End Sub

Private Property Let iJasperField_Expression(RHS As String)
    sExpression = RHS
End Property

Private Property Get iJasperField_Expression() As String
    iJasperField_Expression = sExpression
End Property

Private Property Get iJasperField_JasperElement() As String
    iJasperField_JasperElement = "<formula><![CDATA[" & escapeHTML(oReportElement.PrintCondition & " = " & Replace(sExpression, vbCrLf, " ## ")) & "]]></formula>"
End Property

Private Property Get iJasperField_ReportElement() As ReportElement
    Set iJasperField_ReportElement = oReportElement
End Property

Private Property Set iJasperField_ReportElement(RHS As ReportElement)
    Set oReportElement = RHS
End Property

Private Function escapeHTML(psString As String) As String
    escapeHTML = Replace(Replace(Replace(psString, "<", "&lt"), ">", "&gt"), """", "'")
End Function
