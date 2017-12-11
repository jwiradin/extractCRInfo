VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oReportElement As ReportElement
Dim sExpression As String
Dim aTemplate() As String
Implements iJasperField

Private Enum DATAPOS
    X = 1
    Y
    W
    H
    Condition
    Field
End Enum
Private Sub Class_Initialize()
    aTemplate = Split("<textField><reportElement positionType='Float' x=',,' y=',,' width=',,' height=',,'><printWhenExpression><![CDATA[,,]]></printWhenExpression></reportElement><textFieldExpression><![CDATA[$F{,,}]]></textFieldExpression></textField>", ",")
    Set oReportElement = New ReportElement
End Sub

Private Property Let iJasperField_Expression(RHS As String)
    sExpression = RHS
End Property

Private Property Get iJasperField_Expression() As String
    iJasperField_Expression = sExpression
End Property

Private Property Get iJasperField_JasperElement() As String
    iJasperField_JasperElement = Convert()
End Property

Private Property Get iJasperField_ReportElement() As ReportElement
    Set iJasperField_ReportElement = oReportElement
End Property

Private Property Set iJasperField_ReportElement(RHS As ReportElement)
    Set oReportElement = RHS
End Property

Private Function Convert() As String
    aTemplate(DATAPOS.X) = oReportElement.Position.X
    aTemplate(DATAPOS.Y * 2 - 1) = oReportElement.Position.Y
    aTemplate(DATAPOS.W * 2 - 1) = oReportElement.Width
    aTemplate(DATAPOS.H * 2 - 1) = oReportElement.Height
    aTemplate(DATAPOS.Condition * 2 - 1) = oReportElement.PrintCondition
    aTemplate(DATAPOS.Field * 2 - 1) = sExpression
    
    'Convert = Replace(Join(aTemplate, ""), "'", """")
    Convert = Join(aTemplate, "")
End Function

