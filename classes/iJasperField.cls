VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iJasperField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Set ReportElement(poElement As ReportElement)

End Property

Public Property Get ReportElement() As ReportElement

End Property

Public Property Let Expression(psExpression As String)

End Property

Public Property Get Expression() As String

End Property

Public Property Get JasperElement() As String

End Property
