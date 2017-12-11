VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JasperReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const WU_LOGPIXELSX = 88
Private Const WU_LOGPIXELSY = 90

Dim oFormulas As Collection
Dim oGroups As Collection
Dim sRpt As String
Dim sName As String
Dim oWriter As StringBuilder

Private Sub Class_Initialize()
    Set oFormulas = New Collection
    Set oGroups = New Collection
    Set oWriter = New StringBuilder
    oWriter.Init
End Sub

Public Sub Init(psRpt As String)
    sRpt = psRpt
End Sub


Public Function ExtractInfo() As String
    Dim oApp As CRAXDRT.Application
    Dim oRpt As CRAXDRT.Report
    Dim oArea As CRAXDRT.Area

    Dim oGroup As JasperGroup
    Dim oFormula As iJasperField

    Set oApp = New Application

    Set oRpt = oApp.OpenReport(sRpt)

    ExtractFormula oRpt

    For Each oArea In oRpt.Areas
        Set oGroup = New JasperGroup
        oGroup.Name = oArea.Name
        oGroup.IsStartNewPage = oArea.NewPageAfter
        If oArea.Kind = crGroupHeader Or oArea.Kind = crGroupFooter Then
            Select Case oArea.GroupConditionField.Kind
                Case crFormulaField
                    oGroup.Expression = oArea.GroupConditionField.FormulaFieldName
                Case crDatabaseField
                    oGroup.Expression = oArea.GroupConditionField.DatabaseFieldName
            End Select
        End If

        printSection oGroup, oArea.Sections, oArea.Kind

        oGroups.Add oGroup, oGroup.Name
    Next

    Set oRpt = Nothing
    Set oApp = Nothing
    
    oWriter.WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
    oWriter.WriteLine "<report>"

    oWriter.WriteLine "<formulas>"
    For Each oFormula In oFormulas
        oWriter.WriteLine oFormula.JasperElement
    Next
    oWriter.WriteLine "</formulas>"


    oWriter.WriteLine "<groups>"
    For Each oGroup In oGroups
        oWriter.WriteLine oGroup.GetJasperElement
    Next
    
    oWriter.WriteLine "</groups>"
    oWriter.WriteLine "</report>"

    ExtractInfo = oWriter.GetString
End Function
Private Sub ExtractFormula(poRpt As CRAXDRT.Report)
    Dim oFormula As CRAXDRT.FormulaFieldDefinition
    Dim oJasperFormula As iJasperField

    For Each oFormula In poRpt.FormulaFields
        Set oJasperFormula = New JasperFormula
        oJasperFormula.Expression = oFormula.Text
        oJasperFormula.ReportElement.PrintCondition = oFormula.Name
        oFormulas.Add oJasperFormula, oFormula.Name
    Next

End Sub

Private Sub printSection(poGroup As JasperGroup, oSections As Sections, poAreaKind As CRAreaKind)
    Dim oSection As CRAXDRT.Section
    Dim oReportObject
    Dim oText As CRAXDRT.TextObject
    Dim oField As CRAXDRT.FieldObject
    Dim oDBField As CRAXDRT.DatabaseFieldDefinition
    Dim oFrame As JasperFrame
    Dim oJasperStatic As iJasperField
    Dim oJasperText As iJasperField

    If oSections.Count > 0 Then

        For Each oSection In oSections
            
            Set oFrame = New JasperFrame
            oFrame.Name = oSection.Name
            oFrame.FrameType = IIf(poAreaKind = crGroupFooter, FrameType.FOOTER, IIf(poAreaKind = crGroupHeader, FrameType.HEADER, FrameType.DETAIL))
            oFrame.Expression = Replace(oSection.ConditionFormula("Suppress"), vbCrLf, "##")
            
            For Each oReportObject In oSection.ReportObjects
                Select Case oReportObject.Kind
                Case crFieldObject
                    Set oField = oReportObject
                    If oField.Field.Kind = crDatabaseField Then
                        Set oDBField = oField.Field
                        Set oJasperText = New TextField
                        
                        With oJasperText
                            .Expression = oDBField.DatabaseFieldName
                            .ReportElement.Height = ConvertTwipsToPixels(oField.Height, 1)
                            .ReportElement.Width = ConvertTwipsToPixels(oField.Width, 0)
                            .ReportElement.Position.X = ConvertTwipsToPixels(oField.Left, 0)
                            .ReportElement.Position.Y = ConvertTwipsToPixels(oField.Top, 1)
                        End With
                        oFrame.AddField oJasperText
                    Else
                        If oField.Field.Kind = crFormulaField Then
                            Set oJasperText = New TextField
                            With oJasperText
                                .Expression = oField.Field.Name
                                .ReportElement.Height = ConvertTwipsToPixels(oField.Height, 1)
                                .ReportElement.Width = ConvertTwipsToPixels(oField.Width, 0)
                                .ReportElement.Position.X = ConvertTwipsToPixels(oField.Left, 0)
                                .ReportElement.Position.Y = ConvertTwipsToPixels(oField.Top, 1)
                            End With
                            oFrame.AddField oJasperText
                        End If
                    End If
                Case crTextObject
                    Set oText = oReportObject
                    Set oJasperStatic = New StaticField
                    With oJasperStatic
                        .Expression = oText.Text
                        .ReportElement.Height = ConvertTwipsToPixels(oText.Height, 1)
                        .ReportElement.Width = ConvertTwipsToPixels(oText.Width, 0)
                        .ReportElement.Position.X = ConvertTwipsToPixels(oText.Left, 0)
                        .ReportElement.Position.Y = ConvertTwipsToPixels(oText.Top, 1)
                    End With
                    oFrame.AddField oJasperStatic
                End Select
            Next
            poGroup.AddFrame oFrame
        Next
    End If
End Sub

Private Function ConvertTwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
    'Handle to device
    Dim lngDC As Long
    Dim lngPixelsPerInch As Long
    Const nTwipsPerInch = 1440
    lngDC = GetDC(0)
    If (lngDirection = 0) Then    'Horizontal
        lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
    Else    'Vertical
        lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
    End If
    lngDC = ReleaseDC(0, lngDC)
    ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch * 0.75
End Function

