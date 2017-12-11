VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extract Crystal Report Info"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6315
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   360
      Left            =   8370
      TabIndex        =   2
      Top             =   105
      Width           =   750
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   255
      TabIndex        =   1
      Top             =   135
      Width           =   7980
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Extract"
      Height          =   360
      Left            =   7770
      TabIndex        =   0
      Top             =   555
      Width           =   1350
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPath As Path
Private Sub cmdExecute_Click()
    Dim oJasper As JasperReport
    Dim oWriter As FileWriter
    
    Set oJasper = New JasperReport
    Set oWriter = New FileWriter
    Set oPath = New Path
    
    oJasper.Init txtFileName.Text
    
    oWriter.Init App.Path & "\" & oPath.GetFileNameOnly(txtFileName.Text) & ".xml"
    oWriter.WriteLine oJasper.ExtractInfo()
    oWriter.CloseFile
    Set oWriter = Nothing
    Set oPath = Nothing
    MsgBox "done"
End Sub

Private Sub Command2_Click()
    CommonDialog1.ShowOpen
    txtFileName.Text = CommonDialog1.FileName
End Sub

