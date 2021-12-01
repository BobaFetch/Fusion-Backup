VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPhu15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Time Type Codes Report"
   ClientHeight    =   2160
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2160
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   960
      Width           =   780
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaPhu15.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaPhu15.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaPhu15.frx":0308
      PictureDn       =   "diaPhu15.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2160
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Codes(s)"
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   6
      Top             =   960
      Width           =   1185
   End
End
Attribute VB_Name = "diaPhu15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 3)
   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs907"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT TYPECODE FROM TmcdTable "
   LoadComboBox cmbCde, -1
   cmbCde = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaPhu15 = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCode As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   If cmbCde <> "ALL" Then sCode = Left(cmbCde, 2) Else sCode = ""
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaValue.Add CStr("'Requested By: " & sFacility & "'")
   aFormulaValue.Add CStr("'Includes " & cmbCde & "...'")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admhu15")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{TmcdTable.TYPECODE} Like '" & sCode & "*' "
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = cmbCde
   SaveSetting "Esi2000", "EsiAdmn", "hu15", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiAdmn", "hu15", sOptions)
   cmbCde = sOptions
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub

Private Sub optDis_Click()
   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   PrintReport
   
End Sub
