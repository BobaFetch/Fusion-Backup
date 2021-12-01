VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaSCp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cost Information (Report)"
   ClientHeight    =   2880
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   17
      Top             =   840
      Width           =   2775
   End
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   4560
      Picture         =   "diaSCp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Show BOM Structure"
      Top             =   840
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   7
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
      PictureUp       =   "diaSCp01a.frx":0342
      PictureDn       =   "diaSCp01a.frx":0488
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2880
      FormDesignWidth =   7260
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   13
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSCp01a.frx":05CE
      PictureDn       =   "diaSCp01a.frx":0714
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions?"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions?"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   9
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1665
   End
End
Attribute VB_Name = "diaSCp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*********************************************************************************
' diaSCp01a - Cost Detail By Part
'
' Notes:
'
' Created: 12/06/02 (nth)
' Revisions:
'   10/22/03 (nth) added get custom report
'
'*********************************************************************************

Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdVew_Click()
   optVew.Value = vbChecked
   ViewParts.Show
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      If Len(cUR.CurrentPart) Then cmbPrt = cUR.CurrentPart
      bOnLoad = False
      FillPartCombo cmbPrt
      cmbPrt = "ALL"
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim i%
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaSCp01a = Nothing
End Sub

Private Sub PrintReport()
   Dim sPart As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   If cmbPrt <> "ALL" Then
      sPart = Compress(cmbPrt)
   Else
      sPart = ""
   End If
   On Error GoTo DiaErr1
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "ShowExDesc"
    aFormulaName.Add "ShowDesc"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Includes Part Numbers " & CStr(cmbPrt) & "...'")
    aFormulaValue.Add optExt.Value
    aFormulaValue.Add optDsc.Value
    
    sCustomReport = GetCustomReport("finpc01.rpt")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
   sSql = "{PartTable.PARTREF} LIKE '" & sPart & "*' "
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
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

Private Sub PrintReport1()
   Dim sPart As String
   
   MouseCursor 13
   If cmbPrt <> "ALL" Then
      sPart = Compress(cmbPrt)
   Else
      sPart = ""
   End If
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='Includes Part Numbers " & cmbPrt & "...'"
   MdiSect.crw.ReportFileName = sReportPath & "finpc01.rpt"
   sSql = "{PartTable.PARTREF} LIKE '" & sPart & "*' "
   If optExt.Value = vbChecked Then
      MdiSect.crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
   Else
      MdiSect.crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
   End If
   If optDsc.Value = vbChecked Then
      MdiSect.crw.SectionFormat(1) = "GROUPHDR.1.0;T;;;"
   Else
      MdiSect.crw.SectionFormat(1) = "GROUPHDR.1.0;F;;;"
   End If
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
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
   cmbPrt = "ALL"
   
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optDsc.Value)) & Trim(str(optExt.Value))
   SaveSetting "Esi2000", "EsiFina", "pc01", sOptions
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", "pc01", sOptions)
   If Len(sOptions) Then
      optDsc = Left(sOptions, 1)
      optExt = Right(sOptions, 1)
   Else
      optDsc.Value = vbUnchecked
      optExt.Value = vbUnchecked
   End If
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then
      ' Part search is closing refresh form
      cmbPrt_LostFocus
   End If
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub cmbPrt_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   
End Sub
