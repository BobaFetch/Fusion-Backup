VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchasing History By Manufacturing Order"
   ClientHeight    =   3300
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3300
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp09a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRun 
      Height          =   288
      Left            =   6040
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Associated PO Runs"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   2420
      Width           =   850
   End
   Begin VB.CheckBox optItm 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2680
      Width           =   850
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PurcPRp09a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PurcPRp09a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Runs From PO Items"
      Top             =   1200
      Width           =   3165
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3300
      FormDesignWidth =   7215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   5460
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2420
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2680
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2265
   End
End
Attribute VB_Name = "PurcPRp09a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   GetRuns
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc, True)
   If lblDsc.ForeColor <> ES_RED Then
      GetRuns
   Else
      cmbRun.Clear
   End If
   
End Sub

Private Sub cmbRun_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   If cmbRun.ListCount > 0 Then
      For iList = 0 To cmbRun.ListCount - 1
         If cmbRun.List(iList) = cmbRun Then b = 1
      Next
      If b = 0 Then cmbRun = cmbRun.List(0)
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PIRUNPART,PARTREF,PARTNUM " _
          & "FROM PoitTable,PartTable WHERE PIRUNPART=PARTREF " _
          & "ORDER BY PIRUNPART"
   LoadComboBox cmbPrt, 1
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
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
   Set PurcPRp09a = Nothing
   
End Sub

Private Sub PrintReport()
    MouseCursor 13
    Dim lRunno As Long
    Dim sPartNumber As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   
   sPartNumber = Compress(cmbPrt)
   lRunno = Val(cmbRun)
   On Error GoTo DiaErr1
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowExtendedDescription"
    aFormulaName.Add "ShowItem"
    aFormulaName.Add "ShowDescription"
    
    aFormulaValue.Add CStr("'CompanyName" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optExt.Value
    aFormulaValue.Add optItm.Value
    aFormulaValue.Add optDsc.Value
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("prdpr11")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{RunsTable.RUNREF}='" & sPartNumber & "' " _
          & "AND {RunsTable.RUNNO}=" & lRunno & " AND {PoitTable.PITYPE}<>16"
   
'   If optExt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.1;T;;;"
'   End If
'
'   If optItm.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(2) = "DETAIL.0.1;F;;;"
'      MDISect.Crw.SectionFormat(3) = "DETAIL.0.2;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(2) = "DETAIL.0.1;T;;;"
'      MDISect.Crw.SectionFormat(3) = "DETAIL.0.2;T;;;"
'   End If
'
'   If optDsc.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(4) = "GROUPHDR.2.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(4) = "GROUPHDR.2.0;T;;;"
'   End If
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
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
   sOptions = Trim(str(optDsc.Value)) _
              & Trim(str(optExt.Value)) _
              & Trim(str(optItm.Value))
   SaveSetting "Esi2000", "EsiProd", "pr11", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "pr11", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.Value = Val(Mid(sOptions, 1, 1))
      optExt.Value = Val(Mid(sOptions, 2, 1))
      optItm.Value = Val(Mid(sOptions, 3, 1))
   Else
      optDsc.Value = vbChecked
      optExt.Value = vbChecked
      optItm.Value = vbChecked
   End If
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
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


Private Sub optItm_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub GetRuns()
   Dim sPartNumber As String
   cmbRun.Clear
   On Error GoTo DiaErr1
   sPartNumber = Compress(cmbPrt)
   sSql = "SELECT DISTINCT PIRUNPART,PIRUNNO FROM " _
          & "PoitTable WHERE PIRUNPART='" & sPartNumber & "'"
   LoadNumComboBox cmbRun, "####0", 1
   If cmbRun.ListCount > 0 Then cmbRun = cmbRun.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
