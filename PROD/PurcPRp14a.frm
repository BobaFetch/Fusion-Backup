VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp14a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order Requested By"
   ClientHeight    =   3345
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3345
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp14a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cmbReq 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "2"
      ToolTipText     =   "Contains Previous Table Entries Including Blanks"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   735
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
         Picture         =   "PurcPRp14a.frx":07AE
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
         Picture         =   "PurcPRp14a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   3480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3345
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Dates From"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   1692
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   252
      Index           =   5
      Left            =   2880
      TabIndex        =   14
      Top             =   1560
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   9
      Left            =   5400
      TabIndex        =   13
      Top             =   1560
      Width           =   1404
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Detail"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5400
      TabIndex        =   11
      Top             =   1080
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Tag             =   " "
      Top             =   2040
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requested By"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "PurcPRp14a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbReq_LostFocus()
   cmbReq = CheckLen(cmbReq, 20)
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      'OpenWebHelp "hs907"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillReqBy
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub FillReqBy()
   On Error GoTo DiaErr1
   cmbReq.Clear
   sSql = "SELECT DISTINCT POREQBY FROM PohdTable ORDER BY POREQBY"
   LoadComboBox cmbReq, -1
   cmbReq = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillreqby"
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
   Set PurcPRp14a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim sBegDte As String
    Dim sEndDte As String
    Dim sReqBy As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If Not IsDate(txtBeg) Then
      sBegDte = "1995,01,01"
   Else
      sBegDte = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDte = "2024,12,31"
   Else
      sEndDte = Format(txtEnd, "yyyy,mm,dd")
   End If
   If cmbReq <> "ALL" Then sReqBy = cmbReq

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDetails"
    aFormulaName.Add "ShowDescription"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbReq) & "...'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optDet.Value
    aFormulaValue.Add optDsc.Value
    
    sCustomReport = GetCustomReport("prdpr17")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{PohdTable.POREQBY} LIKE '" & sReqBy & "*' AND " _
          & "{PohdTable.PODATE} in Date(" & sBegDte & ") to Date(" & sEndDte & ")"
   sSql = sSql & " AND {PohdTable.POCAN} = .000 and {PoitTable.PITYPE} <> 16"
'   If optDet.value = vbChecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.1.1;T;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.0;T;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPFTR.1.0;T;;;"
'      If optDsc.value = vbChecked Then
'         MDISect.Crw.SectionFormat(3) = "DETAIL.0.1;T;;;"
'      Else
'         MDISect.Crw.SectionFormat(3) = "DETAIL.0.1;F;;;"
'      End If
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.1.1;F;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.0;F;;;"
'      MDISect.Crw.SectionFormat(2) = "DETAIL.0.1;F;;;"
'      MDISect.Crw.SectionFormat(3) = "GROUPFTR.1.0;F;;;"
'   End If
   
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

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
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 4)
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = Trim(str(optDet.Value)) & Trim(str(optDsc.Value))
   SaveSetting "Esi2000", "EsiProd", "prp14a", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "prp14a", sOptions)
   If Trim(sOptions) <> "" Then
      optDet.Value = Val(Left(sOptions, 1))
      optDsc.Value = Val(Right(sOptions, 1))
   Else
      optDet.Value = vbChecked
      optDsc.Value = vbChecked
   End If
   If optDet.Value = vbChecked Then optDsc.Enabled = True _
                     Else optDsc.Enabled = False
   
End Sub

Private Sub optDet_Click()
   If optDet.Value = vbChecked Then optDsc.Enabled = True _
                     Else optDsc.Enabled = False
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub
