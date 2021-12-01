VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ShopSHp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Orders"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkShowIntCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   13
      Top             =   4320
      Width           =   735
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   5520
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   1080
      Width           =   1140
   End
   Begin VB.CheckBox chkDocList 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   3060
      TabIndex        =   12
      Top             =   4080
      Width           =   735
   End
   Begin VB.CheckBox chkShowToolList 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   11
      Top             =   3840
      Width           =   735
   End
   Begin VB.CheckBox chkBOM 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   8
      ToolTipText     =   "Document List (Printed MO's Only)"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CheckBox chkTime 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "ShopSHp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp01a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Revision-Select From List"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox chkPickList 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   9
      ToolTipText     =   "Pick List For This Part (Printed MO's Only) Status PL"
      Top             =   3360
      Width           =   726
   End
   Begin VB.CheckBox chkDoc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   7
      ToolTipText     =   "Document List (Printed MO's Only)"
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox chkBudget 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   10
      Top             =   3600
      Width           =   735
   End
   Begin VB.CheckBox optFrom 
      Height          =   255
      Left            =   3960
      TabIndex        =   42
      Top             =   5220
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6240
      TabIndex        =   39
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   320
         Left            =   600
         Picture         =   "ShopSHp01a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
      Begin VB.CommandButton optDis 
         Height          =   320
         Left            =   0
         Picture         =   "ShopSHp01a.frx":0AC2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   18
      Top             =   4860
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   23
      Top             =   5700
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   22
      Top             =   5460
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   9
      Left            =   3960
      TabIndex        =   21
      Top             =   5820
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   20
      Top             =   5220
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   19
      Top             =   4980
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkSoAlloc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   24
      Top             =   5580
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkSvcs 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.CheckBox chkComments 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   14
      Top             =   4740
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   1080
      Width           =   3545
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   4860
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4605
      FormDesignWidth =   7485
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Internal Comments"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   55
      ToolTipText     =   "Show Service Document List"
      Top             =   4320
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Service Part Document List"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   54
      ToolTipText     =   "Show Service Document List"
      Top             =   4080
      Width           =   2865
   End
   Begin VB.Label Label1 
      Caption         =   "Show Tool List"
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show BOM"
      Height          =   255
      Index           =   22
      Left            =   240
      TabIndex        =   52
      ToolTipText     =   "Document List (Printed MO's Only)"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Operation Time"
      Height          =   255
      Index           =   21
      Left            =   240
      TabIndex        =   51
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PL Rev"
      Height          =   255
      Index           =   18
      Left            =   5160
      TabIndex        =   48
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   47
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblQty 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   46
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Pick List"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   45
      ToolTipText     =   "Pick List For This Part (Printed MO's Only) Status PL"
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lblSta 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6840
      TabIndex        =   44
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show MO Budget"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   43
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label lblTyp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   41
      Top             =   1440
      Width           =   300
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type/Status"
      Height          =   255
      Index           =   15
      Left            =   5160
      TabIndex        =   40
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Information"
      Enabled         =   0   'False
      Height          =   255
      Index           =   14
      Left            =   3480
      TabIndex        =   38
      Top             =   4860
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspection Document"
      Enabled         =   0   'False
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   37
      Top             =   5700
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Code"
      Enabled         =   0   'False
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   36
      Top             =   5460
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Allocations"
      Enabled         =   0   'False
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   35
      Top             =   5220
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Comments"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   34
      Top             =   4980
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Document List"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   33
      ToolTipText     =   "Document List (Printed MO's Only)"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show SO Allocations"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   32
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Allocations"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   31
      Top             =   5580
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Outside Service Part Numbers"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   30
      Top             =   2400
      Width           =   2715
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Operation Comments"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   29
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   28
      Top             =   4740
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   27
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   26
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   25
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "ShopSHp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/13/02 Added PKRECORD for new index
' 3/25/04 Removed Jet tables and reorged prdshcvr.rpt
' 3/28/05 Revamped the Cover Sheet and formatting.
Option Explicit
Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

'Dim DbDoc   As Recordset 'Jet
'Dim DbPls   As Recordset 'Jet
Dim bPrinting As Boolean

Dim bGoodPart As Byte
Dim bGoodMo As Byte
Dim bOnLoad As Byte
Dim bTablesCreated As Byte
Dim bUserTypedRun As Byte

Dim sBomRev As String
Dim sRunPkstart As String
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetRevisions()
   cmbRev.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREV FROM BmhdTable WHERE BMHREF='" _
          & Compress(cmbPrt) & "' ORDER BY BMHREV"
   LoadComboBox cmbRev, -1
   Exit Sub
   
DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh01", "000000000")
'   If Len(sOptions) > 0 Then
'      For iList = 1 To 5
'         optInc(iList) = Val(Mid(sOptions, iList, 1))
'      Next
'      For iList = 7 To 11
'         optInc(iList) = Val(Mid(sOptions, iList, 1))
'      Next
'      chkBudget = Val(Mid(sOptions, iList, 1))
'   End If

   chkComments = Mid(sOptions, 1, 1)
   chkTime = Mid(sOptions, 2, 1)
   chkSvcs = Mid(sOptions, 3, 1)
   chkSoAlloc = Mid(sOptions, 4, 1)
   chkDoc = Mid(sOptions, 5, 1)
   chkBOM = Mid(sOptions, 6, 1)
   chkPickList = Mid(sOptions, 7, 1)
   chkBudget = Mid(sOptions, 8, 1)
   chkShowToolList = Mid(sOptions, 9, 1)
   chkDocList = Mid(sOptions, 10, 1)
   If Len(sOptions) > 10 Then chkShowIntCmt = Mid(sOptions, 11, 1) Else chkShowIntCmt.Value = 0

   'chkSoAlloc.Value = GetSetting("Esi2000", "EsiProd", "sh01all", chkSoAlloc.Value)
   lblPrinter = GetSetting("Esi2000", "EsiProd", "sh01Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   
   sOptions = CStr(chkComments) & CStr(chkTime) & CStr(chkSvcs) & CStr(chkSoAlloc) _
      & CStr(chkDoc) & CStr(chkBOM) & CStr(chkPickList) & CStr(chkBudget) & CStr(chkShowToolList) & CStr(chkDocList) & CStr(chkShowIntCmt) & "000000000"
   SaveSetting "Esi2000", "EsiProd", "sh01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "sh01all", Trim(chkSoAlloc.Value)
   SaveSetting "Esi2000", "EsiProd", "sh01Printer", lblPrinter
   
End Sub


Private Sub chkShowToolList_KeyPress(KeyAscii As Integer)
    KeyLock KeyAscii
End Sub


Private Sub cmbPrt_Click()
    bUserTypedRun = 0
   bGoodPart = GetRuns()
   If bGoodPart Then GetRevisions
   optFrom.Value = vbUnchecked     'we have changed the default value/mo now
   bPrinting = False
End Sub


Private Sub cmbPrt_LostFocus()
    If (bPrinting = False) Then
       cmbPrt = CheckLen(cmbPrt, 30)
       bGoodPart = GetRuns()
       cmbPrt = UCase(cmbPrt)
       If bGoodPart Then GetRevisions
    End If

End Sub

Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo Psh01
   sProcName = "printreport"
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
'   Dim aRptPara As New Collection
'   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sSubSql As String
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "PartNumber"
   aFormulaName.Add "RunNumber"
   aFormulaName.Add "ShowOpComments"
   aFormulaName.Add "ShowOpTime"
   aFormulaName.Add "ShowSvcParts"
   aFormulaName.Add "ShowSoAllocs"
   aFormulaName.Add "ShowDocList"
   aFormulaName.Add "ShowBOM"
   aFormulaName.Add "ShowPickList"
   aFormulaName.Add "ShowMoBudget"
   aFormulaName.Add "ShowToolList"
   aFormulaName.Add "ShowServPartDoc"
   aFormulaName.Add "ShowInternalCmt"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(sPartNumber)) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(cmbRun)) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkComments) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkTime) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkSvcs) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkSoAlloc) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkDoc) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkBOM) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkPickList) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkBudget) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkShowToolList) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkDocList) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkShowIntCmt) & "'")
   
   sCustomReport = GetCustomReport("prdsh01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RunsTable.RUNNO} = {@Run} and {PartTable.PARTREF} = {@PartNumber}"
   cCRViewer.SetReportSelectionFormula sSql
   
   sSubSql = "{MopkTable.PKMORUN} = {?Pm-RunsTable.RUNNO} and " _
            & "{MopkTable.PKMORUN} = {?Pm-RunsTable.RUNNO} and " _
            & "{MopkTable.PKMOPART} = {?Pm-RunsTable.RUNREF} and  " _
            & "({MopkTable.PKTYPE} = 10 OR {MopkTable.PKTYPE} = 9)"
            ' PKTYPE=10 is picked type and PickOpenItem = 9
   ' set the sub sql variable pass the sub report name
   cCRViewer.SetSubRptSelFormula "custpklist.rpt", sSubSql

bPrinting = True
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   
'    ' set report parameters
'    Dim aRptPara As New Collection
'    Dim aRptParaType As New Collection
'    aRptPara.Add sPartNumber        ' @PartRef
'    aRptParaType.Add CStr("String")
'    cCRViewer.SetReportDBParameters aRptPara, aRptParaType  'must happen AFTER SetDbTableConnection call!

   cCRViewer.OpenCrystalReportObject Me, aFormulaName
      
   'cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   MouseCursor 0
   DoEvents
   bPrinting = False
   Exit Sub
   
Psh01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
Psh02:
   DoModuleErrors Me
   
End Sub


Private Sub cmbRev_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbRev = CheckLen(cmbRev, 4)
   For iList = 0 To cmbRev.ListCount - 1
      If Trim(cmbRev) = Trim(cmbRev.List(iList)) Then b = 1
   Next
   If b = 0 And cmbRev.ListCount > 0 Then
      Beep
      cmbRev = cmbRev.List(0)
   End If
   sBomRev = Trim(cmbRev)
   
End Sub


Private Sub cmbRun_Click()
   GetThisRun
   
End Sub


Private Sub cmbRun_KeyDown(KeyCode As Integer, Shift As Integer)
    bUserTypedRun = 1
End Sub

Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   GetThisRun
   
End Sub




Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4120
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

'Private Sub cmdVew_Click(Index As Integer)
'   Dim sWindows As String
'
'   'if allcoation button pressed, print allocations
'   If Index = 1 Then
'      PrintAllocations
'      Exit Sub
'   End If
'
'   'otherwise, print cover sheet
'   If chkDoc.Value = vbUnchecked And chkPickList.Value = vbUnchecked Then
'      MsgBox "Please Select One Or Both Cover Sheet Options.", _
'         vbInformation, Caption
'      Exit Sub
'   End If
'   MouseCursor 13
'   BuildPickList
'   BuildDocumentList
'
'   PrintCover
'
'End Sub
'
Private Sub Form_Activate()
   If bOnLoad Then
      'CreateJetTables
      FillAllRuns cmbPrt
      If optFrom.Value = vbChecked Then
         cmbPrt = ShopSHe02a.cmbPrt
         cmbRun = ShopSHe02a.cmbRun
      End If
      bGoodPart = GetRuns()
      If bGoodPart Then GetRevisions
      bOnLoad = 0
      bPrinting = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bUserTypedRun = 0
   
   GetOptions
   bTablesCreated = 0
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV,PARUN,RUNREF,RUNSTATUS," _
          & "RUNNO FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF "
   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   rdoQry.Parameters.Append AdoParameter1
   
   bOnLoad = 1
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set rdoQry = Nothing
   If optFrom.Value = vbChecked Then
      ShopSHe02a.lblStat = lblSta
      ShopSHe02a.Show
   Else
      FormUnload
   End If
   Set ShopSHp01a = Nothing
   
End Sub




Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   Dim iOriginalRun As Integer
   Dim bOriginalRunFound As Byte
   
   bOriginalRunFound = 0
   
   On Error GoTo DiaErr1
   iOriginalRun = Val(cmbRun)
   MouseCursor 13
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   rdoQry.Parameters(0).Value = sPartNumber
'   rdoQry(0) = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, rdoQry)
   If bSqlRows Then
      With RdoRns
         If optFrom Then
            cmbRun = ShopSHe02a.cmbRun
         Else
            cmbRun = Format(!Runno, "####0")
         End If
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(!PALEVEL, "#")
         cmbRev = "" & Trim(!PABOMREV)
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            If iOriginalRun = !Runno Then bOriginalRunFound = 1
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then
         If optFrom And bOriginalRunFound Then
            cmbRun = LTrim(str(iOriginalRun))
         Else
            cmbRun = cmbRun.List(cmbRun.ListCount - 1)
         End If
      Else
        If bOriginalRunFound And bUserTypedRun Then cmbRun = LTrim(str(iOriginalRun))
      End If
      GetRuns = True
      GetThisRun
   Else
      sPartNumber = ""
      GetRuns = False
   End If
   MouseCursor 0
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblQty_Click()
   'run qty
   
End Sub

Private Sub chkbudget_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbPrt.SetFocus
      Exit Sub
   Else
      PrintReport
      MouseCursor 0
   End If
   
End Sub

Private Sub chkdoc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFrom_Click()
   'dummy to check if from Revise mo
   
End Sub



Private Sub chkpicklist_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   Dim b As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbPrt.SetFocus
      Exit Sub
   Else
      On Error Resume Next
      
      'Truncate and insert here so that there will always be a hook
'      sSql = "TRUNCATE TABLE EsReportCvrTable"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "INSERT INTO EsReportCvrTable (DLSPart1) VALUES('')"
'      RdoCon.Execute sSql, rdExecDirect
'
      If chkPickList.Value = vbChecked Then
         If lblSta = "SC" Or lblSta = "RL" Then
            cmbRev.Enabled = True
            sMsg = "Do You Want To Print The MO Pick " & vbCr _
                   & "List And Move The Run Status To PL?"
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               'Build the pick list and change status
               b = 1
               MouseCursor 13
               BuildPartsList
            Else
               CancelTrans
            End If
         Else
            cmbRev.Enabled = False
            MouseCursor 13
            b = 1
            BuildPickList
         End If
      End If
      If chkDoc = vbChecked Then
         MouseCursor 13
         b = 1
         'BuildDocumentList
      End If
'      If b = 1 Then PrintCover
      PrintReport
   End If
   
End Sub


Private Sub GetThisRun()
   Dim RdoRun As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNPKSTART,RUNQTY FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbPrt) & "' AND " _
          & "RUNNO=" & cmbRun & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblSta = "" & Trim(!RUNSTATUS)
         If lblSta = "SC" Or lblSta = "RL" Then cmbRev.Enabled = True _
                     Else cmbRev.Enabled = False
         If Not IsNull(!RUNPKSTART) Then
            sRunPkstart = Format(!RUNPKSTART, "mm/dd/yy")
         Else
            sRunPkstart = Format(ES_SYSDATE, "mm/dd/yy")
         End If
         lblQty = Format(!RUNQTY, ES_QuantityDataFormat)
         ClearResultSet RdoRun
      End With
   End If
   Set RdoRun = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


'
'Private Sub BuildDocumentList()
'   Dim RdoDoc As rdoResultset
'   Dim b As Byte
'   Dim iRow As Integer
'   Dim sCol As String * 1
'
'   On Error GoTo DiaErr1
'   '5/06/04 Changed to static table
'   sSql = "SELECT * FROM RndlTable WHERE RUNDLSRUNREF='" & Compress(cmbPrt) & "' " _
'          & "AND RUNDLSRUNNO=" & Val(cmbRun) & " ORDER BY RUNDLSNUM"
'   bSqlRows = GetDataSet(RdoDoc, ES_FORWARD)
'   If bSqlRows Then
'      With RdoDoc
'         Do Until .EOF
'            iRow = iRow + 1
'            Err = 0
'            b = b + 1
'            If b < 6 Then
'               sCol = Trim$(Str(b))
'               sSql = "UPDATE EsReportCvrTable SET " _
'                      & "DLSPart" & sCol & "='" & Trim(!RUNDLSDOCREFLONG) & "'," _
'                      & "DLSDocRev" & sCol & "='" & Trim(!RUNDLSDOCREV) & "'," _
'                      & "DLSSheet" & sCol & "='" & Trim(!RUNDLSDOCREFSHEET) & "'," _
'                      & "DLSClass" & sCol & "='" & Trim(!RUNDLSDOCREFCLASS) & "'," _
'                      & "DLSDocDesc" & sCol & "='" & Trim(!RUNDLSDOCREFDESC) & "'," _
'                      & "DLSDocEco" & sCol & "='" & Trim(!RUNDLSDOCREFECO) & "'," _
'                      & "DLSDocAdcn" & sCol & "='" & Trim(!RUNDLSDOCREFADCN) & "' "
'               RdoCon.Execute sSql, rdExecDirect
'            End If
'            .MoveNext
'         Loop
'         ClearResultSet RdoDoc
'      End With
'   Else
'      sSql = "UPDATE EsReportCvrTable SET DLSDocRef1 = '*** No Documents Recorded ***'"
'      RdoCon.Execute sSql, rdExecDirect
'   End If
'   Set RdoDoc = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "builddoclist"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'SC or RL...no Pick List yet

Private Sub BuildPartsList()
   Dim RdoLst As ADODB.Recordset
   
   Dim b As Byte
   Dim iPkRecord As Integer
   Dim cConversion As Currency
   Dim cQuantity As Currency
   Dim cSetup As Currency
   Dim sUnits As String
   Dim sComt As String
   Dim sCol As String * 1
   
   On Error Resume Next
   
   ' this table does not appear to be used, but better to have no data than someone else's data
   sSql = "TRUNCATE TABLE EsReportPlsTable"
   clsADOCon.ExecuteSql sSql
   
'   sSql = "INSERT INTO EsReportPlsTable (PLSRow,PLSPart) VALUES(1,'')"
'   clsADOCon.ExecuteSql sSql
   
   sSql = "SELECT DISTINCT BMASSYPART FROM BmplTable " _
          & "WHERE BMASSYPART='" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      b = 1
      ClearResultSet RdoLst
   Else
      MouseCursor 0
      b = 0
      MsgBox "This Part Does Not Have A Parts List.", vbExclamation, Caption
   End If
   If b = 1 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALOCATION," _
             & "BMASSYPART,BMPARTREF,BMREV,BMQTYREQD,BMSETUP,BMADDER," _
             & "BMCONVERSION,BMUNITS FROM PartTable,BmplTable WHERE (" _
             & "PARTREF=BMPARTREF AND BMASSYPART='" & Compress(cmbPrt) & "' " _
             & "AND BMREV='" & sBomRev & "') "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_STATIC)
      If bSqlRows Then
         b = 0
         With RdoLst
            clsADOCon.BeginTrans
            clsADOCon.ADOErrNum = 0
            
            Do Until .EOF
               If Not IsNull(!BMSETUP) Then
                  cSetup = !BMSETUP
               Else
                  cSetup = 0
               End If
               sUnits = "" & Trim(!BMUNITS)
               cQuantity = Format(!BMQTYREQD + !BMADDER, ES_QuantityDataFormat)
               cConversion = Format(!BMCONVERSION, "#####0.0000")
               ' Get the Eng Partlist comment
               sComt = GetBMComment(Compress(cmbPrt), sBomRev)
               
               If cConversion = 0 Then cConversion = 1
               cQuantity = cQuantity / cConversion
               cQuantity = (cQuantity * Val(lblQty)) + cSetup
               b = b + 1
               sCol = Trim$(str(b))
'               If b = 1 Then
'                  sSql = "UPDATE EsReportPlsTable SET " _
'                         & "PLSPart='" & Trim(!PartNum) & "'," _
'                         & "PLSDesc='" & Trim(!PADESC) & " '," _
'                         & "PLSPQty=" & cQuantity & "," _
'                         & "PLSUom='" & Trim(!BMUNITS) & "'," _
'                         & "PLSLoc='" & Trim(!PALOCATION) & "' " _
'                         & "WHERE PLSRow=1"
'               Else
'                  sSql = "INSERT INTO EsReportPlsTable " _
'                         & "(PLSRow,PLSPart,PLSDesc,PLSPQty,PLSUom,PLSLoc) Values" _
'                         & "(" & sCol & ",'" _
'                         & Trim(!PartNum) & "','" _
'                         & Trim(!PADESC) & "'," _
'                         & cQuantity & ",'" _
'                         & Trim(!BMUNITS) & "','" _
'                         & Trim(!PALOCATION) & "')"
'               End If
'               clsADOCon.ExecuteSql sSql
               Err.Clear
               clsADOCon.ADOErrNum = 0
               
               If sRunPkstart = "" Then sRunPkstart = Format(ES_SYSDATE, "mm/dd/yy")
               iPkRecord = iPkRecord + 1
               sSql = "INSERT INTO MopkTable (PKPARTREF," _
                      & "PKMOPART,PKMORUN,PKTYPE,PKPDATE," _
                      & "PKPQTY,PKBOMQTY,PKRECORD,PKUNITS,PKCOMT) VALUES('" _
                      & Trim(!PartRef) & "','" & Compress(cmbPrt) & "'," _
                      & cmbRun & ",9,'" & sRunPkstart & "'," & cQuantity _
                      & "," & cQuantity & "," & iPkRecord & ",'" & sUnits & "','" & Trim(sComt) & "') "
               clsADOCon.ExecuteSql sSql
               .MoveNext
            Loop
            ClearResultSet RdoLst
            If clsADOCon.ADOErrNum = 0 Then
               clsADOCon.CommitTrans
               Sleep 500
               sSql = "UPDATE RunsTable SET RUNSTATUS='PL'," _
                      & "RUNPLDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
                      & "WHERE RUNREF='" & Compress(cmbPrt) & "' " _
                      & "AND RUNNO=" & cmbRun & " "
               clsADOCon.ExecuteSql sSql
               lblSta = "PL"
            Else
               clsADOCon.RollbackTrans
               clsADOCon.ADOErrNum = 0
            End If
            ClearResultSet RdoLst
         End With
      End If
   Else
'      sSql = "UPDATE EsReportPlsTable SET PLSPart = '*** No Pick List Recorded ***' " _
'             & "WHERE PLSRow=1"
'      clsADOCon.ExecuteSql sSql
   End If
   Set RdoLst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "buildpartslist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Cover and pick list

'Private Sub PrintCover()
'   Dim sWindows As String
'   MouseCursor 13
'
'   On Error GoTo DiaErr1
'   DoEvents
'   sWindows = GetWindowsDir()
'   SetMdiReportsize MDISect
'   sProcName = "printcover"
'
'#If True Then
'   'OLD WAY
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   sCustomReport = GetCustomReport("prdshcvr")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   MDISect.Crw.Formulas(1) = "Includes='" & cmbPrt & " Run " & cmbRun & "'"
'   MDISect.Crw.Formulas(2) = "Includes2='" & lblDsc & "'"
'   If chkPickList.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.1;T;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;T;;;"
'   End If
'   If chkDoc.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(3) = "GROUPHDR.0.0;F;;;"
'      MDISect.Crw.SectionFormat(4) = "GROUPHDR.0.1;F;;;"
'      MDISect.Crw.SectionFormat(5) = "GROUPHDR.0.2;F;;;"
'      MDISect.Crw.SectionFormat(6) = "GROUPHDR.0.3;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(3) = "GROUPHDR.0.0;T;;;"
'      MDISect.Crw.SectionFormat(4) = "GROUPHDR.0.1;T;;;"
'      MDISect.Crw.SectionFormat(5) = "GROUPHDR.0.2;T;;;"
'      MDISect.Crw.SectionFormat(6) = "GROUPHDR.0.3;T;;;"
'   End If
'#End If
'
'#If False Then
'   'NEW WAY REQUESTED BY LARRY
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   sCustomReport = GetCustomReport("prdshcvr")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   MDISect.Crw.Formulas(1) = "Includes='" & cmbPrt & " Run " & cmbRun & "'"
'   MDISect.Crw.Formulas(2) = "Includes2='" & lblDsc & "'"
'   MDISect.Crw.Formulas(3) = "PartNumber='" & Trim(sPartNumber) & "'"
'   MDISect.Crw.Formulas(4) = "RunNumber='" & Trim(cmbRun) & "'"
'   MDISect.Crw.Formulas(5) = "ShowDocList='" & chkDoc & "'"
'   MDISect.Crw.Formulas(6) = "ShowPickList='" & chkPickList & "'"
'#End If
'
'   SetCrystalAction Me
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   Resume Psh02
'Psh02:
'   DoModuleErrors Me
'
'End Sub
'
'Pick list is active

Private Sub BuildPickList()
   'Dim DbPls  As rdoResultset
   Dim RdoLst As ADODB.Recordset
   Dim b As Byte
   Dim sCol As String * 1
   
   'On Error Resume Next
   On Error GoTo DiaErr1
   sSql = "TRUNCATE TABLE EsReportPlsTable"
   clsADOCon.ExecuteSql sSql
   
'   sSql = "INSERT INTO EsReportPlsTable (PLSRow,PLSPart) VALUES(1,'')"
'   clsADOCon.ExecuteSql sSql
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALOCATION," _
          & "PKPARTREF,PKMOPART,PKPQTY,PKADATE,PKAQTY,PKUNITS FROM PartTable," _
          & "MopkTable WHERE (PARTREF=PKPARTREF AND PKMOPART='" _
          & Compress(cmbPrt) & "' AND PKMORUN=" & cmbRun & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         Do Until .EOF
            b = b + 1
            sCol = Trim$(str(b))
'            If b = 1 Then
'               sSql = "UPDATE EsReportPlsTable SET " _
'                      & "PLSPart='" & Trim(!PartNum) & "'," _
'                      & "PLSDesc='" & Trim(!PADESC) & " '," _
'                      & "PLSPQty=" & Format(!PKPQTY, ES_QuantityDataFormat) & "," _
'                      & "PLSAQty=" & Format(!PKAQTY, ES_QuantityDataFormat) & "," _
'                      & "PLSUom='" & Trim(!PKUNITS) & "'," _
'                      & "PLSLoc='" & Trim(!PALOCATION) & "' " _
'                      & "WHERE PLSRow=1"
'            Else
'               sSql = "INSERT INTO EsReportPlsTable (" _
'                      & "PLSRow,PLSPart,PLSDesc,PLSPQty,PLSAQty,PLSUom,PLSLoc) Values(" _
'                      & sCol & ",'" _
'                      & Trim(!PartNum) & "','" _
'                      & Trim(!PADESC) & "'," _
'                      & Format(!PKPQTY, ES_QuantityDataFormat) & "," _
'                      & Format(!PKAQTY, ES_QuantityDataFormat) & ",'" _
'                      & Trim(!PKUNITS) & "','" _
'                      & Trim(!PALOCATION) & "')"
'            End If
'            clsADOCon.ExecuteSql sSql
'            If Not IsNull(!PKADATE) Then
'               sSql = "UPDATE EsReportPlsTable SET " _
'                      & "PLSADate='" & Format(!PKADATE, "mm/dd/yy") & "' " _
'                      & "WHERE PLSRow=" & sCol & " "
'               clsADOCon.ExecuteSql sSql
'            End If
            Err.Clear
            .MoveNext
         Loop
         ClearResultSet RdoLst
      End With
   Else
'      sSql = "UPDATE EsReportPlsTable SET PLSPart = '*** No Pick List Recorded ***' " _
'             & "PLSRow=1"
'      clsADOCon.ExecuteSql sSql
   End If
   DoEvents
   On Error Resume Next
   Set RdoLst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "buildpicklist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetBMComment(ByVal strPrt As String, ByVal strBomRev As String) As String
    On Error GoTo DiaErr1
    Dim strCmt As String
    
    Dim rdoQry As ADODB.Command
    Dim AdoParameter1 As ADODB.Parameter
    Dim ADOParameter2 As ADODB.Parameter
    
    Dim RdoBMCmt  As ADODB.Recordset
   
    sSql = "SELECT BMCOMT FROM PartTable,BmplTable WHERE (" _
           & "PARTREF=BMPARTREF AND BMASSYPART=? " _
           & " AND BMREV=?) "
   
   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adChar
   ADOParameter2.SIZE = 2
   
   rdoQry.Parameters.Append AdoParameter1
   rdoQry.Parameters.Append ADOParameter2
   
   rdoQry.Parameters(0).Value = strPrt
   rdoQry.Parameters(1).Value = strBomRev
   
   
'   rdoQry(0) = strPrt
'   rdoQry(1) = strBomRev
'
   bSqlRows = clsADOCon.GetQuerySet(RdoBMCmt, rdoQry, ES_STATIC)
   If bSqlRows Then
      With RdoBMCmt
         strCmt = "" & !BMCOMT
      End With
      On Error Resume Next
   Else
      strCmt = ""
   End If
   RdoBMCmt.Close
   
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set RdoBMCmt = Nothing
   Set rdoQry = Nothing
   
   
   GetBMComment = strCmt
   Exit Function
   
DiaErr1:
   sProcName = "GetBMComment"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function





'Private Sub PrintAllocations()
'
'   PrintReportSalesOrderAllocations Me, Trim(sPartNumber), cmbRun
'
'End Sub
