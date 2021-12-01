VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form AdmnProd25 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Labor Efficiency Reports"
   ClientHeight    =   3660
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3660
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Analyze by"
      Height          =   1215
      Left            =   660
      TabIndex        =   14
      Top             =   2160
      Width           =   3255
      Begin VB.OptionButton optWorkCenter 
         Caption         =   "WorkCenter"
         Height          =   435
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optEmployee 
         Caption         =   "Employee"
         Height          =   315
         Left            =   720
         TabIndex        =   15
         Top             =   720
         Width           =   1875
      End
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      Tag             =   "4"
      Top             =   1620
      Width           =   1095
   End
   Begin VB.ComboBox cboShop 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With Invoices"
      Top             =   660
      Width           =   1555
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1140
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6300
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6300
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "AdmnProd25.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "AdmnProd25.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
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
      PictureUp       =   "AdmnProd25.frx":0308
      PictureDn       =   "AdmnProd25.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3660
      FormDesignWidth =   7980
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   10
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
      PictureUp       =   "AdmnProd25.frx":0594
      PictureDn       =   "AdmnProd25.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Op Comp Date"
      Height          =   285
      Index           =   3
      Left            =   660
      TabIndex        =   13
      Top             =   1620
      Width           =   1395
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   9
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   1
      Left            =   660
      TabIndex        =   8
      Top             =   660
      Width           =   1125
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Op Comp Date"
      Height          =   285
      Index           =   4
      Left            =   660
      TabIndex        =   7
      Top             =   1140
      Width           =   1545
   End
End
Attribute VB_Name = "AdmnProd25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
' ShopSHp24 - Priority Dispatch Report.
'
' Revision:
' 10/29/2018 TEL Created report for EBM

'************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodVendor As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub FillCombo()
   Dim rs As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SHPNUM from ShopTable ORDER BY SHPNUM"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs)
   If bSqlRows Then
      With rs
         Do Until .EOF
            AddComboStr cboShop.hwnd, "" & Trim(!SHPNUM)
            .MoveNext
         Loop
      End With
   End If
   Set rs = Nothing
   Exit Sub
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   txtBeg = Format(ES_SYSDATE, "mm/dd/yy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   optWorkCenter.Value = True
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdmnProd25 = Nothing
End Sub

Private Sub PrintReport()
   
   Dim sDateColumn As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim aSortList As New Collection
   
   On Error GoTo DiaErr1
   
   MouseCursor 13
   
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   
   aFormulaName.Add "RequestBy"
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   
   ' Report path based on detail or summary types of reports
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   If optWorkCenter Then
      aFormulaValue.Add "'Labor Efficiency by WorkCenter'"
      sCustomReport = GetCustomReport("AdmnProd25a.rpt")
   Else
      aFormulaValue.Add "'Labor Efficiency by Employee'"
      sCustomReport = GetCustomReport("AdmnProd25b.rpt")
   End If
   aFormulaValue.Add "'For operation Completions from " & Trim(txtBeg) & " through " & Trim(txtEnd) & "'"
   
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   ' view report
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection   ' zeros parameters.  Set parameters immediately before OpenCrystalReportObject call
   cCRViewer.ShowGroupTree False
   
   ' report parameter
   aRptParaType.Add CStr("String")  'parameter 1 = @Shop
   aRptParaType.Add CStr("String")  'parameter 2 = @StartDate
   aRptParaType.Add CStr("String")  'parameter 3 = @EndDate
   
   aRptPara.Add CStr(cboShop)
   aRptPara.Add CStr(txtBeg)
   aRptPara.Add CStr(txtEnd)
   
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType   'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   cCRViewer.ClearFieldCollection aSortList
   
   ' Show Report
   MouseCursor 0
   Exit Sub
   ' Handle runtime errors
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub


Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

