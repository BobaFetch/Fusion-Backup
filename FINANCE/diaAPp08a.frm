VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Payable Aging (Old Report)"
   ClientHeight    =   3465
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox optPst 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With Invoices"
      Top             =   840
      Width           =   1555
   End
   Begin VB.CheckBox optDtl 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaAPp08a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaAPp08a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   7
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
      PictureUp       =   "diaAPp08a.frx":0308
      PictureDn       =   "diaAPp08a.frx":044E
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
      FormDesignHeight=   3465
      FormDesignWidth =   6930
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   15
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
      PictureUp       =   "diaAPp08a.frx":0594
      PictureDn       =   "diaAPp08a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Default Sort By Vendor)"
      Height          =   285
      Index           =   7
      Left            =   3720
      TabIndex        =   19
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Vendor Type"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Age By Post Date"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   3720
      TabIndex        =   14
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cutoff Date"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1545
   End
End
Attribute VB_Name = "diaAPp08a"
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
' diaPap08a - Display Accounts Payable Aging Summary Or Detail.
'
' Notes: Print a "As Of" Accounts Receivable Aging
'
' Created: (cjs)
' Revision:
' 06/03/01 (nth) Removed jet database code, now only using crystal.
' 06/05/01 (nth) Allow for both summary and detail version of report.
' 01/28/03 (nth) Fixed error with as of date not allowing prior dates.
' 10/22/03 (nth) Added custom report.
' 09/10/04 (nth) Added sort by vendor type.
'
'************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodVendor As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmbVnd_Click()
   If cmbVnd <> "ALL" Then
      bGoodVendor = FindVendor(Me)
   Else
      lblNme = "All Vendors."
   End If
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) = 0 Then cmbVnd = "ALL"
   If cmbVnd <> "ALL" Then
      bGoodVendor = FindVendor(Me)
   Else
      lblNme = "All Vendors."
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
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
   Dim RdoVed As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT VIVENDOR,VEREF,VENICKNAME " _
          & "FROM VihdTable,VndrTable WHERE VIVENDOR=VEREF ORDER BY VIVENDOR"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      With RdoVed
         cmbVnd = "ALL"
         AddComboStr cmbVnd.hWnd, "ALL"
         Do Until .EOF
            AddComboStr cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
      End With
   End If
   lblNme = "All Vendors."
   Set RdoVed = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
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
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
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
   Set diaAPp08a = Nothing
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
   
   ' Set report titles and headers
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Terms"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "AsOfDate"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Detail Accounts Payable Aging As Of " & CStr(Trim(txtBeg)) & "'")
   aFormulaValue.Add CStr("' Invoices Aged Based On Terms From Purchase Order/Vendor? N'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(txtBeg)) & "'")
   
   
   If optPst.Value = vbChecked Then
       aFormulaName.Add "AgePost"
       aFormulaValue.Add "'1'"
       aFormulaName.Add "Title2"
       aFormulaValue.Add CStr("'Invoices Aged By Posting Date'")
      sDateColumn = "{VihdTable.VIDTRECD}"
      aSortList.Add "{VihdTable.VETYPE}"
   Else
       aFormulaName.Add "AgePost"
       aFormulaValue.Add "'0'"
       aFormulaName.Add "Title2"
       aFormulaValue.Add CStr("'Invoices Aged By Vendor Invoice Date'")
      sDateColumn = "{VihdTable.VIDATE}"
      aSortList.Add "{VihdTable.VIDATE}"
   End If
   
   ' Report path based on detail or summary types of reports
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   If optDtl.Value = vbChecked Then
      sCustomReport = GetCustomReport("finap08b.rpt")
      'sCustomReport = "finap08b.rpt"
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
   Else
      sCustomReport = GetCustomReport("finap08a.rpt")
      'sCustomReport = "finap08a.rpt"
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
   End If
   
      
   sSql = ""
   sSql = cCRViewer.GetReportSelectionFormula
   
   If (sSql <> "") Then
      sSql = sSql & " AND "
   End If
   
   ' Determine selection formula
   If cmbVnd = "ALL" Then
      sSql = sSql & sDateColumn & " <=#" & Trim(txtBeg) & "#"
   Else
      sSql = sSql & sDateColumn & " <=#" & Trim(txtBeg) & "# " _
             & "AND {VndrTable.VENICKNAME}='" & Trim(cmbVnd) & "'"
   End If
   
   
   cCRViewer.SetSortFields aSortList
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
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

Private Sub PrintReport1()
   Dim sDateColumn As String
   Dim sCustomReport As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   'SetMdiReportsize MdiSect
   
   ' Set report titles and headers
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Title1='Detail Accounts Payable Aging As Of " & Trim(txtBeg) & "'"
   MdiSect.crw.Formulas(2) = "Terms=' Invoices Aged Based On Terms From Purchase Order/Vendor? N'"
   MdiSect.crw.Formulas(3) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(4) = "AsOfDate='" & Trim(txtBeg) & "'"
   
   If optPst.Value = vbChecked Then
      MdiSect.crw.Formulas(5) = "AgePost='1'"
      MdiSect.crw.Formulas(6) = "Title2 ='Invoices Aged By Posting Date'"
      sDateColumn = "{VihdTable.VIDTRECD}"
   Else
      MdiSect.crw.Formulas(5) = "AgePost='0'"
      MdiSect.crw.Formulas(6) = "Title2 = 'Invoices Aged By Vendor Invoice Date'"
      sDateColumn = "{VihdTable.VIDATE}"
   End If
   
   ' Report path based on detail or summary types of reports
   If optDtl.Value = vbChecked Then
      sCustomReport = GetCustomReport("finap08b.rpt")
      MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   Else
      sCustomReport = GetCustomReport("finap08a.rpt")
      MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   End If
   
   ' Determine selection formula
   If cmbVnd = "ALL" Then
      sSql = sDateColumn & " <=#" & Trim(txtBeg) & "#"
   Else
      sSql = sDateColumn & " <=#" & Trim(txtBeg) & "# " _
             & "AND {VndrTable.VENICKNAME}='" & Trim(cmbVnd) & "'"
   End If
   
   MdiSect.crw.SortFields(0) = "+{VndrTable.VETYPE}"
   
   MdiSect.crw.SelectionFormula = sSql
   
   ' Show Report
   'SetCrystalAction Me
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

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = optPst.Value & optDtl.Value & optTyp
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optPst.Value = Mid(sOptions, 1, 1)
      optDtl.Value = Mid(sOptions, 2, 1)
      optTyp.Value = Mid(sOptions, 3, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name _
                & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
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

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
'   If CDate(txtBeg) > CDate(Now) Then
'      Beep
'      txtBeg = Format(Now, "mm/dd/yy")
'   End If
End Sub
