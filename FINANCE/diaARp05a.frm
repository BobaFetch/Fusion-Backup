VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Receivable Aging (Report)"
   ClientHeight    =   3660
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3660
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optPIF 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox optAdv 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   255
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox optPad 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   255
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Aging"
      Top             =   840
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp05a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   570
         Picture         =   "diaARp05a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   1
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
      PictureUp       =   "diaARp05a.frx":0308
      PictureDn       =   "diaARp05a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3660
      FormDesignWidth =   6465
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
      PictureUp       =   "diaARp05a.frx":0594
      PictureDn       =   "diaARp05a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Payment In Full"
      Height          =   405
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Advance Payments"
      Height          =   405
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "As Of"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   3720
      TabIndex        =   10
      Top             =   840
      Width           =   2025
   End
End
Attribute VB_Name = "diaARp05a"
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

'*************************************************************************************
' diaARp05 - Print A/R Aging
'
' Notes:
'
' Created: (cjs)
' Modified:
' 06/25/01 (nth) Added "as of" logic and dumped the jet temp db.
' 10/30/03 (nth) fixed issue with cash receipts.
' 11/01/04 (nth) Added advance payment option to report.
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodCustomer As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbCst_Click()
   If cmbCst <> "ALL" Then
      bGoodCustomer = FindThisCustomer(Me)
   Else
      lblNme = "All Customers Selected."
   End If
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Trim(cmbCst) <> "" Then
      bGoodCustomer = FindThisCustomer(Me)
   Else
      lblNme = "All Customers Selected."
   End If
End Sub

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
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT INVCUST,CUREF,CUNICKNAME " _
          & "FROM CihdTable,CustTable WHERE INVCUST=CUREF ORDER BY CUREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      optPIF.Value = 1
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   txtBeg = Format(ES_SYSDATE, "mm/dd/yy")
   bOnLoad = True
   GetOptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaARp05a = Nothing
End Sub
Private Sub PrintReport()
   Dim sTitle As String
   MouseCursor 13
   On Error GoTo DiaErr1
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   ' Which report detail or summary
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   
   If optPad.Value = vbChecked Then
        aFormulaName.Add "Title1"
        aFormulaValue.Add CStr("'Detail Accounts Receivable Aging As Of " & CStr(Trim(txtBeg)) & "'")
        sCustomReport = GetCustomReport("finar05b.rpt")
   Else
        aFormulaName.Add "Title1"
        aFormulaValue.Add CStr("'Summary Accounts Receivable Aging As Of " & CStr(Trim(txtBeg)) & "'")
        sCustomReport = GetCustomReport("finar05a.rpt")
   End If
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
     
   ' Build selection formula
   If Compress(cmbCst) = "" Or Compress(UCase(cmbCst)) = "ALL" Then
      sSql = "{CihdTable.INVNO} > 0 AND {CihdTable.INVDATE} <= #" & Trim(txtBeg) & "#"
   Else
      sSql = "{CihdTable.INVNO} > 0 AND {CihdTable.INVDATE} <= #" & Trim(txtBeg) & _
            "# AND {CihdTable.INVCUST} = '" & Compress(cmbCst) & "'"
   End If
   
   If optAdv Then
      aFormulaName.Add "Title2"
      aFormulaValue.Add CStr("'*** Includes Advance Payments ***'")
   Else
      sSql = sSql & " AND {CihdTable.INVTYPE} <> 'CA'"
   End If
   sSql = sSql & " AND {CihdTable.INVCANCELED} = 0"
   
   aFormulaName.Add "Terms"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "AsOfDate"
   aFormulaName.Add "IncludePIF"
   
   aFormulaValue.Add CStr("''")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(txtBeg)) & "'")
   aFormulaValue.Add CStr("'" & CStr(optPIF) & "'")
   
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
   Dim sTitle As String
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   
   ' Which report detail or summary
   If optPad.Value = vbChecked Then
      sTitle = "Title1='Detail Accounts Receivable Aging As Of " _
               & Trim(txtBeg) & "'"
      sCustomReport = GetCustomReport("finar05b.rpt")
      MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   Else
      sTitle = "Title1='Summary Accounts Receivable Aging As Of " _
               & Trim(txtBeg) & "'"
      MdiSect.crw.ReportFileName = sReportPath & "finar05a.rpt"
      sCustomReport = GetCustomReport("finar05a.rpt")
      MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   End If
   
   ' Set report titles and headers
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = sTitle
   MdiSect.crw.Formulas(2) = "Terms=''"
   MdiSect.crw.Formulas(3) = "RequestBy='Requested By: " _
                        & sInitials & "'"
   MdiSect.crw.Formulas(4) = "AsOfDate='" & Trim(txtBeg) & "'"
   
   ' Build selection formula
   If Compress(cmbCst) = "" Or Compress(UCase(cmbCst)) = "ALL" Then
      sSql = "{CihdTable.INVDATE} <= #" & Trim(txtBeg) & "#"
   Else
      sSql = "{CihdTable.INVDATE} <= #" & Trim(txtBeg) & _
             "# AND {CihdTable.INVCUST} = '" & Compress(cmbCst) & "'"
   End If
   If optAdv Then
      MdiSect.crw.Formulas(5) = "Title2 = '*** Includes Advance Payments ***'"
   Else
      sSql = sSql & " AND {CihdTable.INVTYPE} <> 'CA'"
   End If
   sSql = sSql & " AND {CihdTable.INVCANCELED} = 0"
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
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = Trim(optPad.Value) & Trim(optAdv.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optPad.Value = Val(Left(sOptions, 1))
      optAdv.Value = Val(Mid(sOptions, 2, 1))
   Else
      optPad.Value = vbUnchecked
      optAdv.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPad_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
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
End Sub

