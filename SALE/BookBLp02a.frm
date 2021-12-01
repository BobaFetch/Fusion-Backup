VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form BookBLp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backlog By Scheduled Date"
   ClientHeight    =   4950
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbSlp 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Salesperson From List"
      Top             =   1500
      Width           =   975
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBLp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCls 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   3000
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   3840
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   3600
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optStatCd 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   4080
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BookBLp02a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BookBLp02a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Sales Orders"
      Top             =   1920
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   3720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4950
      FormDesignWidth =   7260
   End
   Begin VB.Label lblSlp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3120
      TabIndex        =   29
      Top             =   1500
      Width           =   3060
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salesperson(s)"
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   28
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   12
      Left            =   3480
      TabIndex        =   27
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   11
      Left            =   4560
      TabIndex        =   25
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   10
      Left            =   4560
      TabIndex        =   24
      Top             =   2640
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   4560
      TabIndex        =   23
      Top             =   1920
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   9
      Left            =   5520
      TabIndex        =   22
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   21
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   3000
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Tag             =   " "
      Top             =   3600
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Desc"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Tag             =   " "
      Top             =   3840
      Width           =   1605
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Status Code"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Tag             =   " "
      Top             =   4080
      Width           =   1605
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduled Dates"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1425
   End
End
Attribute VB_Name = "BookBLp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/25/05 Changed dates and Options
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   If Len(cmbCls) = 0 Then cmbCls = "ALL"
   
End Sub

Private Sub cmbSlp_Click()
   GetSalesPerson
   GetValidCustomer
End Sub

Private Sub cmbSlp_LostFocus()
   cmbSlp = CheckLen(cmbSlp, 4)
   If Len(cmbSlp) = 0 Then cmbSlp = "ALL"
   GetSalesPerson
   GetValidCustomer
End Sub

Private Sub GetSalesPerson()
   Dim rdoSlp As ADODB.Recordset
   On Error GoTo DiaErr1
   If lblSlp.ForeColor = vbRed Then Exit Sub
   sSql = "Qry_GetSalesPerson '" & cmbSlp & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp)
   If bSqlRows Then
      cmbSlp = "" & Trim(rdoSlp!SPNumber)
      lblSlp = "" & Trim(rdoSlp!SPFIRST) & " " & Trim(rdoSlp!SPLAST)
   Else
      lblSlp = "*** Range Of Salespersons ***"
   End If
   On Error Resume Next
   Set rdoSlp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsalesper"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillSales()
   On Error GoTo DiaErr1
   sSql = "Qry_FillSalesPersons"
   LoadComboBox cmbSlp, -1
   If cmbSlp.ListCount > 0 Then
      cmbSlp.AddItem "ALL"
   Else
      lblSlp = "*** No Salespersons Installed ***"
   End If
   cmbSlp = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillsales"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbCst_Click()
   GetCustomer
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(cmbCst) = 0 Then cmbCst = "ALL"
   GetCustomer
   
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
   sSql = "Qry_GetCustomerSalesOrder"
   LoadComboBox cmbCst
   If Not bSqlRows Then
      lblNme = "*** No Customers With SO's Found ***"
   Else
      cmbCst = "ALL"
      GetCustomer
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
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes
      FillProductClasses
      FillSales
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
   
'   optExt.Enabled = True   ' somehow gets greyed out
'   optStatCd.Enabled = True
'   optExt.Visible = True
'   optStatCd.Visible = True
   
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
   Set BookBLp02a = Nothing
   
End Sub

Private Sub PrintReport()
   
   Dim sCust As String
   Dim sCode As String
   Dim sClass As String
   Dim sEnd As String
   Dim sStart As String
   Dim sSlp As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   If cmbCst <> "ALL" Then sCust = Compress(cmbCst)
   If Not IsDate(txtEnd) Then sEnd = "12/31/2024" Else _
                 sEnd = Format(txtEnd, "mm/dd/yyyy")
                 
   If Not IsDate(txtBeg) Then sStart = "1/1/2008" Else _
                 sStart = Format(txtBeg, "mm/dd/yyyy")
                 
   If cmbCde <> "ALL" Then sCode = Compress(cmbCde)
   If cmbCls <> "ALL" Then sClass = Compress(cmbCls)
   If cmbSlp <> "ALL" Then sSlp = Compress(cmbSlp)
   
   On Error GoTo DiaErr1
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowPartDesc"
   aFormulaName.Add "ShowPartExtDesc"
   aFormulaName.Add "ShowStatusCodes"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer(s) " & CStr(cmbCst & ". Sales Person(s) " _
                     & CStr(cmbSlp) & "." _
                    & " From " & txtBeg & " Through " & txtEnd) & "...'")
   
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc
   aFormulaValue.Add optExt
   aFormulaValue.Add optStatCd
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slebl02")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   
'   MdiSect.Crw.StoredProcParam(0) = sEnd ' Start Time
'   MdiSect.Crw.StoredProcParam(1) = cmbCst  ' End Time
'   MdiSect.Crw.StoredProcParam(2) = cmbCls  ' Part Class
'   MdiSect.Crw.StoredProcParam(3) = cmbCde  ' Part Code
    
    aRptPara.Add CStr(sStart)
    aRptPara.Add CStr(sEnd)
    aRptPara.Add CStr(sCust)
    aRptPara.Add CStr(cmbCls.Text)
    aRptPara.Add CStr(cmbCde.Text)
    aRptPara.Add CStr(cmbSlp.Text)
    
    
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    
    ' Set report parameter
   
   sSql = ""
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType     'must happen AFTER SetDbTableConnection call!
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
   cmbCls = "ALL"
   cmbCde = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtEnd) & Trim(str(optExt.Value)) _
              & Trim(str(optDsc.Value)) & Trim(str(optStatCd.Value))
   SaveSetting "Esi2000", "EsiSale", "bL02", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "bL02", Trim(sOptions))
   If Len(sOptions) Then
      optExt = Mid(sOptions, 11, 1)
      optDsc = Mid(sOptions, 12, 1)
      optStatCd = Mid(sOptions, 13, 1)
   Else
      optExt.Value = vbChecked
      optDsc.Value = vbChecked
      optStatCd.Value = vbChecked
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   
End Sub

Private Sub lblNme_Change()
   If Left(lblNme, 9) = "*** No Cu" Then
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optStatCd_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
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

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me

End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If Trim(txtBeg) <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   cmbCst = "ALL"
   cmbSlp = "ALL"
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub GetCustomer()
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetCustomerBasics '" & Compress(cmbCst) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cmbCst = "" & Trim(.Fields(1))
         If Len(cmbCst) > 3 Then
            lblNme = "" & Trim(.Fields(2))
         Else
            lblNme = "*** Range Of Customers Selected ***"
         End If
         ClearResultSet RdoCst
      End With
   Else
      lblNme = "*** Range Of Customers Selected ***"
   End If
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetValidCustomer()
   Dim RdoCst As ADODB.Recordset
   Dim sStart As String
   Dim sEnd As String
   Dim sSalesName As String
   
   On Error GoTo DiaErr1
   If Not IsDate(txtEnd) Then sEnd = "12/31/2024" Else _
                 sEnd = Format(txtEnd, "mm/dd/yyyy")
                 
   If Not IsDate(txtBeg) Then sStart = "1/1/2008" Else _
                 sStart = Format(txtBeg, "mm/dd/yyyy")
   
   sSalesName = Compress(cmbSlp)
   If cmbSlp = "ALL" Then sSalesName = "" Else sSalesName = Compress(cmbSlp)
   
   sSql = "SELECT DISTINCT SOCUST FROM sohdTable WHERE sosalesman LIKE '" _
            & sSalesName & "%' AND SODATE BETWEEN '" & sStart & "' AND '" & sEnd & "'"

   LoadComboBox cmbCst, -1
   Exit Sub
   
DiaErr1:
   sProcName = "getcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   cmbCst = "ALL"
   cmbSlp = "ALL"
End Sub

