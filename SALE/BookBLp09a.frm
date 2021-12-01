VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form BookBLp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backlog By Part Number"
   ClientHeight    =   5160
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
   ScaleHeight     =   5160
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "BookBLp09a.frx":0000
      Height          =   315
      Left            =   5040
      Picture         =   "BookBLp09a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   960
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2040
      TabIndex        =   30
      Tag             =   "3"
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   2040
      TabIndex        =   27
      Top             =   915
      Width           =   3015
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   25
      Tag             =   "4"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBLp09a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   24
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
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   2040
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   4320
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   4080
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optStatCd 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   4560
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   0
      Tag             =   "4"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BookBLp09a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BookBLp09a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   10
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
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Sales Orders"
      Top             =   2880
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5160
      FormDesignWidth =   7260
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   32
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   915
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   13
      Left            =   5520
      TabIndex        =   28
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   12
      Left            =   3480
      TabIndex        =   26
      Top             =   2520
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   11
      Left            =   4560
      TabIndex        =   23
      Top             =   2040
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   10
      Left            =   4560
      TabIndex        =   22
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   4560
      TabIndex        =   21
      Top             =   2880
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   9
      Left            =   5520
      TabIndex        =   20
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Tag             =   " "
      Top             =   4080
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Desc"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Tag             =   " "
      Top             =   4320
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Status Code"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Tag             =   " "
      Top             =   4560
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduled Dates"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1995
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   1425
   End
End
Attribute VB_Name = "BookBLp09a"
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


Private Sub cmbCst_Click()
   GetCustomer
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(cmbCst) = 0 Then cmbCst = "ALL"
   GetCustomer
   
End Sub


Private Sub cmbPart_Click()
      GetPart
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPart = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPart
   ViewParts.Show
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
      txtPrt = "ALL"
      cmbPart = "ALL"
      FillProductCodes
      FillProductClasses
      FillCombo
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartNumber
      
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

Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
    GetPart
End Sub

Private Sub lblDsc_Click()
   If Left(lblDsc, 12) = "*** No Parts" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   cmbPart = txtPrt
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
End Sub


Private Sub FillPartNumber()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable, SoitTable  " _
        & " WHERE ITPART=PARTREF AND" _
      & " ITCANCELED=0 AND SoitTable.ITPSNUMBER=''" _
      & " AND ITINVOICE=0 AND SoitTable.ITPSSHIPPED=0"
    
    LoadComboBox cmbPart, 0, False
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
    GetPart
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set BookBLp09a = Nothing
   
End Sub

Private Sub PrintReport()
   
   Dim sCust As String
   Dim sCode As String
   Dim sClass As String
   Dim sEnd As String
   Dim sStart As String
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
   
   Dim partKey As String
   partKey = Compress(cmbPart.Text)
   If lblDsc = "*** Range Of Parts Selected ***" And partKey <> "ALL" Then
      'partKey = partKey & "%"  -- crystal doesn't pass % to sql
   End If

   
   On Error GoTo DiaErr1
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowPartDesc"
   aFormulaName.Add "ShowPartExtDesc"
   aFormulaName.Add "ShowStatusCodes"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Parts " & partKey & " Customer(s) " & CStr(cmbCst & "." _
                    & " From " & txtBeg & " Through " & txtEnd) & "...'")
   
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc
   aFormulaValue.Add optExt
   aFormulaValue.Add optStatCd
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slebl09")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   'cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   
'   MdiSect.Crw.StoredProcParam(0) = sEnd ' Start Time
'   MdiSect.Crw.StoredProcParam(1) = cmbCst  ' End Time
'   MdiSect.Crw.StoredProcParam(2) = cmbCls  ' Part Class
'   MdiSect.Crw.StoredProcParam(3) = cmbCde  ' Part Code
    
    aRptPara.Add CStr(sStart)
    aRptPara.Add CStr(sEnd)
    
   'aRptPara.Add CStr(Compress(cmbPart.Text))
'   Dim partKey As String
'   partKey = Compress(cmbPart.Text)
'   If lblDsc = "*** Range Of Parts Selected ***" And partKey <> "ALL" Then
'      partKey = partKey & "%"
'   End If
   aRptPara.Add partKey

    aRptPara.Add CStr(cmbCst.Text)
    aRptPara.Add CStr(cmbCls.Text)
    aRptPara.Add CStr(cmbCde.Text)
    
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    
    ' Set report parameter
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   
   sSql = ""
   cCRViewer.SetReportSelectionFormula sSql
   
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
      optExt = Mid(sOptions, 9, 1)
      optDsc = Mid(sOptions, 10, 1)
      optStatCd = Mid(sOptions, 11, 1)
   Else
      optExt.Value = vbChecked
      optDsc.Value = vbChecked
      optStatCd.Value = vbChecked
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   
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
   ShowCalendar Me

End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If Trim(txtBeg) <> "ALL" Then txtBeg = CheckDate(txtBeg)

End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
   
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

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDate(txtEnd)
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPart.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPart.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

Private Sub GetPart()
   Dim RdoPrt As ADODB.Recordset
   sSql = "Qry_GetPartNumberBasics '" & Compress(cmbPart) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPart = "" & Trim(.Fields(1))
         If Len(cmbPart) > 3 Then
            lblDsc = "" & Trim(.Fields(2))
         Else
            lblDsc = "*** Range Of Parts Selected ***"
         End If
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "*** Range Of Parts Selected ***"
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


