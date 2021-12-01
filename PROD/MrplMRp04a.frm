VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MrplMRp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP Projected Inventory by Week"
   ClientHeight    =   3255
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   4020
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox cboStartDate 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cboEndDate 
      Height          =   315
      Left            =   4020
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox chkPriorActivity 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2940
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "MrplMRp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
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
      Picture         =   "MrplMRp04a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox chkExtDesc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2460
      Width           =   735
   End
   Begin VB.CheckBox chkExceptions 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2700
      Width           =   735
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
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
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp04a.frx":0938
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
         Picture         =   "MrplMRp04a.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3255
      FormDesignWidth =   7260
   End
   Begin VB.Label Label1 
      Caption         =   "Not currently used"
      Height          =   255
      Left            =   3420
      TabIndex        =   32
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   31
      Top             =   1005
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   12
      Left            =   180
      TabIndex        =   30
      Top             =   1380
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   1
      Left            =   2820
      TabIndex        =   29
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   28
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   27
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5580
      TabIndex        =   26
      Top             =   1740
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   9
      Left            =   5580
      TabIndex        =   25
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5580
      TabIndex        =   24
      Top             =   1380
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Prior To First Period"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   23
      Top             =   2940
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   20
      Top             =   2220
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Only Parts With Exceptions"
      Height          =   285
      Index           =   6
      Left            =   180
      TabIndex        =   19
      Top             =   2700
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   7
      Left            =   180
      TabIndex        =   18
      Top             =   2460
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last MRP"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   17
      Top             =   360
      Width           =   975
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   16
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblMrp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1500
      TabIndex        =   15
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblUsr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   12
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "MrplMRp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/19/06 Revised report and selections. Removed extra report.
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub

Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 6)
   If cmbCls = "" Then cmbCls = "ALL"
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

'Private Sub cmdFnd_Click() '
'   ViewParts.lblControl = "TXTPRT"
'   ViewParts.txtPrt = txtPrt
'   optVew.Value = vbChecked
'   ViewParts.Show
'End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub



Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetMRPDates
      GetLastMrp
      cmbCde.AddItem "ALL"
      FillProductCodes
      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      cmbCls.AddItem "ALL"
      FillProductClasses
      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      FillCombos
      bOnLoad = 0
   End If
   If optVew.Value = vbChecked Then
      optVew.Value = vbUnchecked
      Unload ViewParts
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
   Set MrplMRp04a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim sParts As String
    Dim sCode As String
    Dim sClass As String
    Dim sBegDate As String
    Dim sEndDate As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   GetMRPCreateDates sBegDate, sEndDate
   
   If Trim(cboStartDate) = "" Then cboStartDate = "ALL"
   If Trim(cboEndDate) = "" Then cboEndDate = "ALL"

   If Trim(cmbPart) = "" Then cmbPart = "ALL"
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
   If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
   If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdmr04")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowExceptionsOnly"
    aFormulaName.Add "DateDeveloped"
    aFormulaName.Add "ShowExtDesc"
    aFormulaName.Add "ShowPriorPeriodExceptions"
   
    aFormulaValue.Add CStr("'" & sFacility & "...'")
    aFormulaValue.Add CStr("'Parts" & Trim(cmbPart) _
                        & ", Prod Code(s) " & Trim(cmbCde) & ", Class(es) " _
                        & Trim(cmbCls) & "...'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & chkExceptions.Value & "'")
    aFormulaValue.Add CStr("'MRP Created" & lblMrp & "...'")
    aFormulaValue.Add CStr("'" & chkExtDesc.Value & "'")
    aFormulaValue.Add CStr("'" & chkPriorActivity.Value & "'")
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "Includes='Parts " & Trim(cmbpart) _
'                        & ", Prod Code(s) " & Trim(cmbCde) & ", Class(es) " _
'                        & Trim(cmbCls) & "'"
'   MDISect.Crw.Formulas(2) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MDISect.Crw.Formulas(3) = "ShowExceptionsOnly='" & chkExceptions.value & "'"
'   MDISect.Crw.Formulas(4) = "DateDeveloped = 'MRP Created " & lblMrp & "'"
'
'   MDISect.Crw.Formulas(5) = "ShowExtDesc='" & chkExtDesc.value & "'"
'
'   MDISect.Crw.Formulas(6) = "ShowPriorPeriodExceptions='" & chkPriorActivity.value & "'"

   sSql = "{MrplTable.MRP_PARTREF} LIKE '" & sParts & "*'" & vbCrLf _
      & "AND {MrplTable.MRP_PARTPRODCODE} LIKE '" & sCode & "*'" & vbCrLf _
      & "AND {MrplTable.MRP_PARTCLASS} LIKE '" & sClass & "*'" & vbCrLf
   If IsDate(cboStartDate) Then
      sSql = sSql & "AND {MrplTable.MRP_PARTDATERQD} >= Date(" & Format(cboStartDate, "yyyy,mm,dd") & ")" & vbCrLf
   End If
   If IsDate(cboEndDate) Then
      sSql = sSql & "AND {MrplTable.MRP_PARTDATERQD} <= Date(" & Format(cboEndDate, "yyyy,mm,dd") & ") "
   End If
   sSql = sSql & " AND not ({PartTable.PALEVEL} in [6, 5])"
'   MDISect.Crw.SelectionFormula = sSql
   cCRViewer.SetReportSelectionFormula sSql
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

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
   Dim sCode As String * 6
   Dim sClass As String * 4
   sCode = cmbCde
   sClass = cmbCls
   sOptions = sCode & sClass & chkExtDesc.Value & chkExceptions.Value _
      & Me.chkPriorActivity.Value & "0000"
   SaveSetting "Esi2000", "EsiProd", "Prdmr04", sOptions
   SaveSetting "Esi2000", "EsiProd", "Prdmr04Printer", lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr04", sOptions)
   If Len(sOptions) = 0 Then
      sOptions = "          00000000"
   Else
      cmbCde = Mid$(sOptions, 1, 6)
      cmbCls = Mid$(sOptions, 7, 4)
   End If
   
   chkExtDesc.Value = Val(Mid$(sOptions, 11, 1))
   chkExceptions.Value = Val(Mid$(sOptions, 12, 1))
   chkPriorActivity.Value = Val(Mid$(sOptions, 13, 1))
   
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Prdmr04Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"

End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub chkExtDesc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub chkExceptions_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub cboStartDate_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cboStartDate_LostFocus()
   If Len(Trim(cboStartDate)) = 0 Then cboStartDate = "ALL"
   If cboStartDate <> "ALL" Then cboStartDate = CheckDateEx(cboStartDate)
End Sub

Private Sub cboEndDate_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cboEndDate_LostFocus()
   If Len(Trim(cboEndDate)) = 0 Then cboEndDate = "ALL"
   If cboEndDate <> "ALL" Then cboEndDate = CheckDateEx(cboEndDate)
End Sub

Private Sub GetMRPDates()
   
   'by default, show all
   cboStartDate = "ALL"
   cboEndDate = "ALL"
   
End Sub




Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub
