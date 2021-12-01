VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchasing History By Part"
   ClientHeight    =   3675
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "PurcPRp10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
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
      Picture         =   "PurcPRp10a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Character Search Or Select Contains Part Numbers On Purchase Orders"
      Top             =   1080
      Width           =   3545
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "PurcPRp10a.frx":0938
      Height          =   315
      Left            =   6120
      Picture         =   "PurcPRp10a.frx":0C7A
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   3240
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   2880
      Width           =   1555
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Enter/Revise Product Class (4 Char)"
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise Product Code (6 Char)"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbTyp 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
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
         Height          =   320
         Left            =   0
         Picture         =   "PurcPRp10a.frx":0FBC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   320
         Left            =   560
         Picture         =   "PurcPRp10a.frx":113A
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
      FormDesignHeight=   3675
      FormDesignWidth =   7215
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   27
      Top             =   0
      Width           =   1680
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   5500
      TabIndex        =   23
      Top             =   2880
      Width           =   1400
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   5500
      TabIndex        =   22
      Top             =   2520
      Width           =   1400
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   5500
      TabIndex        =   21
      Top             =   2160
      Width           =   1400
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   5500
      TabIndex        =   20
      Top             =   1800
      Width           =   1400
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   3
      Left            =   5500
      TabIndex        =   19
      Top             =   1440
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Due from"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   5500
      TabIndex        =   12
      Top             =   1080
      Width           =   1400
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "PurcPRp10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/21/07 Changed Level and Vendor to ALL
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
   
End Sub

Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   
End Sub

Private Sub cmbTyp_LostFocus()
   If Trim(cmbTyp) = "" Then
      cmbTyp = "ALL"
   Else
      cmbTyp = Format(Abs(Val(cmbTyp)), "0")
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
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
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PIPART,PARTREF,PARTNUM " _
          & "FROM PoitTable,PartTable WHERE PIPART=PARTREF " _
          & "ORDER BY PIPART"
   LoadComboBox txtPrt, 1
   txtPrt = "ALL"
   cmbCde = "ALL"
   cmbCls = "ALL"
   For b = 1 To 7
      cmbTyp.AddItem b
   Next
   cmbTyp.AddItem b
   cmbTyp = "ALL"
   If Trim(txtPrt) = "" Then txtPrt = "ALL"
   cmbVnd = "ALL"
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
      cmbCde.AddItem "ALL"
      FillProductCodes
      cmbCls.AddItem "ALL"
      FillProductClasses
      cmbVnd.AddItem "ALL"
      FillVendors
      FillCombo
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
   Set PurcPRp10a = Nothing
   
End Sub
Private Sub PrintReport()
    Dim sPart As String
    Dim sBDate As String
    Dim sEDate As String
    Dim sCode As String
    Dim sClass As String
    Dim sVendor As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If txtPrt <> "ALL" Then
      sPart = Compress(txtPrt)
   Else
      sPart = ""
   End If
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If Trim(txtEnd) = "" Then txtEnd = "ALL"
   
   If Not IsDate(txtBeg) Then
      sBDate = "1995,01,01"
   Else
      sBDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEDate = "2024,12,31"
   Else
      sEDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   If cmbCde = "ALL" Then
      sCode = ""
   Else
      sCode = Compress(cmbCde)
   End If
   If cmbCls = "ALL" Then
      sClass = ""
   Else
      sClass = Compress(cmbCls)
   End If
   If cmbVnd = "ALL" Then
      sVendor = ""
   Else
      sVendor = Compress(cmbVnd)
   End If
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtPrt) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("prdpr12")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "{PoitTable.PIPART} LIKE '" & sPart & "*' " _
          & "AND {PoitTable.PIPDATE} In Date(" & sBDate & ") " _
          & " To Date(" & sEDate & ") "
   If Val(cmbTyp) > 0 Then sSql = sSql & "AND {PartTable.PALEVEL}=" & cmbTyp & " "
   sSql = sSql & "AND {PartTable.PAPRODCODE} LIKE '" & sCode & "*' "
   sSql = sSql & "AND {PartTable.PACLASS} LIKE '" & sClass & "*' "
   sSql = sSql & "AND {PohdTable.POVENDOR} LIKE '" & sVendor & "*' "
   sSql = sSql & " AND {PoitTable.PIITEM} > 0.00"
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
   
   'txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   'txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 4)
   txtBeg = ""
   txtEnd = ""
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiProd", "Pmr01", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Ppr01", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then
      txtBeg = "ALL"
   Else
      txtBeg = CheckDateEx(txtBeg)
   End If
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then
      txtEnd = "ALL"
   Else
      txtEnd = CheckDateEx(txtEnd)
   End If
   
End Sub


Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Then txtPrt = "ALL"
   
End Sub
