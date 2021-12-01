VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSAp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Top / Bottom Sales Analysis"
   ClientHeight    =   3765
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3765
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTop 
      Height          =   315
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   26
      Tag             =   "1"
      Text            =   "5"
      ToolTipText     =   "Requires A Number"
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton optTop 
      Caption         =   "Top Sales"
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OptionButton optBottom 
      Caption         =   "Bottom Sales"
      Height          =   255
      Left            =   2040
      TabIndex        =   23
      Top             =   2280
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CheckBox chkGraph 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtBottom 
      Height          =   315
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   18
      Tag             =   "1"
      Text            =   "5"
      ToolTipText     =   "Requires A Number"
      Top             =   3300
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtNSales 
      Height          =   315
      Left            =   840
      MaxLength       =   3
      TabIndex        =   17
      Tag             =   "1"
      ToolTipText     =   "Requires A Number"
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox cboCategory 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "diaSAp02a.frx":0000
      Left            =   1320
      List            =   "diaSAp02a.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   420
      Width           =   2175
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox cboEnd 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1140
      Width           =   1515
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "4"
      Top             =   780
      Width           =   1515
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5040
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3765
      FormDesignWidth =   6720
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4680
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaSAp02a.frx":0084
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
         Picture         =   "diaSAp02a.frx":0202
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
      TabIndex        =   8
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
      PictureUp       =   "diaSAp02a.frx":038C
      PictureDn       =   "diaSAp02a.frx":04D2
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   9
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
      PictureUp       =   "diaSAp02a.frx":0618
      PictureDn       =   "diaSAp02a.frx":075E
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select "
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Graph"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   19
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Top"
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Bottom"
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   825
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaSAp02a"
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

'*********************************************************************************
' diaSAp02a - Top / Bottom Sales Analysis
'
' Notes:
'
' Created: 01/29/03 (nth)
' Revisions:
'   03/04/04 (nth) Added detail for Linda
'   03/05/04 (nth) Added cash basis version of report for Linda
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   cboStart = Format(Now, "mm/01/yy")
   cboEnd = Format(Now, "mm/31/yy")
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
   Set diaARp15a = Nothing
End Sub

'Private Sub optCsh_Click()
'    If optCsh Then
'        optDtl = vbChecked
'        optDtl.enabled = False
'    Else
'        optDtl.enabled = True
'    End If
'End Sub
'

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

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub
Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo DiaErr1
   
   txtTop.Text = 2
   txtBottom.Text = 2
   
   optPrn.enabled = False
   optDis.enabled = False
   'SetMdiReportsize MdiSect
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   
   If (optTop.Value) Then
      sCustomReport = GetCustomReport("finsa02.rpt")
   Else
      sCustomReport = GetCustomReport("finsa02a.rpt")
   End If
   
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Category"
   aFormulaName.Add "StartDate"
   aFormulaName.Add "EndDate"
   aFormulaName.Add "Top"
   aFormulaName.Add "Bottom"
   aFormulaName.Add "NSales"
   aFormulaName.Add "ShowDetail"
   aFormulaName.Add "ShowGraph"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboCategory) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboStart) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboEnd) & "'")
   aFormulaValue.Add CStr("'" & CStr(txtTop) & "'")
   aFormulaValue.Add CStr("'" & CStr(txtBottom) & "'")
   aFormulaValue.Add CStr("'" & CStr(txtNSales) & "'")
   aFormulaValue.Add chkDetails.Value
   aFormulaValue.Add chkGraph.Value

   
   'MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
'   MdiSect.Crw.Formulas(2) = "Category='" & cboCategory & "'"
'   MdiSect.Crw.Formulas(3) = "StartDate='" & cboStart & "'"
'   MdiSect.Crw.Formulas(4) = "EndDate='" & cboEnd & "'"
'   MdiSect.Crw.Formulas(5) = "Top='" & txtTop & "'"
'   MdiSect.Crw.Formulas(6) = "Bottom='" & txtBottom & "'"
'   MdiSect.Crw.Formulas(7) = "ShowDetail=" & chkDetails.Value
'   MdiSect.Crw.Formulas(8) = "ShowGraph=" & chkGraph.Value
   
   'MdiSect.crw.SelectionFormula = sSql
   '{Vw_Sales.INVNO} = 121887 AND {Vw_Sales.INVCUST} = 'BOEMIL' AND
   sSql = "{Vw_Sales.INVCANCELED} <> 1.00 and {Vw_Sales.INVDATE} in CDateTime ({@StartDate}) to CDateTime ({@EndDate})"
   cCRViewer.SetReportSelectionFormula (sSql)
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

'   SetCrystalAction Me
   optPrn.enabled = True
   optDis.enabled = True
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   MouseCursor 13
   On Error GoTo DiaErr1
   optPrn.enabled = False
   optDis.enabled = False
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("finsa02.rpt")
   
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Category='" & cboCategory & "'"
   MdiSect.crw.Formulas(3) = "StartDate='" & cboStart & "'"
   MdiSect.crw.Formulas(4) = "EndDate='" & cboEnd & "'"
   MdiSect.crw.Formulas(5) = "Top='" & txtTop & "'"
   MdiSect.crw.Formulas(6) = "Bottom='" & txtBottom & "'"
   MdiSect.crw.Formulas(7) = "ShowDetail=" & chkDetails.Value
   MdiSect.crw.Formulas(8) = "ShowGraph=" & chkGraph.Value
   
   'MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   optPrn.enabled = True
   optDis.enabled = True
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sOptNTop As String
   
   sOptions = chkDetails
   sOptNTop = optTop.Value
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptNTop
   
   SaveSetting "Esi2000", "EsiFina", Me.Name & "StartDate", cboStart
   SaveSetting "Esi2000", "EsiFina", Me.Name & "EndDate", cboEnd
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Category", cboCategory
   SaveSetting "Esi2000", "EsiFina", Me.Name & "NSales", txtNSales

   'SaveSetting "Esi2000", "EsiFina", Me.Name & "Top", txtTop
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Bottom", txtBottom
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim sOptNTop As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, "0")
   '    If Len(Trim(sOptions)) >= 2 Then
   '        chkDetails.Value = Mid(sOptions, 2, 1)
   '    Else
   '        chkDetails.Value = vbUnchecked
   '    End If
   chkDetails.Value = CInt(sOptions)
   
   sOptNTop = GetSetting("Esi2000", "EsiFina", Me.Name, "0")
   optTop.Value = sOptNTop
   
   Dim defaultDate As String
   defaultDate = Format(Date, "mm/dd/yyyy")
   cboStart = GetSetting("Esi2000", "EsiFina", Me.Name & "StartDate", defaultDate)
   cboEnd = GetSetting("Esi2000", "EsiFina", Me.Name & "EndDate", defaultDate)
   cboCategory = GetSetting("Esi2000", "EsiFina", Me.Name & "Category", cboCategory.List(0))
   txtNSales = GetSetting("Esi2000", "EsiFina", Me.Name & "NSales", "10")
   'txtTop = GetSetting("Esi2000", "EsiFina", Me.Name & "Top", "10")
   txtBottom = GetSetting("Esi2000", "EsiFina", Me.Name & "Bottom", "10")
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub cboEnd_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboEnd_LostFocus()
   cboStart = CheckDate(cboStart)
End Sub

Private Sub cboStart_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboStart_LostFocus()
   cboStart = CheckDate(cboStart)
End Sub

