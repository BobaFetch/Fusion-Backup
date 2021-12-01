VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSAp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gross Margin Analysis"
   ClientHeight    =   2685
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2685
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBottom 
      Height          =   315
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "Requires A Number"
      Top             =   1860
      Width           =   495
   End
   Begin VB.TextBox txtTop 
      Height          =   315
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   17
      Tag             =   "1"
      ToolTipText     =   "Requires A Number"
      Top             =   1500
      Width           =   495
   End
   Begin VB.ComboBox cboCategory 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "diaSAp01a.frx":0000
      Left            =   1320
      List            =   "diaSAp01a.frx":0022
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
      Top             =   2280
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
      FormDesignHeight=   2685
      FormDesignWidth =   5835
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
         Picture         =   "diaSAp01a.frx":0091
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
         Picture         =   "diaSAp01a.frx":020F
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
      PictureUp       =   "diaSAp01a.frx":0399
      PictureDn       =   "diaSAp01a.frx":04DF
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
      PictureUp       =   "diaSAp01a.frx":0625
      PictureDn       =   "diaSAp01a.frx":076B
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percent"
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   20
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percent"
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   19
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Top"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Bottom"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1920
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
      Top             =   2280
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
Attribute VB_Name = "diaSAp01a"
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
' diaSAp01a - Gross Margin Analysis
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
                             X As Single, Y As Single)
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
  
   sCustomReport = GetCustomReport("finsa01")
   
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Category"
   aFormulaName.Add "StartDate"
   aFormulaName.Add "EndDate"
   aFormulaName.Add "Top"
   aFormulaName.Add "Bottom"
   aFormulaName.Add "ShowDetail"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboCategory) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboStart) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboEnd) & "'")
   aFormulaValue.Add CStr("'" & CStr(txtTop) & "'")
   aFormulaValue.Add CStr("'" & CStr(txtBottom) & "'")
   aFormulaValue.Add chkDetails.Value

   
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
'   MdiSect.Crw.Formulas(2) = "Category='" & cboCategory & "'"
'   MdiSect.Crw.Formulas(3) = "StartDate='" & cboStart & "'"
'   MdiSect.Crw.Formulas(4) = "EndDate='" & cboEnd & "'"
'   MdiSect.Crw.Formulas(5) = "Top='" & txtTop & "'"
'   MdiSect.Crw.Formulas(6) = "Bottom='" & txtBottom & "'"
'   MdiSect.Crw.Formulas(7) = "ShowDetail=" & chkDetails.Value

'{Vw_Sales_New.INTYPE} IN [24.00, 25.00, 26.00, 3.00, 4.00] and

   cCRViewer.SetDbTableConnection
   sSql = "{CihdTable.INVCANCELED} <> 1.00 and {CihdTable.INVDATE} in CDateTime ({@StartDate}) to CDateTime ({@EndDate})"
   cCRViewer.SetReportSelectionFormula (sSql)
   'MdiSect.crw.SelectionFormula = sSql
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.CRViewerSize Me
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

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
Private Sub PrintReport1()
   MouseCursor 13
   On Error GoTo DiaErr1
   optPrn.enabled = False
   optDis.enabled = False
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("finsa01.rpt")
   
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
'   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
'   MdiSect.crw.Formulas(2) = "Category='" & cboCategory & "'"
'   MdiSect.crw.Formulas(3) = "StartDate='" & cboStart & "'"
'   MdiSect.crw.Formulas(4) = "EndDate='" & cboEnd & "'"
'   MdiSect.crw.Formulas(5) = "Top='" & txtTop & "'"
'   MdiSect.crw.Formulas(6) = "Bottom='" & txtBottom & "'"
'   MdiSect.crw.Formulas(7) = "ShowDetail=" & chkDetails.Value
   
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
   sOptions = chkDetails
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & "StartDate", cboStart
   SaveSetting "Esi2000", "EsiFina", Me.Name & "EndDate", cboEnd
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Category", cboCategory
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Top", txtTop
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Bottom", txtBottom
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, "0")
   '    If Len(Trim(sOptions)) >= 2 Then
   '        chkDetails.Value = Mid(sOptions, 2, 1)
   '    Else
   '        chkDetails.Value = vbUnchecked
   '    End If
   chkDetails.Value = CInt(sOptions)
   
   Dim defaultDate As String
   defaultDate = Format(Date, "mm/dd/yyyy")
   cboStart = GetSetting("Esi2000", "EsiFina", Me.Name & "StartDate", defaultDate)
   cboEnd = GetSetting("Esi2000", "EsiFina", Me.Name & "EndDate", defaultDate)
   cboCategory = GetSetting("Esi2000", "EsiFina", Me.Name & "Category", cboCategory.List(0))
   txtTop = GetSetting("Esi2000", "EsiFina", Me.Name & "Top", "10")
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
