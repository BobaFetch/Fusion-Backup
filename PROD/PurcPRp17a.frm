VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form PurcPRp17a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Delivery Performance and Inspection Report"
   ClientHeight    =   5310
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optSum 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   29
      Top             =   4800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox optUseOriginalSD 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   27
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame z2 
      Height          =   540
      Left            =   1560
      TabIndex        =   22
      Top             =   3720
      Width           =   3855
      Begin VB.OptionButton optCom 
         Caption         =   "Complete"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCom 
         Caption         =   "Incomplete"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCom 
         Caption         =   "Both"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbLateCalc 
      Height          =   315
      ItemData        =   "PurcPRp17a.frx":0000
      Left            =   1560
      List            =   "PurcPRp17a.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "9"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox cboVendor 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      Top             =   360
      Width           =   1555
   End
   Begin VB.TextBox txtDaysLate 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "1"
      Text            =   "0"
      Top             =   2340
      Width           =   555
   End
   Begin VB.TextBox txtDaysEarly 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Tag             =   "1"
      Text            =   "0"
      Top             =   1980
      Width           =   555
   End
   Begin VB.ComboBox cboEnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1140
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5310
      FormDesignWidth =   6210
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4680
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PurcPRp17a.frx":0050
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PurcPRp17a.frx":01CE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   9
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
      PictureUp       =   "PurcPRp17a.frx":0358
      PictureDn       =   "PurcPRp17a.frx":049E
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
      PictureUp       =   "PurcPRp17a.frx":05E4
      PictureDn       =   "PurcPRp17a.frx":072A
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   31
      Tag             =   " "
      Top             =   4440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Summary Only"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   30
      Top             =   4800
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Original Ship Date"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action"
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "Chart Results"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Late Calculation by"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Nickname"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   435
      Width           =   1395
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   18
      Top             =   795
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2340
      Width           =   1005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Late"
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   16
      Top             =   2340
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1980
      Width           =   1005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Early"
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   14
      Top             =   1980
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1140
      Width           =   825
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "PurcPRp17a"
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
' PurcPRp17a - Vendor Delivery Performance and Inspection report.
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

Private Sub cboVendor_Click()
   lblName = GetVendorName(cboVendor)
End Sub

Private Sub cboVendor_GotFocus()
   ComboGotFocus cboVendor
End Sub

Private Sub cboVendor_KeyUp(KeyCode As Integer, Shift As Integer)
   ComboKeyUp cboVendor, KeyCode
   lblName = GetVendorName(cboVendor)
End Sub

Private Sub cboVendor_LostFocus()
   lblName = GetVendorName(cboVendor)
End Sub

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
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
      LoadComboWithVendors cboVendor, True
      GetOptions
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   'cboStart = Format(Now, "mm/01/yy")
   'cboEnd = Format(Now, "mm/31/yy")
   'GetOptions
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
   Set PurcPRp17a = Nothing
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
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sSubSql As String
   
   optPrn.Enabled = False
   optDis.Enabled = False
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdpr20.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Vendor"
    aFormulaName.Add "StartDate"
    aFormulaName.Add "EndDate"
    aFormulaName.Add "DaysEarly"
    aFormulaName.Add "DaysLate"
    aFormulaName.Add "UseOriginalShipDate"  'BBS Added on 03/08/2010 for Ticket #11364
    aFormulaName.Add "LateCalculationBy"
    aFormulaName.Add "ShowSumOnly"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Vendor Delivery Performance and Inspection Report'")
    aFormulaValue.Add CStr("'Vendors Included: " & lblName & " (" & CStr(cboVendor) & ")'")
    aFormulaValue.Add CStr("'" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboVendor) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboStart) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboEnd) & "'")
    aFormulaValue.Add txtDaysEarly
    aFormulaValue.Add txtDaysLate
    aFormulaValue.Add optUseOriginalSD.Value    'BBS Added on 03/08/2010 for Ticket #11364
    If cmbLateCalc.ListIndex = 0 Then aFormulaValue.Add CStr("'Date Received'") Else _
       If cmbLateCalc.ListIndex = 1 Then aFormulaValue.Add CStr("'On Dock Inspection Date'") Else aFormulaValue.Add CStr("'On Dock Delivered Date'")
    
    aFormulaValue.Add optSum.Value
    
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
    
    aRptPara.Add CStr(cboVendor)
    aRptPara.Add CStr(cboStart)
    aRptPara.Add CStr(cboEnd)
    aRptPara.Add txtDaysEarly
    aRptPara.Add txtDaysLate
    aRptPara.Add optUseOriginalSD
    aRptPara.Add cmbLateCalc.ListIndex
    
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")

    
   If ((CStr(cboVendor) <> "ALL") And (CStr(cboVendor) <> "<ALL>")) Then
      sSubSql = "{RjhdTable.REJVENDOR} = '" & CStr(cboVendor) & "' AND "
   Else
      sSubSql = ""
   End If
   
   sSubSql = sSubSql & " {RjhdTable.REJDATE} In Date(" & Format(cboStart, "yyyy,mm,dd") & ") " _
          & "To Date(" & Format(cboEnd, "yyyy,mm,dd") & ")"
   
   If optCom(0).Value = True Then
      sSubSql = sSubSql & " AND {RjitTable.RITACT}=1"
   Else
      If optCom(1).Value = True Then sSubSql = sSubSql & " AND {RjitTable.RITACT}=0"
   End If
   sSubSql = sSubSql & " and {RjhdTable.REJTYPE} = 'V'"
   cCRViewer.SetSubRptSelFormula "quarj06.rpt", sSubSql
    
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType    'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   cCRViewer.ShowGroupTree False
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   optPrn.Enabled = True
   optDis.Enabled = True
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim strRegApp As String
   strRegApp = GetRegistryAppTitle()

   SaveSetting "Esi2000", strRegApp, Me.Name & "StartDate", cboStart
   SaveSetting "Esi2000", strRegApp, Me.Name & "EndDate", cboEnd
   SaveSetting "Esi2000", strRegApp, Me.Name & "Vendor", cboVendor
   SaveSetting "Esi2000", strRegApp, Me.Name & "DaysEarly", txtDaysEarly
   SaveSetting "Esi2000", strRegApp, Me.Name & "DaysLate", txtDaysLate
   
   
   Dim sOptions As String
   SaveSetting "Esi2000", strRegApp, Me.Name & "_Printer", lblPrinter
   SaveSetting "Esi2000", strRegApp, Me.Name & "LateCalc", cmbLateCalc.ListIndex
End Sub

Private Sub GetOptions()
   Dim defaultDate As String
   defaultDate = Format(Date, "mm/dd/yyyy")
   
   Dim strRegApp As String
   strRegApp = GetRegistryAppTitle()
   
   cboStart = GetSetting("Esi2000", strRegApp, Me.Name & "StartDate", defaultDate)
   cboEnd = GetSetting("Esi2000", strRegApp, Me.Name & "EndDate", defaultDate)
   cboVendor = GetSetting("Esi2000", strRegApp, Me.Name & "Vendor", cboVendor.List(0))
   lblName = GetVendorName(cboVendor)
   txtDaysEarly = GetSetting("Esi2000", strRegApp, Me.Name & "DaysEarly", 0)
   txtDaysLate = GetSetting("Esi2000", strRegApp, Me.Name & "DaysLate", 0)
   
   On Error Resume Next
   lblPrinter = GetSetting("Esi2000", strRegApp, Me.Name & "_Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
   cmbLateCalc.ListIndex = GetSetting("Esi2000", strRegApp, Me.Name & "LateCalc", 0)
   
End Sub

Private Sub cboEnd_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboEnd_LostFocus()
   cboEnd = CheckDate(cboEnd)
End Sub

Private Sub cboStart_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboStart_LostFocus()
   cboStart = CheckDate(cboStart)
End Sub

