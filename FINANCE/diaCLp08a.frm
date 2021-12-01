VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaCLp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory movement From WIP (Report)"
   ClientHeight    =   4260
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4260
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1680
      TabIndex        =   34
      Tag             =   "9"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox typ 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   33
      Top             =   1440
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   32
      Top             =   1440
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   31
      Top             =   1440
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   30
      Top             =   1440
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox chkTotCost 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   28
      Top             =   3840
      Width           =   200
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1680
      TabIndex        =   25
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   24
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox chkProj 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkStdProd 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkJobs 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   5950
      Picture         =   "diaCLp08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Print The Report"
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   5400
      Picture         =   "diaCLp08a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Display The Report"
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox chkJGL 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   4605
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkSummary 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   5400
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkDsc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   3240
      Width           =   200
   End
   Begin VB.CheckBox chkExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   3525
      Width           =   200
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Tag             =   "8"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ComboBox txtJDate 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Tag             =   "4"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5400
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   12
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
      PictureUp       =   "diaCLp08a.frx":0308
      PictureDn       =   "diaCLp08a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   3960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4260
      FormDesignWidth =   6900
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Code"
      Height          =   285
      Index           =   14
      Left            =   120
      TabIndex        =   37
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL) "
      Height          =   285
      Index           =   15
      Left            =   3000
      TabIndex        =   36
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   35
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Total Cost detail"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   29
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   255
      Index           =   13
      Left            =   480
      TabIndex        =   27
      Top             =   600
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   26
      Top             =   960
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Movement for Standard Products"
      Height          =   285
      Index           =   12
      Left            =   360
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Movement For Projects"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Movement For Jobs"
      Height          =   285
      Index           =   8
      Left            =   4560
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Journal to G.L"
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   20
      Top             =   4605
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Only"
      Height          =   285
      Index           =   4
      Left            =   4800
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   3525
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class"
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   15
      Top             =   2400
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL) "
      Height          =   285
      Index           =   10
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal Date :"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "diaCLp08a"
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
' diaCLp08a - inventory movement to WIP
'
' Notes:
'
' Created: 9/6/08
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbCls_LostFocus()
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
End Sub


Private Sub cmbCde_LostFocus()
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
End Sub


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
      FillProductClasses Me
      Me.cmbCls = "ALL"
      Me.cmbCde = "ALL"
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   'txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   'txtBeg = Format(txtEnd, "mm/01/yy")
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaCLp08a = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sType As String
   
   On Error GoTo whoops
   
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   If Trim(cmbCde) = "" Then cmbCde = "ALL"

   
   'get custom report name if one has been defined
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("fincl08.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   'pass formulas
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "PartClass"
   aFormulaName.Add "ShowPartDesc"
   aFormulaName.Add "ShowExtDesc"
   aFormulaName.Add "ShowTotDet"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'From " & CStr(txtBeg & " Through " & txtEnd & " for Classes " & cmbCls) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbCls) & "'")
   aFormulaValue.Add chkDsc
   aFormulaValue.Add chkExt
   aFormulaValue.Add chkTotCost
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
'   'pass Crystal SQL if required
'   sSql = "{InvaTable.INADATE} >= #" & txtBeg & "# AND {InvaTable.INADATE} <= #" & txtEnd & "# AND {InvaTable.INTYPE} = 6"
'
'   If cmbCde <> "ALL" Then
'      sSql = sSql & " AND {PartTable.PACLASS} = '" & cmbCls & "'"
'   End If
'
'   If cmbCls <> "ALL" Then
'      sSql = sSql & " AND {PartTable.PACLASS} = '" & cmbCls & "'"
'   End If
'
'   Dim b As Integer
'   sType = ""
'   For b = 1 To 4
'      If typ(b) = vbChecked Then
'         sType = sType & CStr(b) & IIf(b = 4, "", ",")
'      End If
'   Next
'
'   If (sType <> "") Then
'      sSql = sSql & " AND {PartTable.PALEVEL} in [" & sType & "]"
'   End If
'
'   cCRViewer.SetReportSelectionFormula sSql
   
   ' set the sub sql variable pass the sub report name
'   cCRViewer.SetSubRptSelFormula "sr_ClGL_Acc_Sum", sSql
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   
   ' report parameter
   aRptPara.Add CStr(txtBeg)
   aRptPara.Add CStr(txtEnd)
   aRptPara.Add CStr(typ(1))
   aRptPara.Add CStr(typ(2))
   aRptPara.Add CStr(typ(3))
   aRptPara.Add CStr(typ(4))
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   
   ' Set report parameter
   
   ' set the sub sql variable pass the sub report name
   cCRViewer.SetSubRptDBParameters "sr_ClGL_Acc_Sum", aRptPara, aRptParaType
   
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType    'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   Exit Sub
   
whoops:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub PrintReport1()
   Dim sCustomReport As String
   On Error GoTo whoops
   
   'setmdireportsizemdisect
   
   'get custom report name if one has been defined
   sCustomReport = GetCustomReport("fincl08.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   'pass formulas
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='Jornal Date " & txtJDate & " for Classes " & cmbCls & "'"
   MdiSect.crw.Formulas(3) = "PartClass='" & cmbCls & "'"
   
   MdiSect.crw.Formulas(4) = "ShowStdProducts=" & chkStdProd
   MdiSect.crw.Formulas(5) = "ShowProject=" & chkProj
   MdiSect.crw.Formulas(6) = "ShowJobs=" & chkJobs
   
   MdiSect.crw.Formulas(7) = "ShowPartDesc=" & chkDsc
   MdiSect.crw.Formulas(8) = "ShowExtDesc=" & chkExt
   MdiSect.crw.Formulas(9) = "ShowSummary=" & chkSummary
   MdiSect.crw.Formulas(10) = "ShowGLTransferJournal=" & chkJGL
   
   'pass Crystal SQL if required
   sSql = ""
   MdiSect.crw.SelectionFormula = sSql
   'setcrystalaction me
   Exit Sub
   
whoops:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtBeg.Text) & Trim(txtEnd.Text)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
   
     If dToday >= 20 Then
       txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
       txtBeg = Format(txtEnd, "mm/01/yy")
     Else
       txtBeg = Mid(sOptions, 1, 8)
       txtEnd = Mid(sOptions, 9, 8)
     End If
     
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub

Private Sub txtJDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtJDate_LostFocus()
   txtJDate = CheckDate(txtJDate)
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub


