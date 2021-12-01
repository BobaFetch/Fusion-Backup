VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Orders (Report)"
   ClientHeight    =   4680
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CheckBox optDocList 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2880
      TabIndex        =   45
      Top             =   4080
      Width           =   735
   End
   Begin VB.CheckBox optSerDocList 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2880
      TabIndex        =   43
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      MaskColor       =   &H8000000F&
      Picture         =   "PurcPRp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
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
      Picture         =   "PurcPRp01a.frx":0592
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optAddr 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbCpy 
      ForeColor       =   &H00800000&
      Height          =   288
      ItemData        =   "PurcPRp01a.frx":0D40
      Left            =   6480
      List            =   "PurcPRp01a.frx":0D42
      TabIndex        =   13
      ToolTipText     =   "Copies To Print (Printed Only)"
      Top             =   1440
      Width           =   585
   End
   Begin VB.CheckBox optPvw 
      Caption         =   "PoView"
      Height          =   195
      Left            =   1440
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "PurcPRp01a.frx":0D44
      Height          =   320
      Left            =   3600
      Picture         =   "PurcPRp01a.frx":16B6
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Show Existing Purchase Orders"
      Top             =   840
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox cmbPon 
      Height          =   285
      Left            =   240
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cboEndPo 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "000000"
      ToolTipText     =   "Select Or Enter PO (Contains All But Canceled PO's)"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cboStartPo 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Select Or Enter PO (Contains All But Canceled PO's And 300 Max)"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   31
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PurcPRp01a.frx":2028
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PurcPRp01a.frx":21B2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.TextBox txtErl 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSrl 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.CheckBox optPck 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   3600
      Width           =   850
   End
   Begin VB.CheckBox optAlo 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   3120
      Width           =   850
   End
   Begin VB.CheckBox optOps 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   3336
      Width           =   850
   End
   Begin VB.CheckBox optCan 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   2844
      Width           =   850
   End
   Begin VB.CheckBox chkPoRemarks 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2880
      TabIndex        =   6
      Top             =   2640
      Width           =   850
   End
   Begin VB.CheckBox chkItemComments 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   850
   End
   Begin VB.CheckBox chkExtDesc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   850
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   3480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4680
      FormDesignWidth =   7140
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Part Document List"
      Height          =   285
      Index           =   18
      Left            =   240
      TabIndex        =   46
      ToolTipText     =   "Show Part Document List"
      Top             =   4080
      Width           =   2625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Service Part Document List"
      Height          =   285
      Index           =   17
      Left            =   240
      TabIndex        =   44
      ToolTipText     =   "Show Service Document List"
      Top             =   4320
      Width           =   2625
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2160
      TabIndex        =   42
      ToolTipText     =   "Nickname (Blank If A Range Is Selected)"
      Top             =   1560
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor (Single PO)"
      Height          =   252
      Index           =   16
      Left            =   240
      TabIndex        =   41
      Top             =   1560
      Width           =   1692
   End
   Begin VB.Label lblPO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   38
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Our Address"
      Height          =   288
      Index           =   14
      Left            =   240
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   1908
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   36
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copies "
      Height          =   252
      Index           =   15
      Left            =   5760
      TabIndex        =   35
      ToolTipText     =   "Copies To Print (Printed Only)"
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(If Different)"
      Height          =   252
      Index           =   13
      Left            =   3600
      TabIndex        =   30
      Top             =   1200
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      Height          =   252
      Index           =   12
      Left            =   4440
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      Height          =   252
      Index           =   11
      Left            =   4440
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Numbers"
      Enabled         =   0   'False
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving Information"
      Height          =   252
      Index           =   9
      Left            =   240
      TabIndex        =   26
      Top             =   3600
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allocations To MO's"
      Height          =   252
      Index           =   8
      Left            =   240
      TabIndex        =   25
      Top             =   3120
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Comments"
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Canceled Items"
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   2880
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Remarks"
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   2640
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2160
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending PO"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   1240
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting PO"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "PurcPRp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/18/04 realigned groups PrintReport
'12/30/04 Added purchasing conversion to report and corrected Revision
'10/19/05 Address Option to Standard PO, Fixed Fax formula
Option Explicit
Dim bOnLoad As Byte
Dim iLastPo As Integer

Dim sCoName As String
Dim sCoPhone As String
Dim sCoFax As String
Dim sCoTaxId As String
Dim sVEmail As String

Dim sCoAdr(5) As String
Dim lPoRange(26) As Long
Dim iUserLogo As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   Dim sCustomReport As String
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtSrl = "0"
   txtErl = "0"
   sCustomReport = GetCustomReport("prdpr01")
   If sCustomReport = "prdpr01.rpt" Then
      z1(14).Visible = True
      optAddr.Visible = True
   End If
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "pr01", sOptions)
   If Len(sOptions) > 0 Then
      chkExtDesc.Value = Val(Left(sOptions, 1))
      chkItemComments.Value = Val(Mid(sOptions, 2, 1))
      chkPoRemarks.Value = Val(Mid(sOptions, 3, 1))
      optCan.Value = Val(Mid(sOptions, 4, 1))
      optOps.Value = Val(Mid(sOptions, 5, 1))
      optAlo.Value = Val(Mid(sOptions, 6, 1))
      optPck.Value = Val(Mid(sOptions, 7, 1))
      optLot.Value = Val(Mid(sOptions, 8, 1))
      optSerDocList.Value = Val(Mid(sOptions, 9, 1))
      optDocList.Value = Val(Mid(sOptions, 10, 1))
   End If
   sOptions = GetSetting("Esi2000", "EsiProd", "pr01c", sOptions)
   If Val(sOptions) > 0 Then
      cmbCpy = Format(sOptions, "0")
   Else
      cmbCpy = "1"
   End If
   If optOps.Value = vbChecked Then optAlo.Value = vbChecked
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Ppr01", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   optAddr.Value = GetSetting("Esi2000", "EsiProd", "PoAddr,", optAddr.Value)
   
End Sub


Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(chkExtDesc.Value) _
              & RTrim(chkItemComments.Value) _
              & RTrim(chkPoRemarks.Value) _
              & RTrim(optCan.Value) _
              & RTrim(optOps.Value) _
              & RTrim(optAlo.Value) _
              & RTrim(optPck.Value) _
              & RTrim(optLot.Value) _
              & RTrim(optSerDocList.Value) _
              & RTrim(optDocList.Value)
              
   SaveSetting "Esi2000", "EsiProd", "pr01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "pr01c", Trim(cmbCpy)
   SaveSetting "Esi2000", "EsiProd", "Ppr01", lblPrinter
   SaveSetting "Esi2000", "EsiProd", "PoAddr,", optAddr.Value
   
End Sub

Private Sub cmbCpy_LostFocus()
   cmbCpy = CheckLen(cmbCpy, 1)
   If Val(cmbCpy) = 0 Then cmbCpy = "1"
   
End Sub

Private Sub cmbPon_Change()
   cboStartPo = cmbPon
   
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

Private Sub cmdVew_Click()
   '    If cmdVew.Value = True Then
   '        PurcPOtree.Show
   '        cmdVew.Value = False
   '    End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetCompany 1
      FillCombo
      ' To check if we need to use company logo
      GetUseLogo
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   cboEndPo.ToolTipText = "Range Of PO's (max 25) Print Only"
   bOnLoad = 1
   GetOptions
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If lblPO = "" Then
      FormUnload
   Else
      PurcPRe02a.lblPO = lblPO
      PurcPRe02a.Show
   End If
   Set PurcPRp01a = Nothing
   
End Sub

Private Sub optAddr_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optAlo_Click()
   If optAlo.Value = vbUnchecked Then optOps.Value = vbUnchecked
   
End Sub

Private Sub optAlo_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optCan_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub chkExtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub chkExtDesc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub chkitemcomments_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub chkitemcomments_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optLot_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optLot_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optOps_Click()
   If optOps.Value = vbChecked Then optAlo.Value = vbChecked
   
End Sub

Private Sub optOps_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optOps_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPck_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optPck_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   If Trim(cboEndPo) > Trim(cboStartPo) Then
      optPrn.Enabled = False
      optDis.Enabled = False
      GetPrintRange
      'PrintRange
      PrintReport
   Else
      PrintReport
   End If
   
End Sub

Private Sub optPvw_Click()
   If optPvw.Value = vbChecked Then
      optPvw.Value = vbUnchecked
      On Error Resume Next
      cboStartPo.SetFocus
   End If
   
End Sub

Private Sub chkporemarks_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub chkporemarks_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cboendpo_LostFocus()
   cboEndPo = CheckLen(cboEndPo, 6)
   cboEndPo = Format(Abs(Val(cboEndPo)), "000000")
   If cboEndPo <> cboStartPo Then lblVnd = ""
   
End Sub


Private Sub txtErl_LostFocus()
   txtErl = CheckLen(txtErl, 2)
   txtErl = Format(Abs(Val(txtErl)), "#0")
   
End Sub

Private Sub cbostartpo_Click()
   GetPoRange
   
End Sub

Private Sub cbostartpo_LostFocus()
   cboStartPo = CheckLen(cboStartPo, 6)
   cboStartPo = Format(Abs(Val(cboStartPo)), "000000")
   GetPoRange
   
End Sub

Private Sub txtSrl_LostFocus()
   txtSrl = CheckLen(txtSrl, 2)
   txtSrl = Format(Abs(Val(txtSrl)), "#0")
   
End Sub
Private Sub PrintReport()
   Dim bForm As Byte
   clsADOCon.ExecuteSQL "UPDATE ComnTable SET CURPONUMBER=" & cboStartPo & " "
   bForm = GetPrintedForm("purchase order")
   On Error GoTo Ppr01
   
   Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    
    sCustomReport = GetCustomReport("prdpr01")
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False

   
    aFormulaName.Add "Company"
    aFormulaName.Add "Phone"
    aFormulaName.Add "Fax"
    aFormulaName.Add "ShowComt"
    aFormulaName.Add "ShowRem"
    aFormulaName.Add "CoAddress1"
    aFormulaName.Add "CoAddress2"
    aFormulaName.Add "CoAddress3"
    aFormulaName.Add "CoAddress4"
    aFormulaName.Add "PoNumber"
    aFormulaName.Add "ShowExtDesc"
    aFormulaName.Add "ShowCanceledItems"
    aFormulaName.Add "ShowMoAllocations"
    aFormulaName.Add "ShowOpComments"
    aFormulaName.Add "ShowRecInfo"
    aFormulaName.Add "ShowServPartDoc"
    aFormulaName.Add "ShowPartDoc"
    aFormulaName.Add "ShowOurAddress"
    aFormulaName.Add "ResaleNumber"
    
    
    aFormulaValue.Add CStr("'" & CStr(Co.Name) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Phone) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Fax) & "'")
    aFormulaValue.Add CStr("'" & CStr(chkItemComments) & "'")
    aFormulaValue.Add CStr("'" & CStr(chkPoRemarks) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(1)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(2)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(3)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(4)) & "'")
    aFormulaValue.Add CStr(cboStartPo)
    aFormulaValue.Add CStr("'" & CStr(chkExtDesc) & "'")
    aFormulaValue.Add CStr("'" & CStr(optCan) & "'")
    aFormulaValue.Add CStr("'" & CStr(optAlo) & "'")
    aFormulaValue.Add CStr("'" & CStr(optOps) & "'")
    aFormulaValue.Add CStr("'" & CStr(optPck) & "'")
    aFormulaValue.Add CStr("'" & CStr(optSerDocList) & "'")
    aFormulaValue.Add CStr("'" & CStr(optDocList) & "'")

    
  
   If (iUserLogo = 1) Then
    aFormulaValue.Add CStr("'" & CStr(0) & "'")

   Else
    aFormulaValue.Add CStr("'" & CStr(optAddr) & "'")
   End If
   
    aFormulaValue.Add CStr("'" & CStr(GetPreferenceValue("RESALENUMBER", True)) & "'")
    aFormulaName.Add "ShowOurLogo"
    aFormulaValue.Add CStr("'" & CStr(iUserLogo) & "'")
    ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue


    If cboStartPo <> cboEndPo Then
        sSql = "{PohdTable.PONUMBER}>=" & Val(cboStartPo) & " AND {PohdTable.PONUMBER}<=" & Val(cboEndPo) & " "
    Else
        sSql = "{PohdTable.PONUMBER}=" & Val(cboStartPo) & " "
    End If
   If optCan.Value = vbUnchecked Then
      sSql = sSql & "AND {PoitTable.PITYPE}<>16 "
   End If
    cCRViewer.SetReportSelectionFormula sSql
    
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
   
   If optPrn Then
      MarkAsPrinted Val(cboStartPo)
   End If
   
   cCRViewer.OpenCrystalReportObject Me, aFormulaName, Val(cmbCpy)
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   
   MouseCursor 0
   Exit Sub
   
Ppr01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Ppr01a
Ppr01a:
   DoModuleErrors Me
   
End Sub



Private Sub GetUseLogo()
    Dim RdoLogo As ADODB.Recordset
    Dim bRows As Boolean
    ' Assumed that COMREF is 1 all the time
    sSql = "SELECT ISNULL(COLUSELOGO, 0) as COLUSELOGO FROM ComnTable WHERE COREF = 1"
    bRows = clsADOCon.GetDataSet(sSql, RdoLogo, ES_FORWARD)

    If bRows Then
        With RdoLogo
            iUserLogo = !COLUSELOGO
        End With
        'RdoLogo.Close
        ClearResultSet RdoLogo
    End If
End Sub



Private Sub MarkAsPrinted(lPonumber As Long)
   Dim sMark As String
   On Error Resume Next
   sMark = "UPDATE PohdTable SET POPRINTED='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
           & "WHERE PONUMBER=" & lPonumber & " "
   clsADOCon.ExecuteSQL sMark
   
End Sub

Private Sub FillCombo()
   Dim b As Byte
   'Dim sYear As String
   On Error GoTo DiaErr1
   For b = 1 To 8
      AddComboStr cmbCpy.hwnd, Format$(b, "0")
   Next
   AddComboStr cmbCpy.hwnd, Format$(b, "0")
   
   'sYear = Format(Now - 730, "yyyy-mm-dd")
   'sSql = "Qry_FillPurchaseOrders '" & sYear & "'"
   sSql = "Qry_FillPurchaseOrders '" & DateAdd("yyyy", -2, Now) & "'"
   LoadNumComboBox cboStartPo, "000000"
   If Not bSqlRows Then _
      MsgBox "There Are No Purchase Orders.", vbInformation, Caption
   If Trim(lblPO) = "" Then
      cboStartPo = cboStartPo.List(0)
   Else
      cboStartPo = lblPO
      cboEndPo = lblPO
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetPoRange()
   Dim RdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
   cboEndPo.Clear
   sSql = "SELECT PONUMBER FROM PohdTable WHERE " _
          & "(PONUMBER>=" & Val(cboStartPo) & " AND POCAN=0)"
   LoadNumComboBox cboEndPo, "000000"
   If bSqlRows Then cboEndPo = cboEndPo.List(0)
   If cboEndPo <> cboStartPo Then
      lblVnd = "*** Range Selected ***"
      sVEmail = ""
   Else
      sSql = "SELECT POVENDOR,VENICKNAME,VEEMAIL FROM PohdTable," _
             & "VndrTable WHERE (PONUMBER=" & Val(cboStartPo) & " AND " _
             & "POVENDOR=VEREF)"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
      If bSqlRows Then
         With RdoVnd
            lblVnd = "" & Trim(!VENICKNAME)
            sVEmail = "" & Trim(!VEEMAIL)
            .Cancel
         End With
         ClearResultSet RdoVnd
      Else
         lblVnd = "*** Range Selected ***"
         sVEmail = ""
      End If
      
   End If
   Set RdoVnd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getporange"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

' TODO: convert to CR11
Private Sub PrintRange()
   Dim iList As Integer
   On Error GoTo Ppr01
'   clsADOCon.ExecuteSQL "UPDATE ComnTable SET CURPONUMBER=" & cboStartPo & " "
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "Company='" & Co.Name & "'"
'   MDISect.Crw.Formulas(1) = "Phone='" & Co.Phone & "'"
'   MDISect.Crw.Formulas(2) = "Fax='" & Co.Fax & "'"
'   sCustomReport = GetCustomReport("prdpr01")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   If chkExtDesc.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.1;T;;;"
'
'   End If
'   If sCustomReport <> "awipr01.rpt" Then
'      If optAlo.Value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(2) = "GROUPFTR.1.0;F;;;"
'         MDISect.Crw.SectionFormat(3) = "GROUPFTR.1.1;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(2) = "GROUPFTR.1.0;T;;;"
'         If optOps.Value = vbUnchecked Then
'            MDISect.Crw.SectionFormat(3) = "GROUPFTR.1.1;F;;;"
'         Else
'            MDISect.Crw.SectionFormat(3) = "GROUPFTR.1.1;T;;;"
'         End If
'      End If
'   Else
'      If optOps.Value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(4) = "GROUPFTR.1.1;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(4) = "GROUPFTR.1.1;T;;;"
'      End If
'   End If
'   If chkItemComments.Value = vbUnchecked Then
'      MDISect.Crw.Formulas(3) = "ShowComt='0'"
'   Else
'      MDISect.Crw.SectionFormat(5) = "GROUPFTR.0.0;T;;;"
'      MDISect.Crw.SectionFormat(5) = "GROUPFTR.0.1;T;;;"
'      MDISect.Crw.Formulas(3) = "ShowComt='1'"
'   End If
'   If chkPoRemarks.Value = vbUnchecked Then
'      MDISect.Crw.Formulas(4) = "ShowRem='0'"
'   Else
'      MDISect.Crw.Formulas(4) = "ShowRem='1'"
'   End If
'   '2/21/01
'   If optPck.Value = vbUnchecked Then
'      MDISect.Crw.Formulas(5) = "ShowRecd='0'"
'
'   Else
'      MDISect.Crw.Formulas(5) = "ShowRecd='1'"
'   End If
'   '10/19/05
'   If sCustomReport = "prdpr01.rpt" Then
'      MDISect.Crw.Formulas(6) = "CoAddress1='" & Co.Addr(1) & "'"
'      MDISect.Crw.Formulas(7) = "CoAddress2='" & Co.Addr(2) & "'"
'      MDISect.Crw.Formulas(8) = "CoAddress3='" & Co.Addr(3) & "'"
'      MDISect.Crw.Formulas(9) = "CoAddress4='" & Co.Addr(4) & "'"
'      If optAddr.Value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(7) = "REPORTHDR.0.1;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(7) = "REPORTHDR.0.1;T;;;"
'      End If
'   End If
'   '7/21/99
'   For iList = 1 To iLastPo
'      MDISect.Crw.Formulas(10) = "PoNumber='" & Format$(lPoRange(iList), "000000") & "'"
'      sSql = "{PohdTable.PONUMBER}=" & lPoRange(iList) & " "
'      If optCan.Value = vbUnchecked Then
'         sSql = sSql & "AND {PoitTable.PITYPE}<>16 "
'      End If
'      MarkAsPrinted lPoRange(iList)
'      MDISect.Crw.CopiesToPrinter = Val(cmbCpy)
'      MDISect.Crw.SelectionFormula = sSql
'      SetCrystalAction Me
'   Next
'   optPrn.Enabled = True
'   optDis.Enabled = True
'   MouseCursor 0
   Exit Sub
   
Ppr01:
   sProcName = "printrange"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   optPrn.Enabled = True
   optDis.Enabled = True
   DoModuleErrors Me
   
End Sub

Private Sub GetPrintRange()
   Dim RdoPpr As ADODB.Recordset
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   Erase lPoRange
   sSql = "SELECT PONUMBER,POCAN FROM PohdTable WHERE " _
          & "PONUMBER BETWEEN " & Val(cboStartPo) & " AND " _
          & Val(cboEndPo) & " AND POCAN=0 "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPpr)
   If bSqlRows Then
      With RdoPpr
         Do Until .EOF
            iList = iList + 1
            lPoRange(iList) = !PONUMBER
            If iList > 25 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoPpr
      End With
   End If
   iLastPo = iList
   cboEndPo = cboEndPo.List(0)
   Set RdoPpr = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getprintrange"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
