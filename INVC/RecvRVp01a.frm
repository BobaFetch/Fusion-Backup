VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RecvRVp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receiving Log By Date"
   ClientHeight    =   4305
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   3075
   End
   Begin VB.ComboBox cmbGroupBy 
      Height          =   315
      ItemData        =   "RecvRVp01a.frx":0000
      Left            =   2040
      List            =   "RecvRVp01a.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RecvRVp01a.frx":0030
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Chars Or Blank For All"
      Top             =   1200
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "RecvRVp01a.frx":07DE
      Height          =   315
      Left            =   5280
      Picture         =   "RecvRVp01a.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1200
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.Frame Z2 
      Height          =   615
      Left            =   2040
      TabIndex        =   24
      ToolTipText     =   "Select Types - Right And Left Arrow Keys"
      Top             =   1560
      Width           =   3255
      Begin VB.OptionButton optShow 
         Caption         =   "Service"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   4
         ToolTipText     =   "Service Part Type 7 Only"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Raw"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Part Types 4 And 5 Only"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optShow 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Show All Part Numbers"
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Tag             =   "4"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Tag             =   "4"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5760
      TabIndex        =   23
      Top             =   450
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "RecvRVp01a.frx":0E62
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "RecvRVp01a.frx":0FEC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5760
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   90
      Width           =   1065
   End
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Tag             =   "3"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.CheckBox optMoa 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4305
      FormDesignWidth =   6885
   End
   Begin VB.Label Label1 
      Caption         =   "Group report by"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   29
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   28
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Part Types"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   25
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   9
      Left            =   5760
      TabIndex        =   21
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   20
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Types"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   16
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Allocations"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
End
Attribute VB_Name = "RecvRVp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sType As String * 2
   If txtTyp = "ALL" Then txtTyp = ""
   sType = txtTyp
   'Save by Menu Option
   sOptions = RTrim(optExt.Value) _
              & RTrim(optMoa.Value) _
              & RTrim(optLot.Value) _
              & sType & Trim(str(cmbGroupBy.ListIndex))
              
   SaveSetting "Esi2000", "EsiProd", "rc01", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "rc01", sOptions)
   If Len(sOptions) > 0 Then
      optExt.Value = Val(Left(sOptions, 1))
      optMoa.Value = Val(Mid(sOptions, 2, 1))
      optLot.Value = Val(Mid(sOptions, 3, 1))
      txtTyp = Trim(Mid(sOptions, 4, 2))
      If Len(sOptions) > 5 Then cmbGroupBy.ListIndex = Val(Mid(sOptions, 6, 1)) Else cmbGroupBy.ListIndex = 1
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   optVew.Value = vbChecked
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

Private Sub Combo1_Change()

End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   FillPartCombo cmbPrt
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "01/01/" & Right(txtEnd, 4)
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
   Set RecvRVp01a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sPart As String
   Dim sType As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   If Len(Trim(txtTyp)) > 0 Then
      sType = Trim(txtTyp)
   Else
      sType = "ALL"
   End If
   If cmbPrt <> "ALL" Then sPart = Compress(cmbPrt)
   If IsDate(txtBeg) Then
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   Else
      sBegDate = "1995,01,01"
   End If
   If IsDate(txtEnd) Then
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   Else
      sEndDate = "2024,12,31"
   End If
   MouseCursor 13
   On Error GoTo Prc01
   sCustomReport = GetCustomReport("prdrc01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowExDesc"
    aFormulaName.Add "ShowMoa"
    aFormulaName.Add "GroupByMonthYear"
    
    aFormulaValue.Add CStr("'Receipts From " & CStr(txtBeg & " Ending " & txtEnd & " And Vendor Types " & sType) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add optExt.Value
    aFormulaValue.Add optMoa.Value
    If cmbGroupBy.ListIndex = 1 Then aFormulaValue.Add CStr("'Y'") Else aFormulaValue.Add CStr(("'N'"))
    
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   sSql = "{PoitTable.PIADATE} in Date(" & sBegDate & ") to Date(" & sEndDate & ") AND " _
          & "{PartTable.PARTREF} LIKE '" & sPart & "*' "
   If sType <> "ALL" Then sSql = sSql & " AND {VndrTable.VETYPE} Like '" & sType & "*' "
   If optShow(1).Value = True Then
      sSql = sSql & " AND ({PartTable.PALEVEL}=4 OR {PartTable.PALEVEL}=5) "
   Else
      If optShow(2).Value = True Then _
         sSql = sSql & " AND {PartTable.PALEVEL}=7"
   End If
   sSql = sSql & " and {PoitTable.PITYPE} IN [15, 17]"
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   cCRViewer.ShowGroupTree False

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
Prc01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optExt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optLot_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optLot_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optMoa_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optMoa_KeyPress(KeyAscii As Integer)
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
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If txtPrt = "" Then txtPrt = "ALL"
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   
End Sub

Private Sub txtTyp_LostFocus()
   txtTyp = CheckLen(txtTyp, 2)
   txtTyp = Compress(txtTyp)
   
End Sub
