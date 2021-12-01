VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimate (Report)"
   ClientHeight    =   3900
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3900
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optComments 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "EstiESp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
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
      Picture         =   "EstiESp01a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optPrc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.CheckBox optSta 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   1080
      Value           =   1  'Checked
      Width           =   252
   End
   Begin VB.CheckBox optCbd 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox cmbBid 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Select Or Enter A Bid Number"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5880
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "EstiESp01a.frx":0938
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
         Picture         =   "EstiESp01a.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   3720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3900
      FormDesignWidth =   6975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   165
      Index           =   6
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Width           =   1665
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   5400
      TabIndex        =   25
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discounts"
      Height          =   165
      Index           =   9
      Left            =   240
      TabIndex        =   24
      Top             =   3120
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   22
      Top             =   2100
      Width           =   1665
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   2100
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   20
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Breakdown"
      Height          =   165
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   2880
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext Part Descriptions"
      Height          =   165
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblTyp 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   16
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblNik 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblCust 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   1770
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Top             =   1440
      Width           =   950
   End
   Begin VB.Label lblCls 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   675
      TabIndex        =   9
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1665
   End
End
Attribute VB_Name = "EstiESp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim bGoodBid As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Function GetTheBid() As Byte
   Dim AdoBid As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT BIDREF,BIDNUM,BIDPRE,BIDCLASS,BIDPART,BIDCUST," _
          & "BIDDATE,BIDRFQ,BIDCANCELED,BIDACCEPTED,BIDCOMPLETE," _
          & "CUREF,CUNICKNAME,CUNAME,PARTREF,PARTNUM " _
          & "FROM EstiTable,CustTable,PartTable WHERE (BIDCUST=CUREF " _
          & "AND BIDPART=PARTREF) AND BIDREF=" & Val(cmbBid) & " "
   
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoBid, ES_FORWARD)
   If bSqlRows Then
      With AdoBid
         GetTheBid = 1
         If !BIDCANCELED = 1 Then
            lblStatus.Caption = "Canceled"
         ElseIf !BIDACCEPTED = 1 Then
            lblStatus.Caption = "Accepted"
         ElseIf !BIDCOMPLETE = 1 Then
            lblStatus.Caption = "Complete"
         Else
            lblStatus.Caption = "Incomplete"
         End If
         lblCls = "" & Trim(!BIDPRE)
         lblNik = "" & Trim(!CUNICKNAME)
         lblTyp = "" & Trim(!BidClass)
         lblCust = "" & Trim(!CUNAME)
         lblDate = "" & Format(!BIDDATE, "mm/dd/yyyy")
         lblPrt = "" & Trim(!PartNum)
         ClearResultSet AdoBid
      End With
   Else
      GetTheBid = 0
      lblCls = ""
      lblNik = ""
      lblTyp = ""
      lblCust = ""
      lblDate = ""
      lblPrt = ""
      MsgBox "This Bid Does Not Exist Or Doesn't Qualify." & vbCrLf _
         & "See Help For Instructions On Marking A Bid Not Accepted.", _
         vbInformation, Caption
   End If
   Set AdoBid = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthebid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function





Private Sub cmbBid_Click()
   bGoodBid = GetTheBid()
   
End Sub

Private Sub cmbBid_LostFocus()
   cmbBid = CheckLen(cmbBid, 6)
   cmbBid = Format(Abs(Val(cmbBid)), "000000")
   bGoodBid = GetTheBid()
   
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
   Dim AdoCmb As ADODB.Recordset
   Dim iList As Integer
   On Error GoTo DiaErr1
   sSql = "SELECT BIDREF,BIDPRE FROM EstiTable " _
          & "WHERE BIDCANCELED=0 ORDER BY BIDREF DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoCmb, ES_FORWARD)
   If bSqlRows Then
      With AdoCmb
         cmbBid = Format(!BIDREF, "000000")
         lblCls = "" & Trim(!BIDPRE)
         Do Until .EOF
            iList = iList + 1
            If iList > 300 Then Exit Do
            AddComboStr cmbBid.hwnd, Format(!BIDREF, "000000")
            .MoveNext
         Loop
         ClearResultSet AdoCmb
      End With
   End If
   If cmbBid.ListCount > 0 Then bGoodBid = GetTheBid()
   Set AdoCmb = Nothing
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
      FillCombo
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set EstiESp01a = Nothing
   
End Sub

Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   sCustomReport = GetCustomReport("enges01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "ShowExDesc"
   aFormulaName.Add "ShowCbd"
   aFormulaName.Add "ShowPrc"
   aFormulaName.Add "ShowComments"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'....'")
   aFormulaValue.Add optDsc.value
   aFormulaValue.Add optCbd.value
   aFormulaValue.Add optPrc.value
   aFormulaValue.Add optComments.value
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{EstiTable.BIDREF}=" & Val(cmbBid) & " "
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
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
   sOptions = RTrim(optDsc.value) _
              & RTrim(optCbd.value) & RTrim(optComments.value)
   SaveSetting "Esi2000", "EsiEngr", "es01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "Pes01", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "es01", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.value = Val(Left(sOptions, 1))
      optCbd.value = Val(Mid(sOptions, 2, 1))
      If Len(sOptions) > 2 Then optComments.value = Val(Mid(sOptions, 3, 1)) Else optComments.value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Pes01", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


Private Sub optCbd_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

