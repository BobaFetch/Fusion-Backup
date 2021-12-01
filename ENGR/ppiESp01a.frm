VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ppiESp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimate (Report)"
   ClientHeight    =   4515
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4515
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "ppiESp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
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
      Picture         =   "ppiESp01a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optPic 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   3480
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optPrc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
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
      Left            =   5040
      TabIndex        =   25
      Top             =   1080
      Value           =   1  'Checked
      Width           =   252
   End
   Begin VB.CheckBox optCbd 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ppiESp01a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ppiESp01a.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4515
      FormDesignWidth =   6975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   168
      Index           =   12
      Left            =   240
      TabIndex        =   32
      Top             =   3120
      Visible         =   0   'False
      Width           =   1668
   End
   Begin VB.Label lblPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   2040
      TabIndex        =   31
      ToolTipText     =   "Uses PartTable.PAPICLINK1. Double Click To View Attachment (If Any)"
      Top             =   3840
      Width           =   4812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture File"
      Height          =   168
      Index           =   11
      Left            =   240
      TabIndex        =   30
      ToolTipText     =   "Uses PartTable.PAPICLINK1"
      Top             =   3960
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Printed Estimates Only"
      Height          =   252
      Index           =   10
      Left            =   3000
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   2628
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Picture"
      Height          =   168
      Index           =   6
      Left            =   240
      TabIndex        =   28
      Top             =   3480
      Width           =   1668
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   5400
      TabIndex        =   27
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discounts"
      Height          =   168
      Index           =   9
      Left            =   240
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   24
      Top             =   2100
      Width           =   1665
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   23
      Top             =   2100
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   22
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Breakdown"
      Enabled         =   0   'False
      Height          =   168
      Index           =   5
      Left            =   480
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext Part Descriptions"
      Height          =   165
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblTyp 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   18
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblNik 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblCust 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   1770
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label cmbCls 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1665
   End
End
Attribute VB_Name = "ppiESp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/3/06 Format and code for QWIK and FULL Bids
'3/2/06 Add print a picture feature
'3/10/06 Added group formatting per Larry
'4/19/06 Changed Bid Combo Sort and Query
Option Explicit
Dim bOnLoad As Byte
Dim bPrintPic As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Function GetTheBid() As Byte
   Dim RdoBid As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT BIDREF,BIDNUM,BIDPRE,BIDCLASS,BIDPART,BIDCUST," _
          & "BIDDATE,BIDRFQ,BIDCANCELED,BIDACCEPTED,BIDCOMPLETE," _
          & "CUREF,CUNICKNAME,CUNAME,PARTREF,PARTNUM,PAPICLINK1 " _
          & "FROM EstiTable,CustTable,PartTable WHERE (BIDCUST=CUREF " _
          & "AND BIDPART=PARTREF) AND BIDREF=" & Val(cmbBid) & " "
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBid, ES_FORWARD)
   If bSqlRows Then
      With RdoBid
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
         cmbCls = "" & Trim(!BIDPRE)
         lblNik = "" & Trim(!CUNICKNAME)
         lblTyp = "" & Trim(!BidClass)
         lblCust = "" & Trim(!CUNAME)
         lblDate = "" & Format(!BIDDATE, "mm/dd/yy")
         lblPrt = "" & Trim(!PartNum)
         lblPicture = "" & Trim(!PAPICLINK1)
         ClearResultSet RdoBid
      End With
   Else
      GetTheBid = 0
      cmbCls = ""
      lblNik = ""
      lblTyp = ""
      lblCust = ""
      lblDate = ""
      lblPrt = ""
      MsgBox "This Bid Does Not Exist Or Doesn't Qualify." & vbCrLf _
         & "See Help For Instructions On Marking A Bid Not Accepted.", _
         vbInformation, Caption
   End If
   Set RdoBid = Nothing
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
   cmbBid.Clear
   'On Error GoTo DiaErr1
   FillEstimateCombo Me, "ALL"
   If cmbBid.ListCount > 0 Then
      cmbBid = cmbBid.List(0)
      bGoodBid = GetTheBid()
   End If
   If cmbBid.ListCount > 0 Then bGoodBid = GetTheBid()
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
   SetMdiReportsize MDISect
   If lblTyp <> "FULL" Then
      sCustomReport = GetCustomReport("enges01")
      MDISect.Crw.ReportFileName = sReportPath & sCustomReport
      MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
      MDISect.Crw.Formulas(1) = "Includes='" & "" & "...'"
      sSql = "{EstiTable.BIDREF}=" & Val(cmbBid) & " "
      If optCbd.value = vbUnchecked Then
         MDISect.Crw.SectionFormat(0) = "REPORTFTR.0.0;F;;;"
      Else
         MDISect.Crw.SectionFormat(0) = "REPORTFTR.0.0;T;;;"
      End If
      If optPrc.value = vbUnchecked Then
         MDISect.Crw.SectionFormat(1) = "DETAIL.0.0;F;;;"
      Else
         MDISect.Crw.SectionFormat(1) = "DETAIL.0.0;T;;;"
      End If
      If optDsc.value = vbUnchecked Then
         MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;F;;;"
      Else
         MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;T;;;"
      End If
   Else
      sCustomReport = GetCustomReport("ppienges01a")
      MDISect.Crw.ReportFileName = sReportPath & sCustomReport
      MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
      MDISect.Crw.Formulas(1) = "Includes='" & "" & "...'"
      sSql = "{EstiTable.BIDREF}=" & Val(cmbBid) & " "
      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.1;T;;;"
      If optDsc.value = vbUnchecked Then
         MDISect.Crw.SectionFormat(1) = "GROUPHDR.0.2;F;;;"
      Else
         MDISect.Crw.SectionFormat(1) = "GROUPHDR.0.2;T;;;"
      End If
      If optPrc.value = vbUnchecked Then
         MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;F;;;"
      Else
         MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;T;;;"
      End If
      If optCmt.value = vbUnchecked Then
         MDISect.Crw.SectionFormat(3) = "GROUPFTR.0.1;F;;;"
      Else
         MDISect.Crw.SectionFormat(3) = "GROUPFTR.0.1;T;;;"
      End If
   End If
   MDISect.Crw.SelectionFormula = sSql
   SetCrystalAction Me
   If bPrintPic = 1 Then
      Sleep 2000
      OpenWebPage lblPicture, "print"
   End If
   MouseCursor 0
   Exit Sub

DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub


'Private Sub PrintReport()
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   SetMdiReportsize MDISect
'   If lblTyp <> "FULL" Then
'      sCustomReport = GetCustomReport("enges01")
'      MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'      MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'      MDISect.Crw.Formulas(1) = "Includes='" & "" & "...'"
'      sSql = "{EstiTable.BIDREF}=" & Val(cmbBid) & " "
'      If optCbd.value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(0) = "REPORTFTR.0.0;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(0) = "REPORTFTR.0.0;T;;;"
'      End If
'      If optPrc.value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(1) = "DETAIL.0.0;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(1) = "DETAIL.0.0;T;;;"
'      End If
'      If optDsc.value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;T;;;"
'      End If
'   Else
'      sCustomReport = GetCustomReport("ppienges01a")
'      MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'      MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'      MDISect.Crw.Formulas(1) = "Includes='" & "" & "...'"
'      sSql = "{EstiTable.BIDREF}=" & Val(cmbBid) & " "
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.1;T;;;"
'      If optDsc.value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(1) = "GROUPHDR.0.2;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(1) = "GROUPHDR.0.2;T;;;"
'      End If
'      If optPrc.value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;T;;;"
'      End If
'      If optCmt.value = vbUnchecked Then
'         MDISect.Crw.SectionFormat(3) = "GROUPFTR.0.1;F;;;"
'      Else
'         MDISect.Crw.SectionFormat(3) = "GROUPFTR.0.1;T;;;"
'      End If
'   End If
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
'   If bPrintPic = 1 Then
'      Sleep 2000
'      OpenWebPage lblPicture, "print"
'   End If
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'













Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = RTrim(optDsc.value) _
              & RTrim(optCbd.value)
   SaveSetting "Esi2000", "EsiEngr", "es01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "Pes01", lblPrinter
   SaveSetting "Esi2000", "EsiEngr", "es01a", optPic.value
   SaveSetting "Esi2000", "EsiEngr", "ppies01b", optPrc.value
   SaveSetting "Esi2000", "EsiEngr", "ppies01c", optCmt.value
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim sPic As String
   Dim sPrc As String
   Dim sCmt As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "es01", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.value = Val(Left(sOptions, 1))
      optCbd.value = Val(Mid(sOptions, 2, 1))
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Pes01", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   sPic = GetSetting("Esi2000", "EsiEngr", "es01a", sPic)
   If sPic <> "" Then optPic.value = Val(sPic)
   
   sPrc = GetSetting("Esi2000", "EsiEngr", "ppies01b", sPrc)
   sCmt = GetSetting("Esi2000", "EsiEngr", "ppies01c", sCmt)
   If sPrc = "" Then optPrc.value = vbChecked Else optPrc.value = Val(sPrc)
   If sCmt = "" Then optCmt.value = vbChecked Else optCmt.value = Val(sCmt)
   
End Sub

Private Sub lblPicture_DblClick()
   If lblPicture.Caption <> "" Then OpenWebPage lblPicture.Caption
   
End Sub


Private Sub lblTyp_Change()
   If lblTyp = "FULL" Then
      z1(5).Visible = True
      z1(9).Visible = True
      z1(12).Visible = True
      optCbd.Visible = True
      optPrc.Visible = True
      optCmt.Visible = True
   Else
      z1(5).Visible = False
      z1(9).Visible = False
      z1(12).Visible = False
      optCbd.Visible = False
      optPrc.Visible = False
      optCmt.Visible = False
   End If
   
End Sub

Private Sub optCbd_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   If optPic.value = vbChecked Then
      If Trim(lblPicture) = "" Then bPrintPic = 0 _
              Else bPrintPic = 1
   Else
      bPrintPic = 0
   End If
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPic_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   bPrintPic = 0
   PrintReport
   
End Sub
