VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaCLp12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory Movement from WIP - Restock"
   ClientHeight    =   2865
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   2280
      TabIndex        =   45
      Tag             =   "5"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   11
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   10
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   9
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   8
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   7
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox chkLotNum 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   8520
      TabIndex        =   34
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox chkActComt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   3360
      Width           =   200
   End
   Begin VB.ComboBox txtActvDate 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   6555
      Picture         =   "diaCLp12a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Print The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   6000
      Picture         =   "diaCLp12a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Display The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox chkJGL 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2760
      TabIndex        =   16
      Top             =   3405
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkSummary 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2760
      TabIndex        =   15
      Top             =   3105
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox chkDsc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   2160
      Width           =   200
   End
   Begin VB.CheckBox chkExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2760
      TabIndex        =   13
      Top             =   2445
      Width           =   200
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "8"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   20
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
      PictureUp       =   "diaCLp12a.frx":0308
      PictureDn       =   "diaCLp12a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2865
      FormDesignWidth =   7680
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   44
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   43
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   42
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   41
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   40
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   39
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   38
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Part Types:"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   37
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Number"
      Height          =   285
      Index           =   13
      Left            =   480
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   8520
      TabIndex        =   35
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Account Type"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   33
      Top             =   3240
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Comment"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   32
      Top             =   3120
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Activity Date"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   31
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Journal to G.L"
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   30
      Top             =   3405
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Only"
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   29
      Top             =   3105
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   28
      Top             =   1845
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   27
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   26
      Top             =   2160
      Width           =   1785
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class"
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   24
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL) "
      Height          =   285
      Index           =   10
      Left            =   3840
      TabIndex        =   23
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through :"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   22
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   21
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "diaCLp12a"
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
' diaCLp12a - Inventory Movement To Overage/Shortage
'
' Notes:
'
' Created: 09/06/08
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim vAccounts(10) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



'*********************************************************************************

Private Sub cmbCls_LostFocus()
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmbAct_LostFocus()
   Dim b As Boolean
   Dim i As Integer
   On Error Resume Next
   cmbAct = CheckLen(cmbAct, 12)
   For i = 0 To cmbAct.ListCount - 1
      If cmbAct = cmbAct.List(i) Then b = True
   Next
   If Not b Then
      cmbAct = cmbAct.List(0)
      cmbAct.ListIndex = 0
   End If
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductClasses Me
      Me.cmbCls = "ALL"
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
   txtActvDate = Format(ES_SYSDATE, "mm/dd/yy")
   
   ' Fill Account Type
   FillAccType
   GetOptions
   
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaCLp12a = Nothing
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
   On Error GoTo whoops
   'get custom report name if one has been defined
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("fincl12.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   'pass formulas
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "PartClass"
   aFormulaName.Add "ShowPartDesc"
   aFormulaName.Add "ShowExtDesc"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'From " & CStr(txtBeg & " Through " & txtEnd & " for Classes " & cmbCls) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbCls) & "'")
   aFormulaValue.Add chkDsc
   aFormulaValue.Add chkExt
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'pass Crystal SQL if required
   sSql = "{InvaTable.INADATE} >= #" & txtBeg & "# AND {InvaTable.INADATE} <= #" & txtEnd & " 23:55# AND ({InvaTable.INTYPE} = 20 OR {InvaTable.INTYPE} = 21) "
   If cmbCls <> "ALL" Then
      sSql = sSql & " AND {PartTable.PACLASS}='" & cmbCls & "'"
   End If
   
   cCRViewer.SetReportSelectionFormula sSql
   
   ' set the sub sql variable pass the sub report name
   cCRViewer.SetSubRptSelFormula "sr_ClGL_Acc_Sum", sSql
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
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
   Dim sPartType As String
   Dim i As Integer
   
   On Error GoTo whoops
   
   'setmdireportsizemdisect
   
   'get custom report name if one has been defined
   sCustomReport = GetCustomReport("fincl12.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
    For i = 0 To 7
      If ChkTyp(i).Value = vbChecked Then
         sPartType = sPartType & Trim(zTyp(i)) & ","
      End If
   Next
   'pass formulas
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='From " & txtBeg & " Through " & txtEnd & " for Classes " & cmbCls & "'"
   MdiSect.crw.Formulas(3) = "PartClass='" & cmbCls & "'"
   MdiSect.crw.Formulas(4) = "ActivityDate='" & txtActvDate & "'"
   MdiSect.crw.Formulas(5) = "ShowActComment='" & chkActComt & "'"
   MdiSect.crw.Formulas(6) = "AccountType='" & cmbAct & "'"
   MdiSect.crw.Formulas(7) = "PartTypes='" & sPartType & "'"
   MdiSect.crw.Formulas(8) = "ShowPartDesc=" & chkDsc
   MdiSect.crw.Formulas(9) = "ShowExtDesc=" & chkExt
   MdiSect.crw.Formulas(10) = "ShowLotNum=" & chkLotNum
   MdiSect.crw.Formulas(11) = "ShowSummary=" & chkSummary
   MdiSect.crw.Formulas(12) = "ShowGLTransferJournal=" & chkJGL
   
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
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   
   If Len(Trim(sOptions)) > 0 Then
     If dToday < 21 Then
      txtBeg = Mid(sOptions, 1, 8)
      txtEnd = Mid(sOptions, 9, 8)
     Else
      txtBeg = Format(Now, "mm/01/yy")
      txtEnd = GetMonthEnd(txtBeg)
     End If
   End If

   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
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
Private Sub txtActvDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtActvDate_LostFocus()
   txtActvDate = CheckDate(txtActvDate)
End Sub

Private Sub FillAccType()
   Dim i As Integer
   Dim RdoGlm As ADODB.Recordset
   Dim sAccount As String
   
   On Error GoTo whoops
   sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
   If bSqlRows Then
      vAccounts(0) = "ALL"
      With RdoGlm
         i = 1
         vAccounts(i) = "" & Trim(!COASSTACCT)
         
         i = 2
         vAccounts(i) = "" & Trim(!COLIABACCT)
         i = 3
         vAccounts(i) = "" & Trim(!COEQTYACCT)
         
         i = 4
         vAccounts(i) = "" & Trim(!COINCMACCT)
         
         i = 6
         vAccounts(i) = "" & Trim(!COEXPNACCT)
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = 5
            vAccounts(i) = "" & Trim(!COCOGSACCT)
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = 7
            vAccounts(i) = "" & Trim(!COOINCACCT)
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = 8
            vAccounts(i) = "" & Trim(!COOEXPACCT)
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = 9
            vAccounts(i) = "" & Trim(!COFDTXACCT)
         End If
         .Cancel
      End With
   End If
   iTotal = i
   For i = 0 To iTotal
      AddComboStr cmbAct.hWnd, Format$(vAccounts(i))
   Next
   If cmbAct.ListCount > 0 Then
      cmbAct = cmbAct.List(0)
   End If
   Set RdoGlm = Nothing
   Exit Sub
   
whoops:
   sProcName = "FillAccType"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

