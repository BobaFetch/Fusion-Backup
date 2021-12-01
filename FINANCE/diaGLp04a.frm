VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Old Detailed General Ledger (Report)"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox OptJrDetail 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      ToolTipText     =   "Show GL posted detail"
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox OptGlDetail 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   24
      ToolTipText     =   "Show GL posted detail"
      Top             =   3360
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5400
      TabIndex        =   22
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaGLp04a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaGLp04a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4275
      FormDesignWidth =   6720
   End
   Begin VB.ComboBox cmbEndAct 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Tag             =   "3"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox cmbStartAct 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox optPag 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox txtend 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox txtstart 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "4"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox optIna 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
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
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLp04a.frx":0308
      PictureDn       =   "diaGLp04a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   16
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
      PictureUp       =   "diaGLp04a.frx":0594
      PictureDn       =   "diaGLp04a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Journal Detail"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   25
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show GL Detail"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   23
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Inactive Accounts"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   21
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page Break By Account"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblEndAct 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblStartAct 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   18
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Account"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Account"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Beginning"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "diaGLp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

' See the UpdateTables prodecure for database revisions

'*************************************************************************************
'
' diaGLp04a - Trial Balance (Report) / Detail GL (Report)
'
' Notes: Same form used for both reports.
'
' Created: 03/20/01 (nth)
' Revisions:
' 09/17/03 (nth) Added beginning balance to Detail GL per WCK.
' 09/17/03 (nth) Revised and updated trial balance.
' 01/19/05 (nth) Corrected beginning and ending account filter.
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

' bForm documentation
' 0 = Detail GL
' 1 = Trial Balance
Dim bForm As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub CreateActTable()
   Dim NewTb1 As TableDef
   Dim NewTb2 As TableDef
   Dim NewIdx1 As Index
   Dim NewIdx2 As Index
   
   On Error Resume Next
   JetDb.Execute "DROP TABLE AccountActivity"
   JetDb.Execute "DROP TABLE AccountBalance"
   
   Set NewTb1 = JetDb.CreateTableDef("AccountActivity")
   Set NewTb2 = JetDb.CreateTableDef("AccountBalance")
   
   With NewTb1
      .Fields.Append .CreateField("ActRef", dbText, 12)
      .Fields.Append .CreateField("ActDeb", dbDouble)
      .Fields.Append .CreateField("ActCrd", dbDouble)
      .Fields.Append .CreateField("ActJETran", dbInteger)
      .Fields.Append .CreateField("ActJERef", dbInteger)
      .Fields.Append .CreateField("ActJE", dbText, 12)
      .Fields.Append .CreateField("ActJEDesc", dbText, 30)
      .Fields.Append .CreateField("ActJEPost", dbDate)
   End With
   
   JetDb.TableDefs.Append NewTb1
   With NewTb2
      .Fields.Append .CreateField("ActNum", dbText, 12)
      .Fields.Append .CreateField("ActRef", dbText, 12)
      .Fields.Append .CreateField("ActDesc", dbText, 40)
      .Fields.Append .CreateField("ActBal", dbDouble)
   End With
   
   JetDb.TableDefs.Append NewTb2
   
   Set NewTb1 = Nothing
   Set NewTb2 = Nothing
   
   'add the table and indexes to Jet.
   
   On Error Resume Next
   Set NewTb1 = JetDb!AccountActivity
   With NewTb1
      Set NewIdx1 = .CreateIndex
      With NewIdx1
         .Name = "ixaActNum"
         .Fields.Append .CreateField("ActRef")
      End With
      .Indexes.Append NewIdx1
   End With
   
   Set NewTb2 = JetDb!AccountBalance
   With NewTb2
      Set NewIdx2 = .CreateIndex
      With NewIdx2
         .Name = "ixbActNum"
         .Unique = True
         .Fields.Append .CreateField("ActRef")
      End With
      .Indexes.Append NewIdx2
   End With
   
   Set NewTb1 = Nothing
   Set NewTb2 = Nothing
   Set NewIdx2 = Nothing
   Set NewIdx1 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "CreateActTable"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub BuildAccountTotals()
   Dim DbActivity As Recordset
   Dim DbBal As Recordset
   Dim RdoBal As ADODB.Recordset
   Dim rdoAct As ADODB.Recordset
   Dim RdoSum As ADODB.Recordset
   Dim sBegDate As String
   Dim sEnddate As String
   Dim sBegAct As String
   Dim sEndAct As String
   Dim sTemp As String
   Dim iCount As Integer
   
   'Check for valid date entries
   If txtstart = "" Then
      MsgBox "Please Enter A Valid Starting Date.", vbInformation
      txtstart.SetFocus
      Exit Sub
   ElseIf txtEnd = "" Then
      MsgBox "Please Enter A Valid Ending Date.", vbInformation
      txtEnd.SetFocus
      Exit Sub
   End If
   
   MouseCursor 13
   
   sBegDate = txtstart
   sEnddate = txtEnd
   
   sBegAct = Compress(cmbStartAct)
   sEndAct = Compress(cmbEndAct)
   
   'If sBegAct = "" Then sBegAct = Trim(cmbStartAct.List(0))
   'If sEndAct = "" Then sEndAct = Trim(cmbEndAct.List(cmbEndAct.ListCount - 1))
   
   On Error Resume Next
   'ReopenJet
   CreateActTable
   
   JetDb.Execute "DELETE * FROM AccountActivity"
   JetDb.Execute "DELETE * FROM AccountBalance"
   
   Set DbActivity = JetDb.OpenRecordset("AccountActivity", dbOpenDynaset)
   Set DbBal = JetDb.OpenRecordset("AccountBalance", dbOpenDynaset)
   
   sSql = "SELECT SUM(JIDEB) AS Debit, SUM(JICRD) " _
          & "AS Credit,GJPOST,JIACCOUNT,GJNAME,JIDESC,JITRAN,JIREF " _
          & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME=JINAME " _
          & "WHERE (GJPOSTED=1) " _
          & "GROUP BY GJNAME,JIACCOUNT,GJPOST,JIDESC,JITRAN,JIREF " _
          & "HAVING (GJPOST >= '" & sBegDate & "' AND GJPOST <= '" & sEnddate & "')"
   'If sBegAct <> "" Or sEndAct <> "" Then
   '    sSql = sSql & " and isnumeric(jiaccount)=1 "
   'End If
   If sBegAct <> "" Then
      sSql = sSql & "AND JIACCOUNT  >= '" & sBegAct & "'"
   End If
   If sEndAct <> "" Then
      sSql = sSql & " AND JIACCOUNT <= '" & sEndAct & "'"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   ' Dump results to temp jet database
   With rdoAct
      Do While Not .EOF
         DbActivity.AddNew
         DbActivity!ActRef = !JIACCOUNT
         DbActivity!ActDeb = !debit
         DbActivity!ActCrd = !credit
         DbActivity!ActJEPost = !GJPOST
         DbActivity!ActJE = !GJNAME
         DbActivity!ActJEDesc = !JIDESC
         DbActivity!ActJETran = !JITRAN
         DbActivity!ActJERef = !JIREF
         DbActivity.Update
         .MoveNext
      Loop
   End With
   DbActivity.Close
   Set rdoAct = Nothing
   
   sSql = "SELECT DISTINCT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable "
   If Not optIna Then
      sTemp = sTemp & "(GLINACTIVE=0)"
   End If
   If sBegAct <> "" Then
      sTemp = sTemp & " AND(GLACCTREF>='" & sBegAct & "')"
   End If
   If sEndAct <> "" Then
      sTemp = sTemp & " AND(GLACCTREF<='" & sEndAct & "')"
   End If
   If Len(sTemp) Then
      sSql = sSql & " WHERE " & sTemp
   End If
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBal, ES_FORWARD)
   With RdoBal
      Do While Not .EOF
         iCount = iCount + 1
         
         DbBal.AddNew
         DbBal!ActNum = !GLACCTNO
         DbBal!ActRef = !GLACCTREF
         DbBal!ActDesc = !GLDESCR
         
         sSql = "SELECT SUM(GjitTable.JIDEB) AS Debit, " _
                & "SUM(GjitTable.JICRD) AS Credit " _
                & "FROM GjhdTable INNER JOIN " _
                & "GjitTable ON GJNAME = JINAME " _
                & "WHERE (JIACCOUNT = '" & !GLACCTREF _
                & "') AND (GJPOSTED = 1) AND (GJPOST < '" _
                & sBegDate & "')"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoSum, ES_FORWARD)
         
         DbBal!ActBal = (RdoSum!debit - RdoSum!credit)
         Set RdoSum = Nothing
         DbBal.Update
         .MoveNext
      Loop
   End With
   DbBal.Close
   Set RdoBal = Nothing
   
   On Error Resume Next
   JetDb.Close
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "BuildAccountTotals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim sSubSql As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   Dim sTemp As String
   Dim sBegAct As String
   Dim sEndAct As String
   
   On Error GoTo DiaErr1
   
   If InStr(Caption, "Trial Balance") > 0 Then
      sTemp = "fingl05.rpt"
   Else
      sTemp = "fingl04.rpt"
'      MdiSect.crw.Formulas(6) = "PrePeriod='" _
 '                          & Format(DateAdd("d", -1, CDate(txtstart)), "m/d/yy") & "'"
   End If
   
   'get custom report name if one has been defined
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport(sTemp)
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   
   'pass formulas
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Start"
   aFormulaName.Add "End"
   aFormulaName.Add "Title1"
   aFormulaName.Add "ShowGLDetail"
   aFormulaName.Add "ShowJournalDetail"
   aFormulaName.Add "InActive"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbStartAct) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbEndAct) & "'")
   aFormulaValue.Add CStr("'Period Beginning " & Format(txtstart, "m/d/yy") & " And Ending " & Format(txtEnd, "m/d/yy") & "'")
   aFormulaValue.Add OptGlDetail
   aFormulaValue.Add OptJrDetail
   aFormulaValue.Add optIna
   
   
   sTemp = "'Include Inactive Accounts? "
   If optIna Then
      sTemp = sTemp & "Y'"
   Else
      sTemp = sTemp & "N'"
   End If
   aFormulaName.Add "Title2"
   aFormulaValue.Add CStr(sTemp)
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sBegAct = Compress(cmbStartAct)
   sEndAct = Compress(cmbEndAct)

   If (sBegAct = "") Then
      sBegAct = "ALL"
   End If
   If (sEndAct = "") Then
      sEndAct = "ALL"
   End If
   
   
   aRptPara.Add CStr(txtstart)
   aRptPara.Add CStr(txtEnd)
   aRptPara.Add CStr(sBegAct)
   aRptPara.Add CStr(sEndAct)
   aRptPara.Add CStr(optIna)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   ' Set report parameter
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType 'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub




Private Sub FillAccounts()
   ' Fill account combo
   ' Need to add account descriptions? maybe
   Dim rdoAct As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         Do Until .EOF
            AddComboStr cmbStartAct.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr cmbEndAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
      End With
   End If
   Set rdoAct = Nothing
   
   lblEndAct = UpdateActDesc(cmbEndAct)
   lblStartAct = UpdateActDesc(cmbStartAct)
   
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "FillAcounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbEndAct_Click()
   lblEndAct = UpdateActDesc(cmbEndAct)
End Sub

Private Sub cmbEndAct_LostFocus()
   lblEndAct = UpdateActDesc(cmbEndAct)
End Sub

Private Sub cmbStartAct_Click()
   lblStartAct = UpdateActDesc(cmbStartAct)
End Sub

Private Sub cmbStartAct_LostFocus()
   lblStartAct = UpdateActDesc(cmbStartAct)
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      If Me.Caption = "Trial Balance (Report)" Then
         optPag.Visible = False
         z1(6).Visible = False
      Else
         optPag.Visible = True
         z1(6).Visible = True
      End If
      FillAccounts
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   txtstart = Format(Now, "mm/01/yy")
   txtEnd = GetMonthEnd(txtstart)
   GetOptions
   ReopenJet
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
   On Error Resume Next
   'JetDb.Execute "DROP TABLE AccountActivity"
   'JetDb.Execute "DROP TABLE AccountBalance"
   Set diaGLp04a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optDis_Click()
   'BuildAccountTotals
   PrintReport
End Sub

Private Sub optPrn_Click()
   'BuildAccountTotals
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtstart_LostFocus()
   txtstart = CheckDate(txtstart)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = Trim(Trim(str(optPag.Value)) & Trim((str(optIna.Value))) & Trim((str(OptGlDetail.Value))) & Trim((str(OptJrDetail.Value))))
   sOptions = Trim(txtstart.Text) & Trim(txtEnd.Text) & Trim(str(optPag.Value)) & Trim((str(optIna.Value))) & Trim((str(OptGlDetail.Value))) & Trim(str(OptJrDetail.Value))
   SaveSetting "Esi2000", "EsiFina", Me.Name & bForm, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & bForm & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   
   On Error Resume Next
   
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name & bForm, sOptions)

   
   If Len(Trim(sOptions)) > 0 Then
     
     If dToday < 21 Then
      txtstart = Mid(sOptions, 1, 8)
      txtEnd = Mid(sOptions, 9, 8)
     Else
      txtstart = Format(Now, "mm/01/yy")
      txtEnd = GetMonthEnd(txtstart)
     End If
     
      optPag = Val(Mid(sOptions, 17, 1))
      optIna = Val(Mid(sOptions, 18, 1))
      OptGlDetail = Val(Mid(sOptions, 19, 1))
      OptJrDetail = Val(Mid(sOptions, 20, 1))
      
   Else
      optPag = vbUnchecked
      optIna = vbUnchecked
      OptGlDetail = vbUnchecked
      OptJrDetail = vbUnchecked
'      txtstart = Format(Now, "mm/01/yy")
'      txtend = GetMonthEnd(txtstart)
           
   End If
   
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & bForm _
                & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

