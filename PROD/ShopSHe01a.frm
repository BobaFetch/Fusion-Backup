VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ShopSHe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Manufacturing Order"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4101
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "ShopSHe01a.frx":0000
      Height          =   315
      Left            =   4800
      Picture         =   "ShopSHe01a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   720
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      Top             =   720
      Width           =   3255
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   240
      TabIndex        =   45
      Top             =   2400
      Width           =   7332
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe01a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6600
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Transaction"
      Top             =   2520
      Width           =   875
   End
   Begin VB.CheckBox optDis 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   120
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Mark MO Released After Creation"
      Top             =   5040
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CheckBox optPrn 
      Alignment       =   1  'Right Justify
      Caption         =   "______"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   5520
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Print The MO After It Is Scheduled And Allocated"
      Top             =   2760
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CheckBox optRel 
      Alignment       =   1  'Right Justify
      Caption         =   "______"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   5520
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Mark MO Released After Creation"
      Top             =   2520
      Width           =   852
   End
   Begin VB.CheckBox optAllocate 
      Caption         =   "Allocate (Sales Only)"
      Enabled         =   0   'False
      Height          =   252
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.ComboBox cmbRev 
      Height          =   315
      Left            =   6360
      Sorted          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Revision-Select From List"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtDoc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Document List Revision. NONE Means That No List Has Been Assigned"
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   2160
      TabIndex        =   10
      Tag             =   "3"
      Top             =   3960
      Width           =   3105
   End
   Begin VB.CheckBox optDmy 
      Caption         =   "Dummy"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox txtStr 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Tag             =   "4"
      Top             =   3240
      Width           =   1250
   End
   Begin VB.ComboBox txtCmp 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Tag             =   "4"
      Top             =   3240
      Width           =   1250
   End
   Begin VB.ComboBox cmbDiv 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   9
      Tag             =   "8"
      ToolTipText     =   "Select Division From List"
      Top             =   3600
      Width           =   860
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6600
      TabIndex        =   27
      ToolTipText     =   "Add The Manufacturing Order And Schedule"
      Top             =   3240
      Width           =   875
   End
   Begin VB.TextBox txtRun 
      Height          =   285
      Left            =   6720
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Next Run Number"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtPri 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Tag             =   "1"
      Text            =   "0"
      ToolTipText     =   "0 to 99 (0 is Highest)"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Run Quantity"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Part Number "
      Top             =   720
      Width           =   3420
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   5040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5205
      FormDesignWidth =   7605
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   300
      Left            =   2160
      TabIndex        =   46
      Top             =   4680
      Visible         =   0   'False
      Width           =   3252
      _ExtentX        =   5741
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Print MO "
      Height          =   252
      Index           =   11
      Left            =   3840
      TabIndex        =   40
      ToolTipText     =   "Print The MO After It Is Scheduled And Allocated"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release MO"
      Height          =   252
      Index           =   10
      Left            =   3840
      TabIndex        =   38
      ToolTipText     =   "Mark MO Released After Creation"
      Top             =   2520
      Width           =   1920
   End
   Begin VB.Label lblSoItemRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SO Item Rev"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5880
      TabIndex        =   37
      Top             =   4800
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label lblSoItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SO Item"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5880
      TabIndex        =   36
      Top             =   4560
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label lblSalesOrder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SO Number"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5880
      TabIndex        =   35
      Top             =   4320
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doc List Rev"
      Height          =   285
      Index           =   52
      Left            =   5280
      TabIndex        =   33
      ToolTipText     =   "Document List Revision. NONE Means That No List Has Been Assigned"
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   31
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   30
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label lblRun 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblTyp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6960
      TabIndex        =   26
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   25
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Production Qty:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division:"
      Enabled         =   0   'False
      Height          =   252
      Index           =   6
      Left            =   3840
      TabIndex        =   23
      Top             =   3600
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Priority:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sched Start Date:"
      Enabled         =   0   'False
      Height          =   252
      Index           =   4
      Left            =   3840
      TabIndex        =   21
      Top             =   3240
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sched Completion Date:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label lblRec 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   18
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   17
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recommended Qty:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Run:"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   1440
      TabIndex        =   14
      Top             =   1335
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "ShopSHe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/5/04 Static Document list logic
'3/15/05 Added GetSchedDate to BackSchedule
'4/1/05 Trap to close if no Company Calendar
'10/12/05 Added column OPMDATE and calculated OPQDATE AND OPMDATE
'11/3/05 Revise Q&M dates to correct Hours/Days difference
'11/3/05 Test Forward Schedule
'11/5/05 Added Saturday and Sunday test
'12/22/05 Added Create MO from Sales Order features (Sales.SaleSLe02b)
'1/27/06 Added GetNextRun
'1/27/06 Added Release Option
'4/27/06 Removed Timer, Added Cancel Button, open/close upper boxes
'4/28/06 Added Time Conversion For BackSchedule Q&M
'7/25/06 Removed Current Part swap.
'1/30/07 Corrected GetDocumentList 7.2.2
'3/20/07 7.3.0 Added UpdateRunsTable (Added Routing Columns)
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter


Dim bGoodWCCal As Boolean
Dim bGoodCoCal As Byte
Dim bGoodRout As Byte
Dim bMoSaved As Byte
Dim bGoodPart As Byte
Dim bMoAdded As Byte
Dim bOldRun As Byte
Dim bOnLoad As Byte

Dim iRunNo As Integer
Dim bView As Integer
Dim cRoutHrs As Currency
Dim cPalevLab As Currency
Dim cPalevExp As Currency
Dim cPalevMat As Currency
Dim cPalevOhd As Currency
Dim cPalevHrs As Currency

Dim sPartNumber As String
Dim sRouting As String
Dim sAssRte As String
Dim sStatus As String

'Passed document stuff
Dim iDocEco As Integer
Dim sDocName As String
Dim sDocClass As String
Dim sDocSheet As String
Dim sDocDesc As String
Dim sDocAdcn As String
Dim sListRef As String
Dim sListRev As String

'3/20/07
'Routings
Dim sRtNumber As String
Dim sRtDesc As String
Dim sRtBy As String
Dim sRtAppBy As String
Dim sRtAppDate As String

Dim sMonths(7) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo Psh01
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   sProcName = "printreport"
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   sCustomReport = GetCustomReport("prdsh01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{RunsTable.RUNREF}='" & sPartNumber & "' " _
          & "AND {RunsTable.RUNNO}=" & Val(txtRun)
'          & " AND {RunsTable.RUNNO} = {@Run} and "
'          & "{PartTable.PARTREF} = {@PartNumber}"
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   MouseCursor 0
   DoEvents
   Exit Sub
   
Psh01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
Psh02:
   DoModuleErrors Me
   
End Sub


Private Sub AllocateItems()
   Dim bByte As Byte
   Dim iRow As Integer
   Dim sPartNumber As String
   On Error Resume Next
   sPartNumber = Compress(cmbPrt)
   sSql = "INSERT INTO RnalTable (RAREF,RARUN,RASO," _
          & "RASOITEM,RASOREV,RAQTY) VALUES('" _
          & sPartNumber & "'," & Val(txtRun) & "," _
          & Val(lblSalesOrder) & "," _
          & Val(lblSoItem) & ",'" _
          & lblSoItemRev & "'," _
          & Val(txtQty) & ")"
   clsADOCon.ExecuteSql sSql
   
End Sub


'10/12/05

Public Function GetWeekEnd(TestDate As Variant) As Byte
   Dim RdoWe As ADODB.Recordset
   GetWeekEnd = 0
   If Left(Format(TestDate, "ddd"), 1) <> "S" Then Exit Function
   
   sSql = "SELECT SUM(COCSHT1+COCSHT2+COCSHT3+COCSHT4) AS AvailHours FROM " _
          & "CoclTable WHERE COCREF='" & Left(TestDate, 3) & "-" _
          & Right(TestDate, 4) & " ' AND COCDAY=" & Val(Mid(TestDate, 5, 2)) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWe, ES_FORWARD)
   If Not IsNull(RdoWe!AvailHours) Then GetWeekEnd = RdoWe!AvailHours
   
   If GetWeekEnd = 0 Then
      If Format(TestDate, "ddd") = "Sun" Then GetWeekEnd = 2 _
                Else GetWeekEnd = 1
   End If
   Set RdoWe = Nothing
   
End Function

Private Sub GetDocumentRevisions()
   cmbRev.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT DLSREV FROM DlstTable WHERE " _
          & "DLSREF='" & Compress(cmbPrt) & "'"
   LoadComboBox cmbRev, -1
   If cmbRev.ListCount > 0 Then cmbRev = cmbRev.List(0)
   cmbRev = txtDoc
   Exit Sub
   
DiaErr1:
   sProcName = "getdocumentrev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function CheckStrings(TestString As String) As String
   Dim iLen As Integer
   Dim K As Integer
   Dim PartNo As String
   Dim NewPart As String
   
   On Error GoTo modErr1
   PartNo = Trim$(TestString)
   iLen = Len(PartNo)
   If iLen > 0 Then
      For K = 1 To iLen
         If Mid$(PartNo, K, 1) = Chr$(34) Or Mid$(PartNo, K, 1) = Chr$(39) _
                 Or Mid$(PartNo, K, 1) = Chr$(44) Then
            Mid$(PartNo, K, 1) = "-"
         End If
      Next
   End If
   CheckStrings = PartNo
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   On Error Resume Next
   CheckStrings = ""
End Function

Private Function GetDocInformation(DocumentRef As String, DocumentRev As String) As String
   Dim RdoDoc As ADODB.Recordset
   sProcName = "getdocinfo"
   sSql = "SELECT DOREF,DONUM,DOREV,DOCLASS,DOSHEET,DODESCR,DOECO," _
          & "DOADCN FROM DdocTable where (DOREF='" & DocumentRef & "' " _
          & "AND DOREV='" & DocumentRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         GetDocInformation = "" & Trim(!DOREF)
         sDocName = "" & Trim(!DONUM)
         sDocClass = "" & Trim(!DOCLASS)
         sDocSheet = "" & Trim(!DOSHEET)
         sDocDesc = "" & Trim(!DODESCR)
         iDocEco = !DOECO
         sDocAdcn = "" & Trim(!DOADCN)
         ClearResultSet RdoDoc
      End With
      sDocName = CheckStrings(sDocName)
      sDocAdcn = CheckStrings(sDocAdcn)
   Else
      sDocName = ""
      sDocClass = ""
      sDocSheet = ""
      sDocDesc = ""
      iDocEco = 0
      sDocAdcn = ""
   End If
   Set RdoDoc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getdocinfo"
   
End Function

Private Sub GetMonths(sStartMonth As String, bBackWard As Byte)
   Dim m As Integer
   m = Val(Right(sStartMonth, 4))
   'pass the month as format(sStartMonth,"mmm-yyyy")
   'returns Jan-1997
   '
   'Get (6) months for now
   Erase sMonths
   On Error GoTo DiaErr1
   If bBackWard Then
      'Backwards
      Select Case Left(sStartMonth, 3)
         Case "Jan"
            sMonths(1) = "Jan-" & Trim(str(m))
            sMonths(2) = "Dec-" & Trim(str(m - 1))
            sMonths(3) = "Nov-" & Trim(str(m - 1))
            sMonths(4) = "Oct-" & Trim(str(m - 1))
            sMonths(5) = "Sep-" & Trim(str(m - 1))
            sMonths(6) = "Aug-" & Trim(str(m - 1))
         Case "Feb"
            sMonths(1) = "Feb-" & Trim(str(m))
            sMonths(2) = "Jan-" & Trim(str(m))
            sMonths(3) = "Dec-" & Trim(str(m - 1))
            sMonths(4) = "Nov-" & Trim(str(m - 1))
            sMonths(5) = "Oct-" & Trim(str(m - 1))
            sMonths(6) = "Sep-" & Trim(str(m - 1))
         Case "Mar"
            sMonths(1) = "Mar-" & Trim(str(m))
            sMonths(2) = "Feb-" & Trim(str(m))
            sMonths(3) = "Jan-" & Trim(str(m))
            sMonths(4) = "Dec-" & Trim(str(m - 1))
            sMonths(5) = "Nov-" & Trim(str(m - 1))
            sMonths(6) = "Oct-" & Trim(str(m - 1))
         Case "Apr"
            sMonths(1) = "Apr-" & Trim(str(m))
            sMonths(2) = "Mar-" & Trim(str(m))
            sMonths(3) = "Feb-" & Trim(str(m))
            sMonths(4) = "Jan-" & Trim(str(m))
            sMonths(5) = "Dec-" & Trim(str(m - 1))
            sMonths(6) = "Nov-" & Trim(str(m - 1))
         Case "May"
            sMonths(1) = "May-" & Trim(str(m))
            sMonths(2) = "Apr-" & Trim(str(m))
            sMonths(3) = "Mar-" & Trim(str(m))
            sMonths(4) = "Feb-" & Trim(str(m))
            sMonths(5) = "Jan-" & Trim(str(m))
            sMonths(6) = "Dec-" & Trim(str(m - 1))
         Case "Jun"
            sMonths(1) = "Jun-" & Trim(str(m))
            sMonths(2) = "May-" & Trim(str(m))
            sMonths(3) = "Apr-" & Trim(str(m))
            sMonths(4) = "Mar-" & Trim(str(m))
            sMonths(5) = "Feb-" & Trim(str(m))
            sMonths(6) = "Jan-" & Trim(str(m))
         Case "Jul"
            sMonths(1) = "Jul-" & Trim(str(m))
            sMonths(2) = "Jun-" & Trim(str(m))
            sMonths(3) = "May-" & Trim(str(m))
            sMonths(4) = "Apr-" & Trim(str(m))
            sMonths(5) = "Mar-" & Trim(str(m))
            sMonths(6) = "Feb-" & Trim(str(m))
         Case "Aug"
            sMonths(1) = "Aug-" & Trim(str(m))
            sMonths(2) = "Jul-" & Trim(str(m))
            sMonths(3) = "Jun-" & Trim(str(m))
            sMonths(4) = "May-" & Trim(str(m))
            sMonths(5) = "Apr-" & Trim(str(m))
            sMonths(6) = "Mar-" & Trim(str(m))
         Case "Sep"
            sMonths(1) = "Sep-" & Trim(str(m))
            sMonths(2) = "Aug-" & Trim(str(m))
            sMonths(3) = "Jul-" & Trim(str(m))
            sMonths(4) = "Jun-" & Trim(str(m))
            sMonths(5) = "May-" & Trim(str(m))
            sMonths(6) = "Apr-" & Trim(str(m))
         Case "Oct"
            sMonths(1) = "Oct-" & Trim(str(m))
            sMonths(2) = "Sep-" & Trim(str(m))
            sMonths(3) = "Aug-" & Trim(str(m))
            sMonths(4) = "Jul-" & Trim(str(m))
            sMonths(5) = "Jun-" & Trim(str(m))
            sMonths(6) = "May-" & Trim(str(m))
         Case "Nov"
            sMonths(1) = "Nov-" & Trim(str(m))
            sMonths(2) = "Oct-" & Trim(str(m))
            sMonths(3) = "Sep-" & Trim(str(m))
            sMonths(4) = "Aug-" & Trim(str(m))
            sMonths(5) = "Jul-" & Trim(str(m))
            sMonths(6) = "Jun-" & Trim(str(m))
         Case "Dec"
            sMonths(1) = "Dec-" & Trim(str(m))
            sMonths(2) = "Nov-" & Trim(str(m))
            sMonths(3) = "Oct-" & Trim(str(m))
            sMonths(4) = "Sep-" & Trim(str(m))
            sMonths(5) = "Aug-" & Trim(str(m))
            sMonths(6) = "Jul-" & Trim(str(m))
      End Select
   Else
      'Forward
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getmonths"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub BackSchedule()

   MouseCursor ccHourglass
   
   Dim mo As New ClassMO
   mo.ScheduleOperations sPartNumber, lblRun, CCur(txtQty), txtCmp, True

   MouseCursor ccDefault
   
End Sub

'3/20/07 7.3.0 Add Actual Routing information

Private Sub UpdateRunsTable()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   
   sSql = "SELECT RUNRTNUM FROM RunsTable WHERE RUNRTNUM='FOOBAR'"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum <> 0 Then
      Err.Clear
      clsADOCon.ADOErrNum = 0
      
      sSql = "ALTER TABLE RunsTable ADD " _
             & "RUNRTNUM CHAR(30) NULL DEFAULT('')," _
             & "RUNRTDESC CHAR(30) NULL DEFAULT('')," _
             & "RUNRTBY CHAR(20) NULL DEFAULT('')," _
             & "RUNRTAPPBY CHAR(20) NULL DEFAULT('')," _
             & "RUNRTAPPDATE CHAR(8) NULL DEFAULT('')"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "UPDATE RunsTable SET RUNRTNUM=''," _
                & "RUNRTDESC='',RUNRTBY=''," _
                & "RUNRTAPPBY='',RUNRTAPPDATE='' " _
                & "WHERE RUNRTNUM IS NULL"
         clsADOCon.ExecuteSql sSql
      End If
   End If
   
End Sub

Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   
End Sub


Private Sub cmbPrt_Click()
   cmdAdd.Enabled = False
   If optAllocate.Value = vbChecked Then cmdAdd.Enabled = True
   bGoodPart = GetPart(True)
   
End Sub


Private Sub cmbPrt_GotFocus()
   cmbPrt_Click
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then Exit Sub
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive.", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodPart = GetPart(False)
   
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
   bView = 0
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bView = 1
End Sub

Private Sub txtPrt_Change()
   cmbPrt = txtPrt
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   'MM  Added
   bMoAdded = 0
End Sub


Private Sub txtPrt_LostFocus()
   If cmdFnd.Value <> 0 Then
      Exit Sub
   End If

   cmbPrt = txtPrt
   txtPrt = CheckLen(txtPrt, 30)
   If bView = 1 Then Exit Sub
   If txtPrt = "" Then Exit Sub
   'MM  Added
   If bMoAdded = 1 Then Exit Sub
   
   If (Not ValidPartNumber(txtPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive.", _
         vbInformation, Caption
      txtPrt = ""
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodPart = GetPart(False)
End Sub



Private Sub cmbRev_Click()
   txtDoc = cmbRev
   
End Sub

Private Sub cmbRev_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If cmbRev.ListCount > 0 Then
      For iList = 0 To cmbRev.ListCount - 1
         If cmbRev = cmbRev.List(iList) Then bByte = 1
      Next
      If bByte = 0 Then cmbRev = txtDoc
   Else
      cmbRev = txtDoc
   End If
   txtDoc = cmbRev
   
End Sub


Private Sub cmbRte_Click()
   If Left(cmbRte, 10) <> "No Routing" Then
      sRouting = Compress(cmbRte)
      GetRouting
   End If
   
End Sub

Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   If Left(cmbRte, 10) <> "No Routing" Then
      sRouting = Compress(cmbRte)
      GetRouting
   End If
   
End Sub


Private Sub cmdAdd_Click()
   Dim sMsg As String
   Dim bResponse As Byte
   'use function errors
   'On Error Resume Next
   bMoAdded = 0
   
   'if required, check if routing is approved
   ' if routing not approved, do not add if that option is selected
   If Not IsRoutingApproved Then Exit Sub
   
   If cmbRte = "" Or Left(cmbRte, 10) = "No Routing" Then
      MsgBox "Requires A Valid Routing.", vbExclamation, Caption
      cmbRte.SetFocus
      Exit Sub
   End If
   
   If Val(txtQty) <= 0 Then
      MsgBox "The Quantity Is Wrong.", vbExclamation, Caption
      txtQty.SetFocus
      Exit Sub
   End If
   If Len(txtCmp) = 0 Then
      MsgBox "The Date Is Not Valid.", vbExclamation, Caption
      txtCmp.SetFocus
      Exit Sub
   End If
   bOldRun = GetRun
   If bOldRun Then Exit Sub
   lblRun = txtRun
   
   sMsg = "Add MO " & Trim(cmbPrt) & " Run " & Trim(txtRun) & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmbPrt.Enabled = False
      txtRun.Enabled = False
      MouseCursor 13
      PrgBar.Visible = True
      cmdAdd.Enabled = False
      bGoodRout = CopyRouting()
      GetDocumentList
      If bMoAdded = 0 Then
         PrgBar.Value = 0
         PrgBar.Visible = False
         Exit Sub
      End If
      If bGoodRout = 1 Then
         BackSchedule
      Else
         optDmy.Value = vbChecked
      End If
      PrgBar.Visible = False
      MouseCursor 0
      MsgBox "MO scheduled"
      
      ' Dummy Option will be checked only if it is called from sales module.
    If optDmy.Value = vbUnchecked Then
      ' Revise the MO
      ShopSHe02a.optFrom = vbChecked
      ShopSHe02a.Show
      
    Else ' Only if this is called from Sales module
    
      sSql = "UPDATE SoitTable SET ITMOCREATED=1 WHERE " _
             & "(ITSO=" & Val(lblSalesOrder) & " AND ITNUMBER=" _
             & Val(lblSoItem) & " AND ITREV='" & lblSoItemRev & "')"
      clsADOCon.ExecuteSql sSql
    
      If optAllocate.Value = vbChecked Then AllocateItems
      If optPrn.Value = vbChecked Then PrintReport
      'Unload Me
    
    End If
    
   Else
      'On Error Resume Next
      cmdAdd.Enabled = False
      cmbPrt.Enabled = True
      txtRun.Enabled = True
      cmbRev.Enabled = True
      cmdQuit.Enabled = False
      cmbPrt.SetFocus
      CancelTrans
   End If
   
   PrgBar.Visible = False
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4101
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdQuit_Click()
   cmdAdd.Enabled = False
   cmbPrt.Enabled = True
   txtRun.Enabled = True
   cmbRev.Enabled = True
   cmdQuit.Enabled = False
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      UpdateRunsTable
      If optAllocate.Value = vbChecked Then cmdAdd.Enabled = True
      bGoodCoCal = GetCompanyCalendar()
      If bGoodCoCal = 0 Then
         MsgBox "There Is No Company Calendar For The Period.", _
            vbInformation, Caption
         If optAllocate.Value = vbUnchecked Then CapaCPe04a.Show
         Unload Me
         Exit Sub
      End If
      bGoodRout = GetDefRoutings()
      If bGoodRout Then
         bGoodRout = 0
'StopwatchStart
         FillRoutings
         FillDivisions
'StopwatchStop "FillRoutings"
         If cmbDiv.ListCount > 0 Then cmbDiv = cmbDiv.List(0)
         
         If optAllocate.Value = vbUnchecked Then
            
            'FillMoParts
            Dim bPartSearch As Boolean

            bPartSearch = GetPartSearchOption
            SetPartSearchOption (bPartSearch)

            If (Not bPartSearch) Then FillMoParts
            
         Else
            cmbPrt.Enabled = True
            txtPrt.Visible = False
            cmdFnd.Visible = False
         End If
         '            If Trim(cUR.CurrentPart) <> "" Then
         '                cmbPrt = cUR.CurrentPart
         '                bGoodPart = GetPart(True)
         '            End If
         bOnLoad = 0
      Else
         Unload Me
      End If
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bGoodSoMo = 0
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAEXTDESC,PARUN,PAUNITS,PARRQ," _
          & "PALEVEL,PAROUTING,PALEVLABOR,PALEVEXP,PALEVMATL,PALEVOH,PALEVHRS," _
          & "PAFLOWTIME,PADOCLISTREF,PADOCLISTREV FROM PartTable WHERE " _
          & "PARTREF= ? AND (PALEVEL<5 OR PALEVEL=8)"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cUR.CurrentPart = Trim(cmbPrt)
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing

   SaveCurrentSelections
   If bMoAdded = 0 Then FormUnload
   Set ShopSHe01a = Nothing
   On Error GoTo 0
   
End Sub

Private Sub lblRun_Click()
   'never visible-tracks run number on load of ShopSHe02a
   
End Sub

Private Sub lblType_Change()
   If Left(cmbRte, 7) = "No Rout" Then
      lblType.ForeColor = ES_RED
   Else
      lblType.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optAllocate_Click()
   'only used in Sales
   
End Sub

Private Sub optDmy_Click()
   'Never visible - Message Switch to revise mo
   If optDmy.Value = vbChecked Then
      On Error Resume Next
      PrgBar.Value = 100
      MouseCursor 0
      'If optAllocate.value = vbChecked Then AllocateItems
      bGoodSoMo = 1
      If bGoodRout Then
         MsgBox "MO Was Successfully Entered And Scheduled.", vbInformation, Caption
      Else
         MsgBox "MO Was Successfully Entered But Not Scheduled.", vbInformation, Caption
      End If
      PrgBar.Visible = False
      If optAllocate.Value = vbUnchecked Then
         ShopSHe02a.cmbPrt = cmbPrt
         ShopSHe02a.cmbRun = txtRun
         ShopSHe02a.optFrom.Value = vbChecked
         ShopSHe02a.Show
         optDmy.Value = vbUnchecked
      Else
'         sSql = "UPDATE SoitTable SET ITMOCREATED=1 WHERE " _
'                & "(ITSO=" & Val(lblSalesOrder) & " AND ITNUMBER=" _
'                & Val(lblSoItem) & " AND ITREV='" & lblSoItemRev & "')"
'         clsADOCon.ExecuteSQL sSql
         ShopSHe01a.txtQty.Enabled = True
         'If optPrn.Value = vbChecked Then PrintReport
         'Unload Me
      End If
   End If
   
End Sub

Private Sub optRel_Click()
   If optRel.Value = vbChecked Then sStatus = "RL" _
                     Else sStatus = "SC"
   
End Sub


Private Sub FillMoParts()
'StopwatchStart
   Dim RdoQry1 As ADODB.Recordset
   Dim sMonth As String
   Dim bType As Byte
   
   bType = GetAllowedTypes()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   
   sSql = "SELECT PARTREF,PARTNUM,PALEVEL,PARUN FROM PartTable WHERE " _
          & "(PALEVEL<" & (bType + 1) & " OR PALEVEL=8) AND PAPRODCODE<>'BID' " _
          & " AND PAINACTIVE = 0 AND PAOBSOLETE = 0 ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQry1, ES_FORWARD)
   If bSqlRows Then
      With RdoQry1
         If cmbPrt = "" Then
            cmbPrt = "" & Trim(!PartNum)
            txtRun = Format(!PARUN + 1, "####0")
            lblTyp = "" & !PALEVEL
         End If
         Do Until .EOF
            AddComboStr cmbPrt.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoQry1
      End With
   End If
   Set RdoQry1 = Nothing
   bGoodWCCal = GetCenterCalendar(Me)
   bGoodPart = GetPart(True)
'StopwatchStop "FillMoParts"
   Exit Sub
   
DiaErr1:
   sProcName = "fillmoparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtCmp_Click()
   If Not bOldRun And Val(txtQty) > 0 Then cmdAdd.Enabled = True Else cmdAdd.Enabled = False
   
End Sub

Private Sub txtCmp_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtCmp_LostFocus()
   txtCmp = CheckDateEx(txtCmp)
   
End Sub


Private Sub txtPri_Click()
   If Not bOldRun And Val(txtQty) > 0 Then cmdAdd.Enabled = True Else cmdAdd.Enabled = False
   
End Sub

Private Sub txtPri_GotFocus()
   txtPri_Click
   
End Sub


Private Sub txtPri_LostFocus()
   txtPri = CheckLen(txtPri, 2)
   txtPri = Format(Abs(Val(txtPri)), "#0")
   
End Sub

Private Sub txtQty_Click()
   If optAllocate.Value = vbChecked Then
      cmdAdd.Enabled = True
   Else
      If Not bOldRun Then cmdAdd.Enabled = True Else cmdAdd.Enabled = True
   End If
   
End Sub

Private Sub txtQty_LostFocus()
   cmdQuit.Enabled = True
   cmbPrt.Enabled = False
   txtRun.Enabled = False
   cmbRev.Enabled = False
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If Val(txtQty) > 0 Then cmdAdd.Enabled = True Else cmdAdd.Enabled = False
   
End Sub

Private Sub txtRun_Click()
   cmdAdd.Enabled = False
   If optAllocate.Value = vbChecked Then cmdAdd.Enabled = True
   
End Sub

Private Sub txtRun_GotFocus()
   iRunNo = Val(txtRun)
   txtRun_Click
   
End Sub


Private Sub txtRun_LostFocus()
   txtRun = CheckLen(txtRun, 5)
   bOldRun = GetRun()
   
End Sub


Private Sub txtStr_Click()
   If Not bOldRun And Val(txtQty) > 0 Then cmdAdd.Enabled = True Else cmdAdd.Enabled = False
   
End Sub

Private Sub txtStr_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtStr_LostFocus()
   txtStr = CheckDateEx(txtStr)
   
End Sub



Public Function GetPart(OnLoad As Byte) As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim iList As Integer
   Dim cManFlow As Currency
   sPartNumber = Compress(cmbPrt)
   sRouting = ""
   If sPartNumber = "" Then
      GetPart = 0
      txtRun = "0"
      txtQty = "1.000"
      lblDsc = ""
      lblExt = ""
      lblRec = "0"
      lblUom(0) = ""
      lblUom(1) = ""
      sPartNumber = ""
      Exit Function
   End If
   On Error GoTo DiaErr1
   GetPart = 0
   AdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry, ES_KEYSET, False, 1)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         txtRun = Format(!PARUN + 1, "####0")
         lblDsc = "" & Trim(!PADESC)
         lblTyp = "" & !PALEVEL
         lblExt = "" & !PAEXTDESC
         lblRec = Format(0 + !PARRQ, "####0")
         lblUom(0) = "" & Trim(!PAUNITS)
         lblUom(1) = "" & Trim(!PAUNITS)
         cManFlow = Format(!PAFLOWTIME, "####0.00")
         If optAllocate.Value = vbUnchecked Then txtCmp = Format(Now + cManFlow, "mm/dd/yyyy")
         
         'Buds 10/14/99
         cPalevLab = !PALEVLABOR
         cPalevExp = !PALEVEXP
         cPalevMat = !PALEVMATL
         cPalevOhd = !PALEVOH
         cPalevHrs = !PALEVHRS
         
         sRouting = "" & Trim(!PAROUTING)
         sAssRte = sRouting
         GetRouting
         '5/5/04
         txtDoc = "" & Trim(!PADOCLISTREV)
         If OnLoad = 0 Then
            sListRef = "" & Trim(!PADOCLISTREF)
            sListRev = "" & Trim(!PADOCLISTREV)
            txtQty.Enabled = True
            If Not IsNull(!PARRQ) Then
               If optAllocate.Value = vbUnchecked Then
                  If !PARRQ = 0 Then
                     txtQty = "1.000"
                  Else
                     txtQty = Format(0 + !PARRQ, "####0.000")
                  End If
               End If
            End If
            txtCmp.Enabled = True
            txtPri.Enabled = True
            cmbDiv.Enabled = True
            cmbRte.Enabled = True
            For iList = 2 To 9
               If iList <> 4 Then z1(iList).Enabled = True
            Next
         Else
            txtQty.Enabled = False
            If optAllocate.Value = vbUnchecked Then txtQty = ""
            txtCmp.Enabled = False
            txtPri.Enabled = False
            cmbDiv.Enabled = False
            cmbRte.Enabled = False
         End If
         ClearResultSet RdoPrt
         GetDocumentRevisions
         GetPart = 1
      End With
   Else
      MsgBox "Part With Wasn't Found (Part Type 1,2,3,8?).", vbExclamation, Caption
      txtRun = "0"
      txtQty = "0"
      lblDsc = ""
      lblExt = ""
      lblRec = "0"
      lblUom(0) = ""
      lblUom(1) = ""
      sPartNumber = ""
      GetPart = 0
      cPalevLab = 0
      cPalevExp = 0
      cPalevMat = 0
      cPalevOhd = 0
      cPalevHrs = 0
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function CopyRouting() As Byte
   Dim RdoRte As ADODB.Recordset
   Dim iCurrentOp As Integer
   Dim sRoutType As String
   Dim sMsg As String
   
   'See if there is a routing of any kind
   If Trim(cmbRte) <> "No Routing Assignment" Then
      sRouting = Compress(cmbRte)
   Else
      sRouting = ""
   End If
   PrgBar.Value = 10
   
   If Val(lblTyp) = 7 Then
      CopyRouting = 0
      bMoAdded = 0
      MouseCursor 0
      MsgBox "Invalid Part Type 7.", vbExclamation, Caption
      Exit Function
   End If
   bMoAdded = 1
   On Error GoTo DiaErr3
   
   'add the Mo
   sSql = "INSERT INTO RunsTable (RUNREF,RUNNO,RUNSCHED," _
          & "RUNSTATUS,RUNDIVISION,RUNQTY,RUNPRIORITY,RUNBUDLAB," _
          & "RUNBUDEXP,RUNBUDMAT,RUNBUDOH,RUNBUDHRS," _
          & "RUNREMAININGQTY,RUNRTNUM,RUNRTDESC,RUNRTBY,RUNRTAPPBY,RUNRTAPPDATE) " _
          & "VALUES('" & sPartNumber & "'," _
          & Val(txtRun) & ",'" _
          & txtCmp & "','" _
          & sStatus & "','" _
          & cmbDiv & "'," _
          & Val(txtQty) & "," _
          & Val(txtPri) & "," _
          & cPalevLab & "," _
          & cPalevExp & "," _
          & cPalevMat & "," _
          & cPalevOhd & "," _
          & cPalevHrs & "," _
          & Val(txtQty) & ",'" _
          & sRtNumber & "','" _
          & sRtDesc & "','" _
          & sRtBy & "','" _
          & sRtAppBy & "','" _
          & sRtAppDate & "')"
          
   Debug.Print sSql
   
   
   
'   Dim strMsg As String
'
'   Dim strFileName As String
'   Dim strFullPath As String
'   Dim nFileNum As Integer
'
'   strFileName = "FusionLog.txt"
'   strFullPath = App.Path & "\" & strFileName
'   nFileNum = FreeFile
'   Open strFullPath For Append As nFileNum
'
'   strMsg = "********START:" & sProgName & ":**********"
'   AddLog nFileNum, strMsg
'
'   AddLog nFileNum, sSql

   clsADOCon.ExecuteSql sSql
     
   If clsADOCon.RowsAffected = 0 Then
      MouseCursor 0
      MsgBox "Could Not Add The MO..", vbExclamation, Caption
      PrgBar.Visible = False
      bMoAdded = 0
      CopyRouting = 0
      Exit Function
   Else
      bMoAdded = 1
   End If
   
   sSql = "UPDATE PartTable SET PARUN=" & Val(txtRun) & " " _
          & "WHERE PARTREF='" & sPartNumber & "'"
   clsADOCon.ExecuteSql sSql
   PrgBar.Value = 20
   
   
   'No routing, use default (if any)
   If Compress(sRouting) = "" Then
      sRoutType = "RTEPART" & Trim(lblTyp)
      sSql = "SELECT " & sRoutType & " FROM ComnTable WHERE COREF=1"
      Set RdoRte = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
      If Not RdoRte.BOF And Not RdoRte.EOF Then sRouting = "" & Trim(RdoRte.Fields(0))
   End If
   On Error GoTo DiaErr1
   sRouting = Compress(sRouting)
   

   If Len(Trim(sRouting)) > 0 Then
      CopyRouting = 1
   Else
      MouseCursor 0
      CopyRouting = 0
      MsgBox "No Routing Or Default For This Part.", vbExclamation, Caption
      Exit Function
   End If
   'There's a routing
   PrgBar.Value = 30

   
   On Error Resume Next
   'Delete possible duplicate keys
   sSql = "DELETE FROM RnopTable WHERE OPREF='" & sPartNumber _
          & "' AND OPRUN=" & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   
   On Error GoTo DiaErr1
   sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPSETUP,OPUNIT," _
          & "OPPICKOP,OPSERVPART,OPQHRS,OPMHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT FROM " _
          & "RtopTable WHERE OPREF='" & sRouting & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_KEYSET)
   
   
   If bSqlRows Then
      With RdoRte
         Do Until .EOF

            
            On Error Resume Next
            If iCurrentOp = 0 Then iCurrentOp = !opNo
            sRoutType = "" & Trim(!OPCOMT)
            sRoutType = ReplaceString(sRoutType)
            sSql = "INSERT INTO RnopTable (OPREF,OPRUN,OPNO,OPSHOP,OPCENTER," _
                   & "OPQHRS,OPMHRS,OPPICKOP,OPSERVPART,OPSUHRS,OPUNITHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT) " _
                   & "VALUES('" & sPartNumber & "'," _
                   & Trim(txtRun) & "," _
                   & !opNo & ",'" _
                   & Trim(!OPSHOP) & "','" _
                   & Trim(!OPCENTER) & "'," _
                   & !OPQHRS & "," _
                   & !OPMHRS & "," _
                   & !OPPICKOP & ",'" _
                   & Trim(!OPSERVPART) & "'," _
                   & !OPSETUP & "," _
                   & !OPUNIT & "," _
                   & !OPSVCUNIT & ",'" _
                   & Trim(!OPTOOLLIST) & "','" _
                   & Trim(sRoutType) & "')"
            clsADOCon.ExecuteSql sSql
   
   'AddLog nFileNum, sSql
            
            .MoveNext
         Loop
         ClearResultSet RdoRte
      End With
      sSql = "UPDATE RunsTable SET RUNOPCUR=" & iCurrentOp & " " _
             & "WHERE RUNREF='" & sPartNumber & "' AND RUNNO=" _
             & Val(txtRun) & " "
      clsADOCon.ExecuteSql sSql
      CopyRouting = 1
   Else
      MouseCursor 0
      sMsg = "Wasn't Able To Copy The Routing " & vbCrLf _
             & "Or There Are No Operations To Copy. " & vbCrLf _
             & "MO Was Successfully Added."
      MsgBox sMsg, vbExclamation, Caption
      CopyRouting = 0
      bMoAdded = 1
      Exit Function
   End If
   CopyRouting = 1
   Set RdoRte = Nothing
   PrgBar.Value = 40
   
'   strMsg = "**************END******************"
'   AddLog nFileNum, strMsg
'
'   ' Close the file
'   Close nFileNum
   
   Exit Function
   
DiaErr1:
MsgBox "DiaErr1"
   
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
MsgBox "DiaErr2"
   PrgBar.Visible = False
   MouseCursor 0
   clsADOCon.RollbackTrans
   Set RdoRte = Nothing
   CopyRouting = 0
   sMsg = str(CurrError.Number) & vbCrLf & CurrError.Description & vbCrLf _
          & "Couldn't Copy Routing."
   On Error GoTo 0
   
   Exit Function
   
DiaErr3:
MsgBox "DiaErr3"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr4
DiaErr4:
   On Error Resume Next
   PrgBar.Visible = False
   MouseCursor 0
   CopyRouting = 0
   sMsg = str(CurrError.Number) & vbCrLf & CurrError.Description & vbCrLf _
          & "Couldn't Add MO.."
   clsADOCon.ExecuteSql "DELETE FROM RnopTable WHERE OPREF='" & sPartNumber & "' AND OPRUN=" & txtRun & " "

MsgBox "DiaErr4" & sMsg
   
   Set RdoRte = Nothing
   On Error GoTo 0
   
End Function

Private Function GetRun() As Byte
   Dim RdoRun As ADODB.Recordset
   
   On Error GoTo DiaErr1
   If Val(txtRun) = 0 Then
      MsgBox "Next Run Is Either Blank Or Zero.", vbExclamation, Caption
      GetRun = 1
      txtRun = Format(iRunNo, "####0")
   Else
      sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
             & "RUNREF='" & sPartNumber & "' AND RUNNO=" & Val(Trim(txtRun)) & " "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun)
      If bSqlRows Then
         MsgBox "That Run Has Already Been Recorded.", vbExclamation, Caption
         GetRun = 1
         ClearResultSet RdoRun
         bGoodPart = GetPart(False)
      Else
         GetRun = 0
      End If
   End If
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FinishMo()
   On Error Resume Next
   PrgBar.Value = 100
   MouseCursor 0
   optDmy.Value = vbChecked
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPri = "99"
   txtDoc.BackColor = Me.BackColor
   sStatus = "SC"
   
End Sub


Private Sub GetRouting()
   Dim RdoRte As ADODB.Recordset
   Dim sRoutType As String
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM RthdTable WHERE RTREF='" & Compress(sRouting) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         cmbRte = "" & Trim(!RTNUM)
         If Compress(sAssRte) = Compress(sRouting) Then
            lblType = "Assigned Routing."
         Else
            lblType = "" & Trim(!RTDESC)
         End If
         sRtNumber = "" & Trim(!RTNUM)
         sRtDesc = "" & Trim(!RTDESC)
         sRtBy = "" & Trim(!RTBY)
         sRtAppBy = "" & Trim(!RTAPPBY)
         If Not IsNull(!RTAPPDATE) Then
            sRtAppDate = Format$(!RTAPPDATE, "mm/dd/yyyy")
         Else
            sRtAppDate = ""
         End If
         ClearResultSet RdoRte
      End With
   Else
      sRoutType = "RTEPART" & Trim(lblTyp)
      sSql = "SELECT " & sRoutType & " FROM ComnTable WHERE COREF=1"
      Set RdoRte = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
      If Not RdoRte.BOF And Not RdoRte.EOF Then
         cmbRte = "" & Trim(RdoRte.Fields(0))
         lblType = "Default Routing."
      Else
         cmbRte = ""
         lblType = ""
      End If
   End If
   If Trim(cmbRte) = "" Or Left(cmbRte, 5) = "No Ro" Then
      cmbRte = "No Routing Assignment"
      lblType = "*** Requires A Routing ***"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getrouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetDefRoutings() As Byte
   Dim b As Byte
   Dim iList As Integer
   
   Dim RdoDef As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RTEPART1,RTEPART2,RTEPART3,RTEPART4,RTEPART5," _
          & "RTEPART6,RTEPART8 FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDef, ES_FORWARD)
   If bSqlRows Then
      With RdoDef
         For iList = 0 To 6
            If Trim(.Fields(iList)) = "" Then
               b = 1
            End If
         Next
      End With
   End If
   If b = 1 Then
      MsgBox "Please Enter A Default Routing For Each Part" & vbCrLf _
         & "Type Before Entering A Manufacturing Order..", _
         vbExclamation, Caption
      GetDefRoutings = 0
   Else
      GetDefRoutings = 1
   End If
   Set RdoDef = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getdefrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'10/16/03 to trap Types allowed (type 3 or 4) see Sys Settings

Private Function GetAllowedTypes() As Byte
   Dim RdoTyp As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT COALLOWTYPEFOURMO FROM ComnTable " _
          & "WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTyp, ES_FORWARD)
   If bSqlRows Then
      With RdoTyp
         If Not IsNull(.Fields(0)) Then _
                       GetAllowedTypes = (.Fields(0) + 3) Else _
                       GetAllowedTypes = 3
         ClearResultSet RdoTyp
      End With
   End If
   If Err > 0 Or GetAllowedTypes = 0 Then GetAllowedTypes = 3
   If GetAllowedTypes = 4 Then
      cmbPrt.ToolTipText = "Contains Part Types 1 Through 4"
   Else
      cmbPrt.ToolTipText = "Contains Part Types 1 Through 3"
   End If
   
   Set RdoTyp = Nothing
End Function

Private Sub GetDocumentList()
   Dim RdoList As ADODB.Recordset
   
   Dim iRow As Integer
   Dim sDocRef As String
   Dim sRev As String
   
   On Error GoTo DiaErr1
   sSql = "DELETE FROM RndlTable WHERE RUNDLSRUNREF='" & sListRef & " ' AND " _
          & "RUNDLSRUNNO=" & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   
   sListRef = Compress(cmbPrt)
   If Trim(txtDoc) = "NONE" Then
      MouseCursor 13
      sSql = "SELECT MAX(DLSREV) FROM DlstTable WHERE DLSREF='" & sListRef & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoList, ES_FORWARD)
      If bSqlRows Then
         With RdoList
            If Not IsNull(.Fields(0)) Then
               sListRev = "" & Trim(.Fields(0))
            Else
               On Error Resume Next
               'Dummy Row for joins
               sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF, RUNDLSRUNNO) " _
                      & "VALUES(1,'" & sListRef & "'," & Val(txtRun) & ")"
               clsADOCon.ExecuteSql sSql
               Exit Sub
            End If
            ClearResultSet RdoList
         End With
      End If
   End If
   On Error GoTo DiaErr1
   sSql = "DELETE FROM RndlTable WHERE RUNDLSRUNREF='" & sListRef & " ' AND " _
          & "RUNDLSRUNNO=" & Val(txtRun) & " "
   clsADOCon.ExecuteSql sSql
   
   ' In partTable the Rev is NONE, but the DocList table has a empty string
   ' 3/7/2010
   If (Trim(sListRev) = "NONE") Then
     sListRev = ""
   End If
   
   sSql = "SELECT * FROM DlstTable WHERE DLSREF='" & sListRef & "' " _
          & "AND DLSREV='" & sListRev & "' ORDER BY DLSDOCCLASS,DLSDOCREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoList, ES_FORWARD)
   If bSqlRows Then
      With RdoList
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         
         Do Until .EOF
            iRow = iRow + 1
            sDocRef = GetDocInformation("" & Trim(!DLSDOCREF), "" & Trim(!DLSDOCREV))
            sProcName = "updatemolist"
            sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF," _
                   & "RUNDLSRUNNO,RUNDLSREV,RUNDLSDOCREF,RUNDLSDOCREV," _
                   & "RUNDLSDOCREFLONG,RUNDLSDOCREFDESC,RUNDLSDOCREFSHEET," _
                   & "RUNDLSDOCREFCLASS,RUNDLSDOCREFADCN," _
                   & "RUNDLSDOCREFECO) VALUES(" & iRow & ",'" & Compress(cmbPrt) & "'," _
                   & Val(txtRun) & ",'" & sListRev & "','" & Trim(!DLSDOCREF) & "','" _
                   & Trim(!DLSDOCREV) & "','" & sDocName & "','" & sDocDesc & "','" _
                   & sDocSheet & "','" & sDocClass & "','" & sDocAdcn & "'," _
                   & iDocEco & ")"
            clsADOCon.ExecuteSql sSql
            .MoveNext
         Loop
         ClearResultSet RdoList
      End With
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         SysMsg "Manufacturing Order Updated.", True
      Else
         MsgBox "Could Not Successfully Update The MO.", _
            vbExclamation, Caption
      End If
   Else
      'Dummy Row for joins - Corrected 1/30/07
      On Error Resume Next
      sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF, RUNDLSRUNNO) " _
             & "VALUES(1,'" & Compress(cmbPrt) & "'," & Val(txtRun) & ")"
      clsADOCon.ExecuteSql sSql
   End If
   Set RdoList = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdocumentli"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetNextRun()
   Dim rdoTmp As ADODB.Recordset
   On Error Resume Next
   sSql = "select max(RUNNO) As LastRun FROM RunsTable WHERE RUNSPLITFROMRUNNO=0 " _
          & "AND RUNREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTmp)
   If bSqlRows Then
      With rdoTmp
         txtRun = Format(!LastRun + 1, "####0")
         .Cancel
      End With
      ClearResultSet rdoTmp
   End If
   rdoTmp.Close
   Set rdoTmp = Nothing
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

Private Function IsRoutingApproved() As Boolean

   'is an approved routine required?
   IsRoutingApproved = True
   Dim Routing As String
   Routing = Compress(cmbRte)
   Dim rs As ADODB.Recordset
   sSql = "select CoRequireApprovedRoutings from ComnTable where COREF = 1 and CoRequireApprovedRoutings = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_KEYSET)
   Set rs = Nothing
   If Not bSqlRows Then Exit Function  'approval not required
   
   ' approved routing is required.  Don't allow blank, which triggers a default routing
   If Routing = "" Then
      IsRoutingApproved = False
      MsgBox "An approved routing is required"
      Exit Function
   End If
   
   ' don't allow to proceed if routing does not have an approval date
   sSql = "select RTAPPDATE from RthdTable where RTREF = '" & Routing & "' and RTAPPDATE is not null"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_KEYSET)
   Set rs = Nothing
   If bSqlRows Then Exit Function  'routing is approved
   
   MsgBox "The selected routing is not approved"
   IsRoutingApproved = False
      
End Function

