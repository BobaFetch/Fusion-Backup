VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CustSh01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MO Priority/Work Center Schedules"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CustSh01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optData 
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optReport 
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton optPrn 
      DownPicture     =   "CustSh01.frx":07AE
      Height          =   320
      Left            =   9000
      Picture         =   "CustSh01.frx":0938
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print This Report"
      Top             =   1440
      Width           =   350
   End
   Begin VB.CheckBox optFrom 
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSel 
      Cancel          =   -1  'True
      Caption         =   "S&elect"
      Height          =   315
      Left            =   8520
      TabIndex        =   3
      ToolTipText     =   "Fill The Grid"
      Top             =   960
      Width           =   875
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Double Click Or Press Enter To Edit Entry"
      Top             =   1920
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      Enabled         =   0   'False
      AllowUserResizing=   1
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Work Center Or leave Blank For All"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   8520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7800
      Top             =   360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5820
      FormDesignWidth =   9525
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2640
      TabIndex        =   18
      Top             =   5400
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   19
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2760
      TabIndex        =   16
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   8160
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Sched Complete from"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "AWI Special"
      Height          =   252
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "CustSh01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/8/05 Removed Jet and created report tables
'10/20/05 Re-added the operation and changed the resizing of the form
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
'Dim DbRpt   As Recordset 'Jet
'Dim DbPoi   As Recordset 'Jet

Dim bOnLoad As Byte
Dim bRefreshed As Byte
Dim lSonumber As Long
Dim iItem As Integer
Dim sRev As String
Dim sDesc As String

Dim sPkPart As String
Dim sPkComt As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbWcn_Click()
   GetWorkCenter
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   If Trim(cmbWcn) = "" Then cmbWcn = "ALL"
   GetWorkCenter
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4155
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdSel_Click()
   FillGrid
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   If optFrom.Value = vbChecked Then
      optFrom.Value = vbUnchecked
      Unload CustSh01a
   End If
   If optReport.Value = vbChecked Then
      optReport.Value = vbUnchecked
      Unload CustSh01b
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT RAREF,RARUN,RASO,RASOITEM,RASOREV,SONUMBER,SOTYPE," _
          & "SOCUST,SOPO,SOTEXT,CUREF,CUNICKNAME FROM RnalTable,SohdTable," _
          & "CustTable WHERE (RASO = SONUMBER AND SOCUST = CUREF) " _
          & "AND (RAREF= ? AND RARUN= ?)"

   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adInteger
   
   AdoQry.Parameters.Append AdoParameter1
   AdoQry.Parameters.Append ADOParameter2
   'Sizes 2/3/04 Janet AWI
   With Grid1
      .Rows = 2
      '.ColWidth(0) = 1100
      .ColWidth(0) = 750
      .ColWidth(1) = 400
      .ColWidth(2) = 2150
      .ColWidth(3) = 400
      '.ColWidth(4) = 600
      .ColWidth(4) = 350
      .ColWidth(5) = 500
      '.ColWidth(5) = 0
      .ColWidth(6) = 925
      .ColWidth(7) = 900
      .ColWidth(8) = 600
      .ColWidth(9) = 1000
      .ColWidth(10) = 1200
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .ColAlignment(4) = 0
      .ColAlignment(6) = 0
      ' .ColAlignment(5) = 0
      .ColAlignment(7) = 0
      ' .ColAlignment(7) = 1
      .ColAlignment(9) = 0
      .ColAlignment(10) = 0
      .row = 0
      .Col = 0
      .Text = "Work Ctr"
      .Col = 1
      .Text = "Prty"
      .Col = 2
      .Text = "Manufacturing Order"
      .Col = 3
      .Text = "Run"
      .Col = 4
      .Text = ""
      .Text = "OP"
      .Col = 5
      .Text = "Time "
      '.Text = "Status"
      .Col = 6
      .Text = "Sched Start"
      .Col = 7
      .Text = "Sched Com"
      .Col = 8
      .Text = "Qty"
      .Col = 9
      .Text = "Customer"
      .Col = 10
      .Text = "Sales Order"
   End With
   bRefreshed = 0
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoQry = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   Grid1.ToolTipText = "Contains Up To 300 Entries. " & Grid1.ToolTipText
   'Janet 12/2/04
   txtBeg = "01/01/" & Format(ES_SYSDATE, "yyyy")
   txtEnd = "12/31/" & Format(ES_SYSDATE, "yyyy")
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "Qry_FillWorkCentersAll"
   AddComboStr cmbWcn.hwnd, "ALL"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then cmbWcn = cmbWcn.List(0)
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Grid1_DblClick()
   On Error Resume Next
   optFrom.Value = vbChecked
   Grid1.Col = 0
   CustSh01a.cmbWcn = Grid1.Text
   Grid1.Col = 1
   CustSh01a.txtPri = Val(Grid1.Text)
   Grid1.Col = 2
   CustSh01a.lblMon = Grid1.Text
   Grid1.Col = 3
   CustSh01a.lblRun = Grid1.Text
   Grid1.Col = 4
   CustSh01a.lblOpno = Grid1.Text
   Grid1.Col = 5
   CustSh01a.lblSta = Grid1.Text
   Grid1.Col = 7
   CustSh01a.lblQty = Grid1.Text
   Grid1.Col = 0
   CustSh01a.Show
   CustSh01a.lblWcn = Grid1.Text
   
End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      On Error Resume Next
      optFrom.Value = vbChecked
      Grid1.Col = 0
      CustSh01a.cmbWcn = Grid1.Text
      Grid1.Col = 1
      CustSh01a.txtPri = Grid1.Text
      Grid1.Col = 2
      CustSh01a.lblMon = Grid1.Text
      Grid1.Col = 3
      CustSh01a.lblRun = Grid1.Text
      Grid1.Col = 4
      CustSh01a.lblOpno = Grid1.Text
      Grid1.Col = 0
      CustSh01a.Show
   End If
   
End Sub


Private Sub optData_Click()
   'if the data in the PopUp has been changed
   
End Sub

Private Sub optFrom_Click()
   'Shows edit
   'Refill when closing the 2nd form
   If optFrom.Value = vbUnchecked And optData.Value = vbChecked Then
      optData.Value = vbUnchecked
      FillGrid
   End If
   
End Sub

Private Sub optPrn_Click()
   BuildReport
   
End Sub


Private Sub optReport_Click()
   'Report box
   
End Sub



Private Sub txtBeg_DropDown()
    ShowCalendarEx Me
End Sub

Private Sub txtBeg_LostFocus()
    txtEnd = CheckDateEx(txtEnd)
End Sub

Private Sub txtDsc_Change()
   If Left(txtDsc, 6) = "*** Wo" Then
      txtDsc.ForeColor = ES_RED
   Else
      txtDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   txtEnd = CheckDateEx(txtEnd)
   
End Sub



Private Sub GetWorkCenter()
   Dim RdoCnt As ADODB.Recordset
   On Error GoTo DiaErr1
   If Trim(cmbWcn) <> "ALL" Then
      sSql = "SELECT WCNNUM,WCNDESC FROM WcntTable " _
             & "WHERE WCNREF='" & Compress(cmbWcn) & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_FORWARD)
      If bSqlRows Then
         With RdoCnt
            cmbWcn = "" & Trim(!WCNNUM)
            txtDsc = "" & Trim(!WCNDESC)
            ClearResultSet RdoCnt
         End With
      Else
         txtDsc = "*** Work Center Wasn't Found ***"
      End If
   Else
      txtDsc = "All Work Centers Selected"
   End If
   Set RdoCnt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getworkcen"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillGrid()
   Dim RdoGrd As ADODB.Recordset
   Dim b As Byte
   Dim iList As Integer
   Dim A As Integer
   Dim C As Integer
   Dim sCenter As String
   Dim sNick As String
   Dim sSales As String
   Dim sSoPon As String
   Grid1.Rows = 1
   Grid1.row = 0
   
   On Error Resume Next
   A = 10
   
   If Trim(cmbWcn) <> "ALL" Then sCenter = Compress(cmbWcn)
   sSql = "SELECT OPREF,OPRUN,OPNO,OPCENTER,OPSUHRS,OPQDATE, RUNREF,RUNNO,RUNSCHED," _
          & "RUNQTY,RUNPRIORITY,RUNSTATUS,PARTREF,PARTNUM FROM RnopTable," _
          & "RunsTable,PartTable " _
          & "WHERE (OPREF = RUNREF AND OPRUN = RUNNO AND OPREF = PARTREF) " _
          & "AND (OPCOMPLETE=0 AND OPCENTER LIKE '" & sCenter & "%' AND " _
          & "RUNSCHED >= '" & Format(txtBeg, "mm/dd/yy") & "' AND " _
          & "RUNSCHED <= '" & Format(txtEnd, "mm/dd/yy") & "') " _
          & "ORDER BY OPCENTER,RUNPRIORITY,RUNSCHED"
   MouseCursor 13
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      lblInfo = "Filling Grid."
      lblInfo.Visible = True
      lblInfo.Refresh
      prg1.Value = A
      prg1.Visible = True
      
      With RdoGrd
         Do Until .EOF
            iList = iList + 1
            C = C + 1
            If C = 10 Then
               A = A + 7
               C = 0
               If A > 95 Then A = 95
               prg1.Value = A
            End If
            sCenter = "" & Trim(!OPCENTER)
            b = GetThisCenter(sCenter)
            b = GetAllocations(Trim(!PartRef), !Runno, sNick, sSales, sSoPon)
            Grid1.Rows = iList + 1
            Grid1.Col = 0
            Grid1.row = iList
            Grid1.Text = sCenter
            
            Grid1.Col = 1
            Grid1.Text = Format$(!RUNPRIORITY, "00")
            
            Grid1.Col = 2
            Grid1.Text = "" & Trim(!PartNum)
            
            Grid1.Col = 3
            Grid1.Text = !Runno
            
            Grid1.Col = 4
            Grid1.Text = Format(!opNo, "##0")
            
            Grid1.Col = 5
            'Grid1.Text = "" & Trim(!RUNSTATUS)
            Grid1.Text = Format$(!OPSUHRS, "#0.0")
            
            Grid1.Col = 6
            Grid1.Text = "" & Format(!OPQDATE, "mm/dd/yy")
            
            Grid1.Col = 7
            Grid1.Text = "" & Format(!RUNSCHED, "mm/dd/yy")
            
            Grid1.Col = 8
            Grid1.Text = Format(!RUNQTY, "#####0")
            
            Grid1.Col = 9
            Grid1.Text = sNick
            
            Grid1.Col = 10
            Grid1.Text = sSales
            lblTotal = iList
            lblTotal.Refresh
            If iList > 300 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
      If Grid1.Rows > 1 Then Grid1.Enabled = True _
                                             Else Grid1.Enabled = False
      prg1.Value = 100
      Grid1.Col = 0
      Grid1.row = 1
      On Error Resume Next
      Grid1.SetFocus
   Else
      MouseCursor 0
      lblTotal = 0
      MsgBox "No Open Operations Found.", _
         vbInformation, Caption
   End If
   prg1.Visible = False
   lblInfo.Visible = False
   MouseCursor 0
   Set RdoGrd = Nothing
   
End Sub

Private Function GetAllocations(sMon As String, lRun As Long, sCnick As String, sSoNum As String, sSoCuPo As String) As Byte
   Dim RdoAlc As ADODB.Recordset
   AdoParameter1.Value = sMon
   ADOParameter2.Value = lRun
   bSqlRows = clsADOCon.GetQuerySet(RdoAlc, AdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoAlc
         sCnick = "" & Trim(!CUNICKNAME)
         sSoNum = "" & Trim(!SOTYPE) & Trim(!SoText) & "-" & Format(!RASOITEM, "##0") & Trim(!RASOREV)
         lSonumber = !RASO
         iItem = !RASOITEM
         sRev = !RASOREV
         sSoCuPo = "" & Trim(!SOPO)
         ClearResultSet RdoAlc
      End With
   Else
      sCnick = ""
      sSoNum = ""
      lSonumber = 0
      iItem = 0
      sRev = ""
   End If
   Set RdoAlc = Nothing
   
End Function


Private Function GetThisCenter(sPassedCenter) As Byte
   Dim RdoCnt As ADODB.Recordset
   sSql = "Qry_GetRoutCenter '" & sPassedCenter & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_FORWARD)
   If bSqlRows Then
      With RdoCnt
         sPassedCenter = "" & Trim(!WCNNUM)
         sDesc = "" & Trim(!WCNDESC)
         ClearResultSet RdoCnt
      End With
   Else
      sPassedCenter = ""
      sDesc = ""
   End If
   Set RdoCnt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthiscen"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub BuildReport()
   Dim RdoRpt As ADODB.Recordset
   Dim RdoCst As ADODB.Recordset
   Dim b As Byte
   Dim iList As Integer
   Dim A As Integer
   Dim C As Integer
   Dim cSoQty As Currency
   Dim cPrice As Currency
   Dim sCenter As String
   Dim sNick As String
   Dim sSales As String
   Dim sSoDate As String
   Dim sSoCmt As String
   Dim sSoPon As String
   '2/2/04
   A = 10
   If Trim(cmbWcn) <> "ALL" Then sCenter = Compress(cmbWcn)
   On Error Resume Next
   sSql = "TRUNCATE TABLE EsReportCustSh01"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "TRUNCATE TABLE EsReportCustSh01P"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "SELECT OPREF,OPRUN,OPNO,OPCENTER,OPSUHRS,OPQDATE,RUNREF,RUNNO,RUNSCHED," _
          & "RUNQTY,RUNPRIORITY,RUNSTATUS,PARTREF,PARTNUM FROM RnopTable," _
          & "RunsTable,PartTable " _
          & "WHERE (OPREF = RUNREF AND OPRUN = RUNNO AND OPREF = PARTREF) " _
          & "AND (OPCOMPLETE=0 AND OPCENTER LIKE '" & sCenter & "%' AND " _
          & "RUNSCHED >= '" & Format(txtBeg, "mm/dd/yy") & "' AND " _
          & "RUNSCHED <= '" & Format(txtEnd, "mm/dd/yy") & "') " _
          & "ORDER BY OPCENTER,RUNPRIORITY,RUNSCHED"
   MouseCursor 13
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      sSql = "SELECT * FROM EsReportCustSh01"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_KEYSET)
      lblInfo = "Building Report."
      lblInfo.Visible = True
      lblInfo.Refresh
      prg1.Value = A
      prg1.Visible = True
      With RdoRpt
         Do Until .EOF
            C = C + 1
            If C = 10 Then
               A = A + 7
               C = 0
               If A > 95 Then A = 95
               prg1.Value = A
            End If
            b = GetAllocations(Trim(!PartRef), !Runno, sNick, sSales, sSoPon)
            sCenter = "" & Trim(!OPCENTER)
            b = GetThisCenter(sCenter)
            b = GetSoInfo(cSoQty, cPrice, sSoDate, sSoCmt)
            b = GetPicks(Trim(!PartRef), !Runno)
            RdoCst.AddNew
            RdoCst!CenterDesc = "" & sDesc
            RdoCst!Center = "" & sCenter
            RdoCst!Priority = Format$(!RUNPRIORITY, "00")
            RdoCst!MoNum = "" & Trim(!PartNum)
            RdoCst!MORUN = !Runno
            RdoCst!MoOpNo = Format(!opNo, "#000")
            RdoCst!MoStatus = "" & Trim(!RUNSTATUS)
            RdoCst!MoStartDte = "" & Format(!OPQDATE, "mm/dd/yy")
            RdoCst!MoShed = "" & Format$(!RUNSCHED, "mm/dd/yy")
            RdoCst!moQty = !RUNQTY
            RdoCst!Customer = sNick
            RdoCst!SalesOrder = lSonumber
            RdoCst!SoItem = iItem
            RdoCst!SoitemRev = sRev
            RdoCst!SoText = sSales
            RdoCst!SoCustPo = sSoPon
            RdoCst!SoQty = cSoQty
            RdoCst!SoShip = sSoDate
            '8/15 to satisfy AWI sort
            sSoDate = Right(sSoDate, 2) & Left(sSoDate, 2) & Mid(sSoDate, 4, 2)
            RdoCst!SoShipSort = sSoDate & "-"
            RdoCst!SoComments = sSoCmt
            RdoCst!OpTime = Format(!OPSUHRS, "###0.000")
            RdoCst!SoDollars = cPrice * cSoQty
            RdoCst!PkPickPart = sPkPart
            RdoCst!PkPickComt = sPkComt
            RdoCst.Update
            'Create a dummy for the join
            sSql = "INSERT INTO EsReportCustSh01P (MoNum,MoRun) " _
                   & "VALUES('" & Trim(!PartNum) & "'," & !Runno & ")"
            clsADOCon.ExecuteSQL sSql
            b = GetPoInfo("" & Trim(!PartNum), !Runno)
            iList = iList + 1
            If iList > 300 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoCst
      End With
      On Error Resume Next
      prg1.Value = 100
      optReport.Value = vbChecked
      CustSh01b.cmbWcn = cmbWcn
      CustSh01b.txtDsc = txtDsc
      CustSh01b.txtBeg = txtBeg
      CustSh01b.txtEnd = txtEnd
      CustSh01b.Show
   Else
      MsgBox "No Report Data Available.", _
         vbInformation, Caption
   End If
   lblInfo.Visible = False
   prg1.Visible = False
   Set RdoRpt = Nothing
   Set RdoCst = Nothing
   
   MouseCursor 0
   
End Sub

Private Function GetSoInfo(cQuantity As Currency, cDollars As Currency, sDate As String, sComments As String) As Byte
   Dim RdoSoi As ADODB.Recordset
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITQTY,ITDOLLARS,ITSCHED,ITCOMMENTS FROM " _
          & "SoitTable WHERE ITSO=" & lSonumber & " AND ITNUMBER=" & iItem _
          & " AND ITREV='" & sRev & "'"
   If lSonumber > 0 Then
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoi, ES_FORWARD)
      If bSqlRows Then
         With RdoSoi
            cQuantity = Format(!ITQTY, ES_QuantityDataFormat)
            sDate = Format$(!ITSCHED, "mm/dd/yy")
            cDollars = Format(!ITDOLLARS, ES_QuantityDataFormat)
            sComments = "" & Trim(!ITCOMMENTS)
            ClearResultSet RdoSoi
         End With
      Else
         cQuantity = 0
         sDate = ""
         sComments = ""
      End If
   Else
      cQuantity = 0
      sDate = ""
      sComments = ""
   End If
   Set RdoSoi = Nothing
End Function

Private Function GetPoInfo(sMoRef As String, lRunno As Long) As Byte
   Dim RdoPos As ADODB.Recordset
   Dim RdoCst As ADODB.Recordset
   
   
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART," _
          & "PIPDATE,PIADATE,PIPQTY,PIAQTY,PIRUNPART,PIRUNNO," _
          & "PONUMBER,POVENDOR,PARTREF,PARTNUM,PADESC," _
          & "VEREF,VENICKNAME " _
          & "FROM PoitTable, PohdTable,PartTable,VndrTable " _
          & "WHERE (PINUMBER=PONUMBER AND PITYPE<>16 AND " _
          & "PARTREF=PIPART AND POVENDOR=VEREF) AND " _
          & "(PIRUNPART='" & Compress(sMoRef) & "' AND PIRUNNO=" & lRunno & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPos, ES_FORWARD)
   If bSqlRows Then
      sSql = "SELECT * FROM EsReportCustSh01P"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_KEYSET)
      With RdoPos
         Do Until .EOF
            RdoCst.AddNew
            RdoCst!MoNum = "" & Trim(sMoRef)
            RdoCst!MORUN = lRunno
            RdoCst!MoPoNum = Format$(!PINUMBER, "000000") & "-" & Format$(!PIITEM) & Trim(!PIREV)
            RdoCst!MoPoVendor = "" & Trim(!VENICKNAME)
            RdoCst!MoPoPart = "" & Trim(!PartNum)
            RdoCst!MoPoPartDesc = "" & Trim(!PADESC)
            RdoCst!MoPoPDate = Format$(!PIPDATE, "mm/dd/yy")
            RdoCst!MoPoADate = Format$(!PIADATE, "mm/dd/yy")
            RdoCst!MoPoPQty = Format(!PIPQTY, ES_QuantityDataFormat)
            RdoCst!MoPoAQty = Format(!PIAQTY, ES_QuantityDataFormat)
            RdoCst.Update
            .MoveNext
         Loop
         ClearResultSet RdoPos
      End With
      RdoCst.Close
   End If
   Set RdoPos = Nothing
   Set RdoCst = Nothing
End Function


Private Function GetPicks(MOPart As String, MORUN As Long) As Byte
   Dim RdoPck As ADODB.Recordset
   sSql = "select PARTREF,PARTNUM,PKPARTREF,PKMOPART,PKMORUN," _
          & "PKCOMT FROM PartTable,MopkTable WHERE (PARTREF=PKPARTREF " _
          & "AND PKMOPART='" & Trim(MOPart) & "' AND PKMORUN=" & MORUN & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_STATIC)
   If bSqlRows Then
      With RdoPck
         sPkPart = "" & Trim(!PartNum)
         sPkComt = "" & Trim(!PKCOMT)
         If Len(sPkComt) > 60 Then
            sPkComt = Left$(sPkComt, 60)
         End If
         ClearResultSet RdoPck
      End With
   Else
      sPkPart = ""
      sPkComt = ""
   End If
   Set RdoPck = Nothing
End Function
