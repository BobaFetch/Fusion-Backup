VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form diaGLp14a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Cash Balance (Report)"
   ClientHeight = 6180
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 7050
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 6180
   ScaleWidth = 7050
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdRes
      Caption = "&Reselect"
      Height = 315
      Left = 6000
      TabIndex = 12
      Top = 1080
      Width = 875
   End
   Begin VB.CommandButton cmdSel
      Caption = "&Select"
      Height = 315
      Left = 6000
      TabIndex = 11
      Top = 720
      Width = 875
   End
   Begin VB.Frame Frame1
      Height = 30
      Left = 240
      TabIndex = 8
      Top = 1560
      Width = 6615
   End
   Begin VB.ComboBox txtThr
      Height = 315
      Left = 1320
      Sorted = -1 'True
      TabIndex = 0
      Tag = "4"
      Top = 1080
      Width = 1095
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5280
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 6180
      FormDesignWidth = 7050
   End
   Begin VB.ComboBox cmbAct
      Height = 315
      Left = 1320
      TabIndex = 1
      Top = 360
      Width = 1695
   End
   Begin VB.CommandButton CmdCan
      Caption = "Close"
      Height = 435
      Left = 6000
      TabIndex = 3
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 2
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaGLp14a.frx":0000
      PictureDn = "diaGLp14a.frx":0146
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2
      Height = 3975
      Left = 240
      TabIndex = 7
      Top = 1680
      Width = 6615
      _ExtentX = 11668
      _ExtentY = 7011
      _Version = 393216
      FixedRows = 0
      FixedCols = 0
      AllowBigSelection = 0 'False
      Enabled = -1 'True
      FocusRect = 0
      HighLight = 0
      FillStyle = 1
      SelectionMode = 1
      AllowUserResizing = 1
   End
   Begin VB.Label lblBal
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1920
      TabIndex = 10
      Top = 5760
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ending Balance"
      Height = 285
      Index = 2
      Left = 240
      TabIndex = 9
      Top = 5760
      Width = 1305
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Through"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 6
      Top = 1080
      Width = 975
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1320
      TabIndex = 5
      Top = 720
      Width = 2775
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 4
      Top = 360
      Width = 945
   End
End
Attribute VB_Name = "diaGLp14a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*********************************************************************************
' diaGLp14a - Cash Account Balance
'
' Notes: Requested by JEVINT
'
' Created:  11/17/04 (nth)
' Revisions:
' 02/14/05 (nth) Remove crystal report and replaced with grid.  FYI this is the 3rd
'                iteration of this report.  JLH cannot make up his mind and will
'                not properly test.  I highly doubt this will see the light of day.
'
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbAct_Click()
   lblDsc = UpdateActDesc(cmbAct)
End Sub

Private Sub cmbAct_LostFocus()
   lblDsc = UpdateActDesc(cmbAct)
   If Trim(cmbAct) = "" Then
      cmbAct = "ALL"
   End If
   'lblBal = Format(CashAccountBalance(cmbAct, txtThr), CURRENCYMASK)
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdSel_Click()
   FillGrid
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   txtThr = Format(ES_SYSDATE, "mm/dd/yy")
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaGLp14a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   Dim rdoAct As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH=1"
   bSqlRows = GetDataSet(rdoAct)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      cmbAct.ListIndex = 0
      lblDsc = UpdateActDesc(cmbAct)
   End If
   Set rdoAct = Nothing
   Exit Sub
   DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.description = Err.description
   DoModuleErrors Me
End Sub

Private Sub txtThr_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtThr_LostFocus()
   txtThr = CheckDate(txtThr)
End Sub

Private Sub FillGrid()
   'sSql = "SELECT JINAME,GJPOST,JIDESC,JITRAN,JIREF,JIACCOUNT,JIDEB,JICRD,RECITEM " _
   '    & "FROM GjitTable INNER JOIN GjhdTable ON JINAME = GJNAME LEFT OUTER JOIN " _
   '    & "ArecTable ON JINAME = RECITEM AND JITRAN = RECTRAN AND " _
   '    & "JIREF = RECREF LEFT OUTER JOIN JrhdTable ON JINAME = MJGLJRNL " _
   '    & "WHERE (MJGLJRNL IS NULL) AND (JICLEAR IS NULL) AND (JIACCOUNT = '" & sCurAcct & "') " _
   '    & "and gjpost <= '" & txtDte & "' ORDER BY GJPOST"
   
   'sSql = "SELECT DISTINCT CACDATE,CACKAMT,CACHECKNO,CUNICKNAME,RECITEM,CACUST " _
   '    & "FROM CashTable INNER JOIN CustTable ON CACUST = CUREF LEFT OUTER JOIN " _
   '    & "ArecTable ON CACUST = RECCUST AND CACHECKNO = RECITEM " _
   '    & "WHERE (CACASHACCT = '" & sCurAcct & "') AND (CACLEAR IS NULL) AND (CACANCELED = 0) " _
   '    & "and carcdate <= '" & txtDte & "' ORDER BY CACDATE"
   
   'sSql = "SELECT CHKNUMBER,CHKAMOUNT,CHKMEMO,CHKPRINTDATE,CHKPOSTDATE,VENICKNAME," _
   '    & "RECITEM FROM VndrTable INNER JOIN ChksTable ON VndrTable.VEREF = " _
   '    & "CHKVENDOR LEFT OUTER JOIN ArecTable ON CHKACCT = RECACCOUNT AND " _
   '    & "CHKNUMBER = RECITEM WHERE (CHKCLEARDATE IS NULL) AND (CHKPRINTED = 1) " _
   '    & "AND (CHKVOID = 0) AND (CHKACCT = '" & sCurAcct & "') and chkpostdate  <='" & txtDte _
   '    & "' OR (CHKCLEARDATE IS NULL) " _
   '    & "AND (CHKACCT = '" & sCurAcct & "') AND " _
   '    & "ISNUMERIC(CHKNUMBER) = 1 and chkvoid = 0 and chkpostdate <= '" & txtDte & "' " _
   '    & "order by CAST(CHKNUMBER AS decimal)"
   
   On Error Resume Next
   RdoCon.Execute "Drop table #temptest", rdExecDirect
   
   RdoCon.Execute "Create table #temptest(mid integer, mname varchar(20))", rdExecDirect
   
End Sub
