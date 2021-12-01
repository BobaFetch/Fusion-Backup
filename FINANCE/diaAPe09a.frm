VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form diaAPe09a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Cash Reconciliation"
   ClientHeight = 7305
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 9630
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 7305
   ScaleWidth = 9630
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton optDis
      Height = 330
      Left = 1320
      Picture = "diaAPe09a.frx":0000
      Style = 1 'Graphical
      TabIndex = 42
      ToolTipText = "Display Selected Transaction"
      Top = 6480
      UseMaskColor = -1 'True
      Width = 495
   End
   Begin VB.CommandButton cmdMrk
      Height = 330
      Left = 120
      Picture = "diaAPe09a.frx":017E
      Style = 1 'Graphical
      TabIndex = 40
      ToolTipText = "Select All"
      Top = 6480
      UseMaskColor = -1 'True
      Width = 495
   End
   Begin VB.CommandButton cmdUnm
      Height = 330
      Left = 720
      Picture = "diaAPe09a.frx":0290
      Style = 1 'Graphical
      TabIndex = 39
      ToolTipText = "Deselect All"
      Top = 6480
      UseMaskColor = -1 'True
      Width = 495
   End
   Begin TabDlg.SSTab SSTab1
      Height = 4455
      Left = 120
      TabIndex = 9
      ToolTipText = "Uncleared Cash Receipts (Sorted By Date \ Receipt Number)"
      Top = 1920
      Width = 9375
      _ExtentX = 16536
      _ExtentY = 7858
      _Version = 393216
      Style = 1
      TabHeight = 520
      TabCaption(0) = "&Checks and Payments"
      TabPicture(0) = "diaAPe09a.frx":030A
      Tab(0).ControlEnabled = -1 'True
      Tab(0).Control(0) = "Grid1"
      Tab(0).Control(0).Enabled = 0 'False
      Tab(0).ControlCount = 1
      TabCaption(1) = "&Deposits and Credits"
      TabPicture(1) = "diaAPe09a.frx":0326
      Tab(1).ControlEnabled = 0 'False
      Tab(1).Control(0) = "Grid2"
      Tab(1).ControlCount = 1
      TabCaption(2) = "&GL Entries"
      TabPicture(2) = "diaAPe09a.frx":0342
      Tab(2).ControlEnabled = 0 'False
      Tab(2).Control(0) = "Grid3"
      Tab(2).ControlCount = 1
      Begin MSFlexGridLib.MSFlexGrid Grid3
         Height = 3975
         Left = -74880
         TabIndex = 41
         ToolTipText = "Uncleared GL Entries (Sorted By Post Date)"
         Top = 360
         Width = 9135
         _ExtentX = 16113
         _ExtentY = 7011
         _Version = 393216
         FocusRect = 0
         HighLight = 0
         FillStyle = 1
         SelectionMode = 1
         AllowUserResizing = 1
      End
      Begin MSFlexGridLib.MSFlexGrid Grid2
         Height = 3975
         Left = -74880
         TabIndex = 11
         Top = 360
         Width = 9135
         _ExtentX = 16113
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
      Begin MSFlexGridLib.MSFlexGrid Grid1
         Height = 3975
         Left = 120
         TabIndex = 10
         ToolTipText = "Uncleared Checks (Sorted By Check Number)"
         Top = 360
         Width = 9135
         _ExtentX = 16113
         _ExtentY = 7011
         _Version = 393216
         FixedRows = 0
         FixedCols = 0
         AllowBigSelection = 0 'False
         FocusRect = 0
         HighLight = 0
         FillStyle = 1
         SelectionMode = 1
         AllowUserResizing = 1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name = "MS Sans Serif"
         Size = 8.25
         Charset = 0
         Weight = 400
         Underline = 0 'False
         Italic = 0 'False
         Strikethrough = 0 'False
         EndProperty
      End
   End
   Begin VB.ComboBox cmbIntAct
      Height = 315
      Left = 5040
      TabIndex = 8
      Top = 1440
      Width = 1455
   End
   Begin VB.ComboBox cmbSerAct
      Height = 315
      Left = 5040
      TabIndex = 5
      Top = 1080
      Width = 1455
   End
   Begin VB.ComboBox txtIntDte
      Height = 315
      Left = 3000
      TabIndex = 7
      Top = 1440
      Width = 1095
   End
   Begin VB.TextBox txtInt
      Height = 285
      Left = 1320
      TabIndex = 6
      Tag = "1"
      Top = 1440
      Width = 1095
   End
   Begin VB.TextBox txtSer
      Height = 285
      Left = 1320
      TabIndex = 3
      Tag = "1"
      Top = 1080
      Width = 1095
   End
   Begin VB.ComboBox txtSerDte
      Height = 315
      Left = 3000
      TabIndex = 4
      Top = 1080
      Width = 1095
   End
   Begin VB.ComboBox cmbAct
      Height = 315
      Left = 1080
      TabIndex = 0
      Top = 120
      Width = 1575
   End
   Begin VB.TextBox txtend
      Height = 285
      Left = 6960
      TabIndex = 2
      Tag = "2"
      Top = 480
      Width = 1215
   End
   Begin VB.TextBox txtBeg
      Height = 285
      Left = 6960
      TabIndex = 1
      Tag = "1"
      Top = 120
      Width = 1215
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 8640
      TabIndex = 14
      TabStop = 0 'False
      ToolTipText = "Save And Exit"
      Top = 120
      Width = 875
   End
   Begin VB.CommandButton cmdRec
      Caption = "&Reconcile "
      Height = 315
      Left = 8640
      TabIndex = 12
      ToolTipText = "Reconcile Marked Transactions"
      Top = 600
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 13
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
      PictureUp = "diaAPe09a.frx":035E
      PictureDn = "diaAPe09a.frx":04A4
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4320
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 7305
      FormDesignWidth = 9630
   End
   Begin Threed.SSFrame SSFrame1
      Height = 30
      Index = 2
      Left = 120
      TabIndex = 35
      Top = 960
      Width = 9375
      _Version = 65536
      _ExtentX = 16536
      _ExtentY = 53
      _StockProps = 14
      Caption = "SSFrame1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.26
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Index = 2
      Left = 6600
      TabIndex = 38
      Top = 1440
      Width = 2775
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Index = 1
      Left = 6600
      TabIndex = 37
      Top = 1080
      Width = 2775
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Index = 0
      Left = 1080
      TabIndex = 36
      Top = 480
      Width = 2775
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account"
      Height = 255
      Index = 18
      Left = 4320
      TabIndex = 34
      Top = 1440
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account"
      Height = 255
      Index = 17
      Left = 4320
      TabIndex = 33
      Top = 1080
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Date"
      Height = 255
      Index = 16
      Left = 2520
      TabIndex = 32
      Top = 1440
      Width = 735
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Date"
      Height = 255
      Index = 15
      Left = 2520
      TabIndex = 31
      Top = 1080
      Width = 735
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Interest Earned"
      Height = 255
      Index = 14
      Left = 120
      TabIndex = 30
      Top = 1440
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Service Charge"
      Height = 255
      Index = 13
      Left = 120
      TabIndex = 29
      Top = 1080
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Marked as Cleared:"
      Height = 255
      Index = 24
      Left = 3480
      TabIndex = 28
      Top = 6480
      Width = 1935
   End
   Begin VB.Label lblDiffBal
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Height = 255
      Left = 7800
      TabIndex = 27
      Top = 6960
      Width = 1095
   End
   Begin VB.Label lblClearBal
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Height = 255
      Left = 7800
      TabIndex = 26
      Top = 6720
      Width = 1095
   End
   Begin VB.Label lblEndBal
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Height = 255
      Left = 7800
      TabIndex = 25
      Top = 6480
      Width = 1095
   End
   Begin VB.Label lblTotalChk
      Alignment = 1 'Right Justify
      Height = 255
      Left = 3480
      TabIndex = 24
      Top = 6960
      Width = 375
   End
   Begin VB.Label lblTotalDep
      Alignment = 1 'Right Justify
      Height = 255
      Left = 3480
      TabIndex = 23
      Top = 6720
      Width = 375
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Checks and Payments"
      Height = 255
      Index = 23
      Left = 3960
      TabIndex = 22
      Top = 6960
      Width = 1695
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Deposits and Credits"
      Height = 255
      Index = 22
      Left = 3960
      TabIndex = 21
      Top = 6720
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Difference"
      Height = 255
      Index = 21
      Left = 6000
      TabIndex = 20
      Top = 6960
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Cleared Balance"
      Height = 255
      Index = 20
      Left = 6000
      TabIndex = 19
      Top = 6720
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ending Balance"
      Height = 255
      Index = 19
      Left = 6000
      TabIndex = 18
      Top = 6480
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account"
      Height = 255
      Index = 6
      Left = 360
      TabIndex = 17
      Top = 120
      Width = 1455
   End
   Begin VB.Image imgdInc
      Height = 180
      Left = 3600
      Picture = "diaAPe09a.frx":05EA
      Top = 0
      Visible = 0 'False
      Width = 255
   End
   Begin VB.Image imgInc
      Height = 180
      Left = 3960
      Picture = "diaAPe09a.frx":089C
      Top = 0
      Visible = 0 'False
      Width = 255
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ending Balance"
      Height = 255
      Index = 0
      Left = 5640
      TabIndex = 16
      Top = 480
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Opening Balance"
      Height = 255
      Index = 2
      Left = 5640
      TabIndex = 15
      Top = 120
      Width = 1335
   End
End
Attribute VB_Name = "diaAPe09a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaAPe09a - Account Reconcilation
'
' Created (nth)
' Revisions
' 10/16/03 (nth) Added GL entries tab per wck.
' 10/23/03 (nth) Added "enter" to grid per wck.
' 06/15/04 (nth) Exclude voided cash receipts and checks.
' 06/16/04 (nth) Added save and restore ArecTable remember your work when you leave.
' 07/27/04 (nth) Added transaction fee to deposits and credits
' 08/10/04 (nth) Corrected CR total multiplying number of invoices
' 08/25/04 (nth) Changed check sort number check numbers then alpha numeric check numbers
' 09/22/04 (nth) Exclude voided checks
' 10/04/04 (nth) Fixed not compressing customer on reconcile deposits and credits per JEVINT
'
'*************************************************************************************

Option Explicit

' Numeric mask used in this transaction formats up to 1 billion $.
Const CURRENCYMASK = "#,###,###,###,##0.00"
Const DATEMASK = "mm/dd/yy"

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sCurAcct As String

' GL Grid
Dim iGLTran() As Integer
Dim iGLRef() As Integer

' Running Total
Dim cOpenBal As Currency
Dim cEndBal As Currency

' Accounts
Dim sCrCashAcct As String
Dim sCrDiscAcct As String
Dim sCrExpAcct As String
Dim sSJARAcct As String
Dim sCrRevAcct As String
Dim sCrCommAcct As String

' Key Handeling
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub Reconcile()
   Dim i As Integer
   Dim iDep As Integer
   Dim iChk As Integer
   Dim sJournalID1 As String 'PJ
   Dim sJournalID2 As String 'CR
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim sVendor As String
   Dim rdoAct As rdoResultset
   Dim sCrCheck As String
   Dim sMsg As String
   Dim bResponse As Byte
   Dim sTemp As String
   Dim sGL As String
   Dim sPst As String
   
   On Error GoTo DiaErr1
   
   sMsg = "Reconcile Cash Account " & Trim(cmbAct) & " ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If
   
   On Error Resume Next
   'On Error GoTo 0
   RdoCon.BeginTrans
   
   Dim sToday As String
   sToday = Format(GetServerDateTime, "mm/dd/yyyy")
   
   
   If CCur(txtSer) > 0 Then
      sGL = "SC" & Format(txtSerDte, "yyyymmdd")
      sPst = GetFYPeriodEnd(CDate(txtSerDte))
      
      sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJOPEN,GJPOST) " _
             & "VALUES('" & sGL & "','SERVICE CHARGE','" _
             & txtSerDte & "','" & sPst & "')"
      RdoCon.Execute sSql
      
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIACCOUNT,JICRD,JIDATE,JICLEAR) " _
             & "VALUES('" & sGL & "',1,1,'" & Compress(cmbAct) & "'," _
             & CCur(txtSer) & ",'" & sToday & "','" & sToday & "')"
      RdoCon.Execute sSql
      
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIACCOUNT,JIDEB,JIDATE,JICLEAR) " _
             & "VALUES('" & sGL & "',1,2,'" & Compress(cmbSerAct) & "'," _
             & CCur(txtSer) & ",'" & sToday & "','" & sToday & "')"
      RdoCon.Execute sSql
   End If
   
   If CCur(txtInt) > 0 Then
      sGL = "IE" & Format(txtIntDte, "yyyymmdd")
      sPst = GetFYPeriodEnd(txtIntDte)
      
      sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJOPEN,GJPOST) " _
             & "VALUES('" & sGL & "','INTEREST EARNED','" _
             & txtIntDte & "','" & sPst & "')"
      RdoCon.Execute sSql
      
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIACCOUNT,JIDEB,JIDATE,JICLEAR) " _
             & "VALUES('" & sGL & "',1,1,'" & Compress(cmbAct) & "'," _
             & CCur(txtInt) & ",'" & sToday & "','" & sToday & "')"
      RdoCon.Execute sSql
      
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIACCOUNT,JICRD,JIDATE,JICLEAR) " _
             & "VALUES('" & sGL & "',1,2,'" & Compress(cmbSerAct) & "'," _
             & CCur(txtInt) & ",'" & sToday & "','" & sToday & "')"
      RdoCon.Execute sSql
   End If
   
   ' checks
   With Grid1
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         If .CellPicture = imgInc Then
            .Col = 2
            sSql = "UPDATE ChksTable Set CHKCLEARDATE='" & sToday _
                   & "' WHERE CHKNUMBER='" & Trim(.Text) & "'"
            RdoCon.Execute sSql
            .Col = 0
         End If
      Next
   End With
   
   ' deposits and credits
   With Grid2
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         If .CellPicture = imgInc Then
            .Col = 3
            sTemp = Compress(.Text)
            .Col = 2
            sSql = "UPDATE CashTable Set CACLEAR='" & sToday _
                   & "' WHERE CACUST='" & sTemp _
                   & "' AND CACHECKNO='" & Trim(.Text) & "'"
            RdoCon.Execute sSql
            .Col = 0
         End If
      Next
   End With
   
   ' gl
   With Grid3
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         If .CellPicture = imgInc Then
            .Col = 2
            sSql = "UPDATE GjitTable SET JICLEAR = '" & sToday _
                   & "' WHERE JINAME = '" & Trim(.Text) & "' AND JITRAN = " _
                   & iGLTran(i) & " AND JIREF = " & iGLRef(i)
            RdoCon.Execute sSql
            .Col = 0
         End If
      Next
   End With
   
   ' Record ending balance, date , and who did it in account record
   sSql = "UPDATE GlacTable SET GLRECDATE = '" & sToday & "',GLRECBAL = " _
          & CCur(txtEnd) & ",GLRECBY = '" & Secure.UserInitials & "'" _
          & " WHERE GLACCTREF = '" & sCurAcct & "'"
   RdoCon.Execute sSql
   
   sSql = "DELETE FROM ArecTable"
   RdoCon.Execute sSql
   
   If Err = 0 Then
      RdoCon.CommitTrans
      Sysmsg "Account Successfully Reconciled.", True
      txtBeg = txtEnd
   Else
      RdoCon.RollbackTrans
      MsgBox "Could Not Reconcile Account.", _
         vbExclamation, Caption
   End If
   
   MouseCursor 0
   cmbAct.Enabled = True
   cmbAct.SetFocus
   Exit Sub
   
   DiaErr1:
   sProcName = "Reconcile"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbAct_Click()
   GetAccount
End Sub

Private Sub cmbAct_LostFocus()
   If Not bCancel Then
      GetAccount
   End If
End Sub

Private Sub cmbIntAct_Click()
   lblDsc(2) = UpdateActDesc(cmbIntAct)
End Sub

Private Sub cmbIntAct_LostFocus()
   lblDsc(2) = UpdateActDesc(cmbIntAct)
End Sub

Private Sub cmbSerAct_Click()
   lblDsc(1) = UpdateActDesc(cmbSerAct)
End Sub

Private Sub cmbSerAct_LostFocus()
   lblDsc(1) = UpdateActDesc(cmbSerAct)
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
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdMrk_Click()
   Mark True
End Sub

Private Sub cmdRec_Click()
   Reconcile
   FillGrid
   UpdateTotals
End Sub

Private Sub cmdUnm_Click()
   Mark False
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      FillAccounts
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim sItem As String
   SetDiaPos Me
   FormatControls
   
   With Grid1
      .RowHeightMin = 300
      .FixedRows = 1
      .Rows = 1
      .FixedCols = 0
      .Cols = 6
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 1500
      .ColWidth(3) = 1500
      .ColWidth(4) = 2700
      .ColWidth(5) = 1500
      .Col = 1
      .Text = "Date"
      .Col = 2
      .Text = "Check Number"
      .Col = 3
      .Text = "Vendor"
      .Col = 4
      .Text = "Memo"
      .Col = 5
      .Text = "Amount"
   End With
   
   With Grid2
      .RowHeightMin = 300
      .FixedRows = 1
      .Rows = 1
      .FixedCols = 0
      .Cols = 6
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 1500
      .ColWidth(3) = 1500
      .ColWidth(4) = 1500
      .ColWidth(5) = 1500
      .Col = 1
      .Text = "Date"
      .Col = 2
      .Text = "Check Number"
      .Col = 3
      .Text = "Customer"
      .Col = 4
      .Text = "Fee"
      .Col = 5
      .Text = "Amount"
   End With
   
   With Grid3
      .RowHeightMin = 300
      .FixedRows = 1
      .Rows = 1
      .FixedCols = 0
      .Cols = 6
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 1500
      .ColWidth(3) = 2700
      .ColWidth(4) = 1500
      .ColWidth(5) = 1500
      .Col = 1
      .Text = "Date"
      .Col = 2
      .Text = "Journal"
      .Col = 3
      .Text = "Description"
      .Col = 4
      .Text = "Debit"
      .Col = 5
      .Text = "Credit"
   End With
   
   optDis.Enabled = False
   txtSerDte = Format(Now, "mm/dd/yy")
   txtIntDte = Format(Now, "mm/dd/yy")
   txtSer = "0.00"
   txtInt = "0.00"
   txtBeg = "0.00"
   txtEnd = "0.00"
   sCurAcct = ""
   
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaGLp08a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillGrid()
   Dim RdoChk As rdoResultset
   Dim rdoDep As rdoResultset
   Dim rdoGL As rdoResultset
   Dim rdoFee As rdoResultset
   Dim sItem As String
   Dim i As Integer
   Dim sDate As String
   
   On Error GoTo DiaErr1
   
   MouseCursor 13
   
   Grid1.Rows = 1
   Grid2.Rows = 1
   Grid3.Rows = 1
   
   i = 1
   ' Fill GL Entries
   sSql = "SELECT JINAME,GJPOST,JIDESC,JITRAN,JIREF,JIACCOUNT,JIDEB,JICRD,RECITEM " _
          & "FROM GjitTable INNER JOIN GjhdTable ON JINAME = GJNAME LEFT OUTER JOIN " _
          & "ArecTable ON JINAME = RECITEM AND JITRAN = RECTRAN AND " _
          & "JIREF = RECREF LEFT OUTER JOIN JrhdTable ON JINAME = MJGLJRNL " _
          & "WHERE (MJGLJRNL IS NULL) AND (JICLEAR IS NULL) AND (JIACCOUNT = '" & sCurAcct & "') " _
          & "ORDER BY GJPOST"
   bSqlRows = GetDataSet(rdoGL)
   If bSqlRows Then
      With rdoGL
         While Not .EOF
            sItem = Chr(9) & Format(!GJPOST, "mm/dd/yy") _
                    & Chr(9) & " " & Trim(!JINAME) _
                    & Chr(9) & " " & Trim(!JIDESC) _
                    & Chr(9) & Format(!JIDEB, CURRENCYMASK) _
                    & Chr(9) & Format(!JICRD, CURRENCYMASK)
            Grid3.AddItem sItem
            Grid3.Row = i
            Grid3.Col = 0
            Grid3.CellPictureAlignment = flexAlignCenterCenter
            If IsNull(!RECITEM) Then
               Set Grid3.CellPicture = imgdInc
            Else
               Set Grid3.CellPicture = imgInc
            End If
            ReDim Preserve iGLTran(i)
            ReDim Preserve iGLRef(i)
            iGLTran(i) = !JITRAN
            iGLRef(i) = !JIREF
            i = i + 1
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set rdoGL = Nothing
   
   ' Fill Deposit/Credit Grid
   sSql = "SELECT DISTINCT CACDATE,CACKAMT,CACHECKNO,CUNICKNAME,RECITEM,CACUST " _
          & "FROM CashTable INNER JOIN CustTable ON CACUST = CUREF LEFT OUTER JOIN " _
          & "ArecTable ON CACUST = RECCUST AND CACHECKNO = RECITEM " _
          & "WHERE (CACASHACCT = '" & sCurAcct & "') AND (CACLEAR IS NULL) AND (CACANCELED = 0) " _
          & "ORDER BY CACDATE"
   bSqlRows = GetDataSet(rdoDep, ES_FORWARD)
   i = 1
   If bSqlRows Then
      With rdoDep
         While Not .EOF
            ' trans fee
            sSql = "SELECT SUM(DCDEBIT) FROM JritTable WHERE DCCHECKNO = '" _
                   & Trim(!CACHECKNO) & "' AND DCCUST = '" & Trim(!CACUST) _
                   & "' AND DCDESC = 'TRANS FEE'"
            bSqlRows = GetDataSet(rdoFee)
            sItem = Chr(9) & Format(!CACDATE, "mm/dd/yy") _
                    & Chr(9) & " " & !CACHECKNO _
                    & Chr(9) & " " & !CUNICKNAME _
                    & Chr(9) & Format(rdoFee.rdoColumns(0), CURRENCYMASK) _
                    & Chr(9) & Format(.rdoColumns(1), CURRENCYMASK)
            Set rdoFee = Nothing
            Grid2.AddItem sItem
            Grid2.Row = i
            Grid2.Col = 0
            Grid2.CellPictureAlignment = flexAlignCenterCenter
            If IsNull(!RECITEM) Then
               Set Grid2.CellPicture = imgdInc
            Else
               Set Grid2.CellPicture = imgInc
            End If
            i = i + 1
            .MoveNext
         Wend
      End With
   End If
   Set rdoDep = Nothing
   
   ' Fill Check Grid
   
   ' Numeric check numbers
   sSql = "SELECT CHKNUMBER,CHKAMOUNT,CHKMEMO,CHKPRINTDATE,CHKPOSTDATE,VENICKNAME," _
          & "RECITEM FROM VndrTable INNER JOIN ChksTable ON VndrTable.VEREF = " _
          & "CHKVENDOR LEFT OUTER JOIN ArecTable ON CHKACCT = RECACCOUNT AND " _
          & "CHKNUMBER = RECITEM WHERE (CHKCLEARDATE IS NULL) AND (CHKPRINTED = 1) " _
          & "AND (CHKVOID = 0) AND (CHKACCT = '" & sCurAcct & "') OR (CHKCLEARDATE IS NULL) " _
          & "AND (CHKACCT = '" & sCurAcct & "') AND (CHKTYPE = 1) AND ISNUMERIC(CHKNUMBER) = 1 and chkvoid = 0 " _
          & "order by CAST(CHKNUMBER AS decimal)"
   bSqlRows = GetDataSet(RdoChk)
   If bSqlRows Then
      FillCheckGrid RdoChk
   End If
   Set RdoChk = Nothing
   
   ' Alpha/Numeric check numbers
   sSql = "SELECT CHKNUMBER,CHKAMOUNT,CHKMEMO,CHKPRINTDATE,CHKPOSTDATE,VENICKNAME," _
          & "RECITEM FROM VndrTable INNER JOIN ChksTable ON VndrTable.VEREF = " _
          & "CHKVENDOR LEFT OUTER JOIN ArecTable ON CHKACCT = RECACCOUNT AND " _
          & "CHKNUMBER = RECITEM WHERE (CHKCLEARDATE IS NULL) AND (CHKPRINTED = 1) " _
          & "AND (CHKVOID = 0) AND (CHKACCT = '" & sCurAcct & "') and ISNUMERIC(CHKNUMBER) = 0 OR (CHKCLEARDATE IS NULL) " _
          & "AND (CHKACCT = '" & sCurAcct & "') AND (CHKTYPE = 1) AND ISNUMERIC(CHKNUMBER) = 0 and chkvoid = 0 order by chknumber "
   bSqlRows = GetDataSet(RdoChk)
   If bSqlRows Then
      FillCheckGrid RdoChk
   End If
   Set RdoChk = Nothing
   MouseCursor 0
   Exit Sub
   DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCheckGrid(RdoChk As rdoResultset)
   Dim sItem As String
   Dim i As Integer
   Dim sDate As String
   On Error GoTo DiaErr1
   With RdoChk
      sItem = ""
      i = Grid1.Rows
      While Not .EOF
         sDate = Format("" & Trim(!CHKPOSTDATE), "mm/dd/yy")
         sItem = Chr(9) & " " & sDate _
                 & Chr(9) & " " & !CHKNUMBER _
                 & Chr(9) & " " & !VENICKNAME _
                 & Chr(9) & " " & Trim(!chkMemo) _
                 & Chr(9) & Format(!CHKAMOUNT, CURRENCYMASK)
         Grid1.AddItem sItem
         Grid1.Row = i
         Grid1.Col = 0
         Grid1.CellPictureAlignment = flexAlignCenterCenter
         If IsNull(!RECITEM) Then
            Set Grid1.CellPicture = imgdInc
         Else
            Set Grid1.CellPicture = imgInc
         End If
         i = i + 1
         .MoveNext
      Wend
      .Cancel
   End With
   Set RdoChk = Nothing
   Exit Sub
   DiaErr1:
   sProcName = "fillchec"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Grid1_Click()
   GridClick Grid1
   Highlight Grid1
   UpdateTotals
End Sub

Private Sub FillAccounts()
   ' Fill account combo
   ' Need to add account descriptions
   Dim rdoAct As rdoResultset
   Dim b As Byte
   
   b = GetCashAccounts()
   If b = 3 Then
      MouseCursor 0
      MsgBox "One Or More Cash Accounts Are Not Active." & vbCr _
         & "Please Set All Cash Accounts In The " & vbCr _
         & "System Setup, Administration Section.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   
   ' Accounts
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH=1"
   bSqlRows = GetDataSet(rdoAct)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      cmbAct.ListIndex = 0
      lblDsc(0) = UpdateActDesc(cmbAct)
   Else
      ' Multiple cash accounts not found so use the default cash account
      cmbAct = sCrCashAcct
      lblDsc(0) = UpdateActDesc(cmbAct)
   End If
   Set rdoAct = Nothing
   
   ' Other Accounts
   sSql = "Qry_FillLowAccounts"
   bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         Do Until .EOF
            AddComboStr cmbSerAct.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr cmbIntAct.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
      End With
      cmbSerAct.ListIndex = 0
      cmbIntAct.ListIndex = 0
   End If
   Set rdoAct = Nothing
End Sub

'Local Errors - from FillCombo
'Flash with 1 = no accounts nec, 2 = OK, 3 = Not enough accounts

Public Function GetCashAccounts() As Byte
   Dim rdocsh As rdoResultset
   Dim i As Integer
   Dim b As Byte
   sSql = "SELECT COGLVERIFY,COCRCASHACCT,COCRDISCACCT,COSJARACCT," _
          & "COCRCOMMACCT,COCRREVACCT,COCREXPACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = GetDataSet(rdocsh, ES_FORWARD)
   sProcName = "getcashacct"
   If bSqlRows Then
      With rdocsh
         For i = 1 To 6
            If "" & Trim(.rdoColumns(i)) = "" Then
               b = 1
               Exit For
            End If
         Next
         sCrCashAcct = "" & Trim(!COCRCASHACCT)
         sCrDiscAcct = "" & Trim(!COCRDISCACCT)
         sSJARAcct = "" & Trim(!COSJARACCT)
         sCrCommAcct = "" & Trim(!COCRCOMMACCT)
         sCrRevAcct = "" & Trim(!COCRREVACCT)
         sCrExpAcct = "" & Trim(!COCREXPACCT)
         .Cancel
         If b = 1 Then GetCashAccounts = 3 Else GetCashAccounts = 2
      End With
   Else
      GetCashAccounts = 0
   End If
   Set rdocsh = Nothing
End Function

Private Sub UpdateTotals()
   Dim i As Integer
   Dim iTotalDep As Integer
   Dim iTotalChk As Integer
   Dim iTotalGL As Integer
   Dim cBalance As Currency
   Dim cSnI As Currency
   Dim iTemp As Integer
   
   ' sum checks
   With Grid1
      iTemp = .Row
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         If .CellPicture = imgInc Then
            .Col = 5
            cBalance = cBalance + CCur(.Text)
            iTotalChk = iTotalChk + 1
            .Col = 0
         End If
      Next
      .Row = iTemp
   End With
   
   ' deposits
   With Grid2
      iTemp = .Row
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         If .CellPicture = imgInc Then
            .Col = 5
            cBalance = cBalance - CCur(.Text)
            iTotalDep = iTotalDep + 1
            .Col = 0
         End If
      Next
      .Row = iTemp
   End With
   
   ' sum gl
   With Grid3
      iTemp = .Row
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         If .CellPicture = imgInc Then
            .Col = 4
            If .Text > 0 Then
               cBalance = cBalance + .Text
            Else
               .Col = 5
               cBalance = cBalance - .Text
            End If
            iTotalGL = iTotalGL + 1
            .Col = 0
         End If
      Next
      .Row = iTemp
   End With
   
   ' Update totals
   
   cSnI = CCur(txtInt) - CCur(txtSer)
   
   lblEndBal = Format(txtEnd, CURRENCYMASK)
   lblClearBal = Format(cBalance + cSnI, CURRENCYMASK)
   lblDiffBal = Format((cBalance + cSnI) + CCur(txtBeg) - _
                CCur(txtEnd), CURRENCYMASK)
   
   lblTotalChk = iTotalChk
   lblTotalDep = iTotalDep
   
   ' Highlight reconcile button if differnce = 0
   With cmdRec
      If Val(lblDiffBal) = 0 Then
         .Enabled = True
      Else
         .Enabled = False
      End If
   End With
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Or KeyAscii = 13 Then
      Grid1_Click
   End If
End Sub

Private Sub Grid2_Click()
   'Dim b As Byte
   'With Grid2
   '    .Row = .RowSel
   '    For b = 0 To .Cols - 1
   '         .Col = b
   '         .CellBackColor = vbBlack
   '    Next
   'End With
   GridClick Grid2
   Highlight Grid2
   UpdateTotals
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Or KeyAscii = 13 Then
      Grid2_Click
   End If
End Sub

Private Sub Grid3_Click()
   GridClick Grid3
   Highlight Grid3
   UpdateTotals
End Sub

Private Sub Grid3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Or KeyAscii = 13 Then
      Grid3_Click
   End If
End Sub

Private Sub optDis_Click()
   Dim sCustomer As String
   Dim sItem As String
   
   Select Case SSTab1.Tab
      Case 0
         With Grid1
            .Col = 2
         End With
      Case 1
         With Grid2
            .Col = 3
            sCustomer = Compress(.Text)
            .Col = 2
            sItem = Trim(.Text)
            diaARp09a.bRemote = True
            diaARp09a.PrintReport sCustomer, sItem
            Unload diaARp09a
         End With
      Case 2
         With Grid3
            .Col = 2
            sItem = Trim(.Text)
            diaGLp03a.bRemote = True
            diaGLp03a.PrintReport sItem
            Unload diaGLp03a
         End With
   End Select
End Sub

Private Sub SSTab1_GotFocus()
   Select Case SSTab1.Tab
      Case 0
         With Grid1
            .SetFocus
         End With
         optDis.Enabled = False
      Case 1
         With Grid2
            .SetFocus
         End With
         optDis.Enabled = True
      Case 2
         With Grid3
            .SetFocus
         End With
         optDis.Enabled = True
   End Select
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = Format(txtBeg, CURRENCYMASK)
   If Trim(txtBeg) = "" Then txtBeg = "0.00"
   cOpenBal = CCur(txtBeg)
   UpdateTotals
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = Format(txtEnd, CURRENCYMASK)
   If Trim(txtEnd) = "" Then txtEnd = "0.00"
   cEndBal = CCur(txtEnd)
   UpdateTotals
End Sub

Private Sub txtInt_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtInt_LostFocus()
   txtInt = CheckLen(txtInt, 11)
   txtInt = Format(txtInt, CURRENCYMASK)
   If Trim(txtInt) = "" Then txtInt = "0.00"
   UpdateTotals
End Sub

Private Sub txtIntDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtSer_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtSer_LostFocus()
   txtSer = CheckLen(txtSer, 11)
   txtSer = Format(txtSer, CURRENCYMASK)
   If Trim(txtSer) = "" Then txtSer = "0.00"
   UpdateTotals
End Sub

Private Sub txtSerDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub GridClick(Whichgrid As MSFlexGrid)
   ' bType doc
   ' 0 = Checks
   ' 1 = Cash Receipt
   ' 2 = GL
   Dim bType As Byte
   Dim sCust As String
   Dim iTran As Integer
   Dim iRef As Integer
   On Error GoTo DiaErr1
   bType = CByte(Right(Whichgrid.Name, 1)) - 1
   With Whichgrid
      If .Rows > 0 And .Row > 0 And .Col = 0 Then
         If bType = 1 Then
            .Col = 3
            sCust = Compress(.Text)
         End If
         .Col = 0
         If .CellPicture = imgdInc Then
            If bType = 2 Then
               iTran = iGLTran(.Row)
               iRef = iGLRef(.Row)
            End If
            .Col = 2
            sSql = "INSERT INTO ArecTable(RECITEM,RECITEMTYPE," _
                   & "RECACCOUNT,RECBY,RECCUST,RECTRAN,RECREF) VALUES ('" _
                   & Trim(.Text) & "'," & bType _
                   & ",'" & sCurAcct & "','" & Secure.UserInitials _
                   & "','" & sCust & "'," & iTran & "," & iRef & ")"
            RdoCon.Execute sSql, rdExecDirect
            .Col = 0
            Set .CellPicture = imgInc
         Else
            On Error GoTo 0
            .Col = 2
            sSql = "DELETE FROM ArecTable WHERE " _
                   & "RECITEM = '" & Trim(.Text) & "' AND " _
                   & "RECITEMTYPE = " & bType & " AND " _
                   & "RECACCOUNT = '" & sCurAcct & "'"
            If bType = 1 Then
               sSql = sSql & " AND RECCUST = '" & sCust & "'"
            End If
            If bType = 2 Then
               sSql = sSql & " AND RECTRAN = " & iGLTran(.Row) _
                      & " AND RECREF = " & iGLRef(.Row)
            End If
            RdoCon.Execute sSql, rdExecDirect
            .Col = 0
            Set .CellPicture = imgdInc
         End If
      End If
   End With
   Exit Sub
   DiaErr1:
   sProcName = "GridClick"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetAccount()
   Dim sTemp As String
   Dim RdoBal As rdoResultset
   On Error GoTo DiaErr1
   sTemp = Compress(cmbAct)
   If sTemp <> sCurAcct Then
      sCurAcct = sTemp
      lblDsc(0) = UpdateActDesc(cmbAct)
      ' Get last reconcialed balance
      sSql = "SELECT GLRECBAL FROM GlacTable WHERE GLACCTREF = '" _
             & sCurAcct & "'"
      bSqlRows = GetDataSet(RdoBal)
      If bSqlRows Then
         With RdoBal
            txtBeg = Format(.rdoColumns(0), CURRENCYMASK)
            .Cancel
         End With
      Else
         txtBeg = "0.00"
      End If
      Set RdoBal = Nothing
      FillGrid
      UpdateTotals
   End If
   Exit Sub
   DiaErr1:
   sProcName = "GetAccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Mark(bType As Byte)
   ' bType Doc
   ' True = Select All
   ' False = Remove All
   Dim i As Integer
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "DELETE FROM ArecTable WHERE RECITEMTYPE = " & SSTab1.Tab
   RdoCon.Execute sSql, rdExecDirect
   Select Case SSTab1.Tab
      Case 0
         With Grid1
            .Col = 0
            For i = 1 To .Rows - 1
               .Row = i
               Set .CellPicture = imgdInc
               If bType Then GridClick Grid1
            Next
         End With
      Case 1
         With Grid2
            For i = 1 To .Rows - 1
               .Row = i
               Set .CellPicture = imgdInc
               If bType Then GridClick Grid2
            Next
         End With
      Case 2
         With Grid3
            .Col = 0
            For i = 1 To .Rows - 1
               .Row = i
               Set .CellPicture = imgdInc
               If bType Then GridClick Grid3
            Next
         End With
   End Select
   UpdateTotals
   MouseCursor 0
   Exit Sub
   DiaErr1:
   sProcName = "Mark"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Highlight(pgrdMygrid As MSFlexGrid)
   Dim pintI As Integer
   Dim plngColorToUse As Long
   Dim plngNormal As Long
   Dim plngHighlight As Long
   Dim iTemp As Integer
   Dim iIndex As Integer
   Static pintLastRow(2) As Integer
   With pgrdMygrid
      iIndex = Right(.Name, 1) - 1
      plngNormal = .BackColor
      plngHighlight = .BackColorSel
      iTemp = .Row
      If pintLastRow(iIndex) > 0 Then
         .Row = pintLastRow(iIndex)
         For pintI = 0 To (.Cols - 1)
            .Col = pintI
            .CellBackColor = plngNormal
            .CellForeColor = .ForeColor
         Next
      End If
      .Row = iTemp
      For pintI = 0 To (.Cols - 1)
         .Col = pintI
         .CellBackColor = plngHighlight
         .CellForeColor = .ForeColorSel
      Next
      pintLastRow(iIndex) = .Row
      .Col = 0
   End With
End Sub
