VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaGLe10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Account Reconciliation"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   517
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCnl 
      Caption         =   "&Reselect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8640
      TabIndex        =   7
      ToolTipText     =   "Cancel Operation"
      Top             =   1200
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   8640
      TabIndex        =   4
      ToolTipText     =   "Display Transactions For Account"
      Top             =   480
      Width           =   875
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "4"
      ToolTipText     =   "Enter New Team Member  (15 Char) Or Select From List"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   1440
      Picture         =   "diaGLe10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Display Selected Transaction"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdMrk 
      Height          =   330
      Left            =   240
      Picture         =   "diaGLe10a.frx":017E
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Select All"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdUnm 
      Height          =   330
      Left            =   840
      Picture         =   "diaGLe10a.frx":03B0
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Deselect All"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Uncleared Cash Receipts (Sorted By Date \ Receipt Number)"
      Top             =   1620
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Checks and Payments"
      TabPicture(0)   =   "diaGLe10a.frx":0407
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Grid1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Deposits and Credits"
      TabPicture(1)   =   "diaGLe10a.frx":0423
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Grid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&GL Entries"
      TabPicture(2)   =   "diaGLe10a.frx":043F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid3"
      Tab(2).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid Grid3 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   39
         ToolTipText     =   "Uncleared GL Entries (Sorted By Post Date)"
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         _Version        =   393216
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid Grid2 
         Height          =   4095
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   12
         ToolTipText     =   "Uncleared Checks (Sorted By Check Number)"
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ComboBox cmbIntAct 
      Height          =   315
      Left            =   3420
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cmbSerAct 
      Height          =   315
      Left            =   3420
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtInt 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Tag             =   "1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtSer 
      Height          =   315
      Left            =   1500
      TabIndex        =   6
      Tag             =   "1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
   Begin VB.TextBox txtend 
      Height          =   285
      Left            =   6540
      TabIndex        =   3
      Tag             =   "2"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtBeg 
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Tag             =   "1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   8640
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Save And Exit"
      Top             =   0
      Width           =   875
   End
   Begin VB.CommandButton cmdRec 
      Caption         =   "Re&concile "
      Enabled         =   0   'False
      Height          =   315
      Left            =   8640
      TabIndex        =   5
      ToolTipText     =   "Reconcile Marked Transactions"
      Top             =   840
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   14
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
      PictureUp       =   "diaGLe10a.frx":045B
      PictureDn       =   "diaGLe10a.frx":05A1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   6840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7755
      FormDesignWidth =   9600
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "GL Transactions"
      Height          =   255
      Index           =   8
      Left            =   3780
      TabIndex        =   51
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label lblTotalGL 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   3300
      TabIndex        =   50
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plus/Minus GL"
      Height          =   240
      Index           =   7
      Left            =   6000
      TabIndex        =   49
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblGl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   7800
      TabIndex        =   48
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plus Deposits"
      Height          =   240
      Index           =   5
      Left            =   6000
      TabIndex        =   47
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label lblDeposits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   7800
      TabIndex        =   46
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      Height          =   240
      Index           =   4
      Left            =   6000
      TabIndex        =   45
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblStartBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   7800
      TabIndex        =   44
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblLst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7320
      TabIndex        =   43
      Top             =   60
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Reconcilied"
      Height          =   255
      Index           =   3
      Left            =   6060
      TabIndex        =   42
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Date"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   41
      Top             =   540
      Width           =   1155
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4980
      TabIndex        =   36
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4980
      TabIndex        =   35
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   3180
      TabIndex        =   34
      Top             =   60
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   255
      Index           =   18
      Left            =   2700
      TabIndex        =   33
      Top             =   1260
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   255
      Index           =   17
      Left            =   2700
      TabIndex        =   32
      Top             =   900
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Earned"
      Height          =   255
      Index           =   14
      Left            =   300
      TabIndex        =   31
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charge"
      Height          =   255
      Index           =   13
      Left            =   300
      TabIndex        =   30
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marked as Cleared:"
      Height          =   255
      Index           =   24
      Left            =   3300
      TabIndex        =   29
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label lblDiffBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   7800
      TabIndex        =   28
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label lblPayments 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   7800
      TabIndex        =   27
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblEndBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   7800
      TabIndex        =   26
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label lblTotalChk 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   3300
      TabIndex        =   25
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label lblTotalDep 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   3300
      TabIndex        =   24
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Checks and Payments"
      Height          =   255
      Index           =   23
      Left            =   3780
      TabIndex        =   23
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposits and Credits"
      Height          =   255
      Index           =   22
      Left            =   3780
      TabIndex        =   22
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Difference"
      Height          =   240
      Index           =   21
      Left            =   6000
      TabIndex        =   21
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Minus Payments"
      Height          =   240
      Index           =   20
      Left            =   6000
      TabIndex        =   20
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Balance"
      Height          =   240
      Index           =   19
      Left            =   6000
      TabIndex        =   19
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Account"
      Height          =   255
      Index           =   6
      Left            =   420
      TabIndex        =   18
      Top             =   120
      Width           =   1155
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   2040
      Picture         =   "diaGLe10a.frx":06E7
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   2760
      Picture         =   "diaGLe10a.frx":073E
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Balance"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   17
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      Height          =   255
      Index           =   2
      Left            =   2700
      TabIndex        =   16
      Top             =   540
      Width           =   1335
   End
End
Attribute VB_Name = "diaGLe10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaGLe10a - Account Reconcilation
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
' 10/05/04 (nth) Moved to General Ledger.
' 12/01/04 (nth) Added statement date.
'
'*************************************************************************************

Option Explicit

' Numeric mask used in this transaction formats up to 1 billion $.
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

Private Sub Reconcile()
   Dim i As Integer
   Dim iDep As Integer
   Dim iChk As Integer
   Dim sJournalID1 As String 'PJ
   Dim sJournalID2 As String 'CR
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim sVendor As String
   'Dim rdoAct As ADODB.Recordset
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
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   Dim sToday As String
   sToday = Format(ES_SYSDATE, "mm/dd/yy")
   
   If CCur(txtSer) > 0 Then
      sGL = "SC" & Format(txtDte, "yyyymmdd")
      sPst = GetFYPeriodEnd(CDate(txtDte))
      
      sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJOPEN,GJPOST) " _
             & "VALUES('" & sGL & "','SERVICE CHARGE','" _
             & txtDte & "','" & sPst & "')"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIACCOUNT,JICRD,JIDATE,JICLEAR) " _
             & "VALUES('" & sGL & "',1,1,'" & Compress(cmbAct) & "'," _
             & CCur(txtSer) & ",'" & sToday & "','" & sToday & "')"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIACCOUNT,JIDEB,JIDATE,JICLEAR) " _
             & "VALUES('" & sGL & "',1,2,'" & Compress(cmbSerAct) & "'," _
             & CCur(txtSer) & ",'" & sToday & "','" & sToday & "')"
      clsADOCon.ExecuteSQL sSql
   End If
   
   If CCur(txtInt) > 0 Then
      sGL = "IE" & Format(txtDte, "yyyymmdd")
      sPst = GetFYPeriodEnd(txtDte)
      
      sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJOPEN,GJPOST) " _
             & "VALUES('" & sGL & "','INTEREST EARNED','" _
             & txtDte & "','" & sPst & "')"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIACCOUNT,JIDEB,JIDATE,JICLEAR) " _
             & "VALUES('" & sGL & "',1,1,'" & Compress(cmbAct) & "'," _
             & CCur(txtInt) & ",'" & sToday & "','" & sToday & "')"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIACCOUNT,JICRD,JIDATE,JICLEAR) " _
             & "VALUES('" & sGL & "',1,2,'" & Compress(cmbIntAct) & "'," _
             & CCur(txtInt) & ",'" & sToday & "','" & sToday & "')"
      clsADOCon.ExecuteSQL sSql
   End If
   
   ' checks
   With Grid1
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         If .CellPicture = imgInc Then
            .Col = 2
            sSql = "UPDATE ChksTable Set CHKCLEARDATE='" & txtDte _
                   & "' WHERE CHKNUMBER='" & Trim(.Text) & "'"
            clsADOCon.ExecuteSQL sSql
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
            sSql = "UPDATE CashTable Set CACLEAR='" & txtDte _
                   & "' WHERE CACUST='" & sTemp _
                   & "' AND CACHECKNO='" & Trim(.Text) & "'"
            clsADOCon.ExecuteSQL sSql
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
            sSql = "UPDATE GjitTable SET JICLEAR = '" & txtDte _
                   & "' WHERE JINAME = '" & Trim(.Text) & "' AND JITRAN = " _
                   & iGLTran(i) & " AND JIREF = " & iGLRef(i)
            clsADOCon.ExecuteSQL sSql
            .Col = 0
         End If
      Next
   End With
   
   ' Record ending balance, date , and who did it in account record
   sSql = "UPDATE GlacTable SET GLRECDATE = '" & txtDte & "',GLRECBAL = " _
          & CCur(txtend) & ",GLRECBY = '" & sInitials & "'" _
          & " WHERE GLACCTREF = '" & sCurAcct & "'"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DELETE FROM ArecTable"
   clsADOCon.ExecuteSQL sSql
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "Account Successfully Reconciled.", True
      txtBeg = txtend
      lblLst = txtDte
      txtend = "0.00"
      txtSer = "0.00"
      txtInt = "0.00"
      DoEvents
      cmdCnl_Click
      SSTab1.enabled = True
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Could Not Reconcile Account.", _
         vbExclamation, Caption
   End If
   
   MouseCursor 0
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

Private Sub cmbAct_GotFocus()
   SelectFormat Me
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
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdCnl_Click()
   cmbAct.enabled = True
   txtDte.enabled = True
   cmdRec.enabled = False
   cmdSel.enabled = True
   Grid1.Clear
   Grid2.Clear: Grid2.Refresh
   Grid3.Clear
   SSTab1.enabled = False
   
   cmbAct.SetFocus
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
   SaveSetting "Esi2000", "Reconcile", "ServiceChargeAcct", cmbSerAct
   SaveSetting "Esi2000", "Reconcile", "InterestAcct", cmbIntAct
   Reconcile
   FillGrid
   UpdateTotals
End Sub

Private Sub cmdSel_Click()
   MouseCursor 13
   SSTab1.enabled = True
   FillGrid
   UpdateTotals
   cmbAct.enabled = False
   txtDte.enabled = True 'False
   cmdSel.enabled = False
   cmdCnl.enabled = True
   MouseCursor 0
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
   FormLoad Me
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
   
   SSTab1.enabled = False
   optDis.enabled = False
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   txtSer = "0.00"
   txtInt = "0.00"
   txtBeg = "0.00"
   txtend = "0.00"
   sCurAcct = ""
   bOnLoad = True
   
   Dim gl As New GLTransaction
   gl.FillComboWithAccounts cmbSerAct
   gl.FillComboWithAccounts cmbIntAct
   
   cmbSerAct = GetSetting("Esi2000", "Reconcile", "ServiceChargeAcct", "")
   cmbIntAct = GetSetting("Esi2000", "Reconcile", "InterestAcct", "")
   lblDsc(1) = UpdateActDesc(cmbSerAct)
   lblDsc(2) = UpdateActDesc(cmbIntAct)
   
   SSTab1.Tab = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaGLe10a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillGrid()
   Dim RdoChk As ADODB.Recordset
   Dim rdoDep As ADODB.Recordset
   Dim rdoGL As ADODB.Recordset
   'Dim rdoFee  As ADODB.RecordSet
   Dim sItem As String
   Dim i As Integer
   Dim sDate As String
   
   On Error GoTo DiaErr1
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
          & "and gjpost <= '" & txtDte & "' and gjtemplate=0  ORDER BY GJPOST"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoGL)
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
   'SELECT DISTINCT CARCDATE,CACKAMT,CACHECKNO,CUNICKNAME,RECITEM,CACUST,
   '(SELECT SUM(DCDEBIT) FROM JritTable
   'WHERE DCCHECKNO = CACHECKNO AND DCCUST = CACUST AND DCDESC = 'TRAN FEE') AS FEE
   'FROM CashTable JOIN CustTable ON CACUST = CUREF
   'LEFT OUTER JOIN ArecTable ON CACUST = RECCUST AND CACHECKNO = RECITEM
   'WHERE (CACASHACCT = '1004') AND (CACLEAR IS NULL)
   'AND (CACANCELED = 0) and carcdate <= '05/10/05'
   'ORDER BY CARCDATE
   
   '    sSql = "SELECT DISTINCT CARCDATE,CACKAMT,CACHECKNO,CUNICKNAME,RECITEM,CACUST " _
   '        & "FROM CashTable INNER JOIN CustTable ON CACUST = CUREF LEFT OUTER JOIN " _
   '        & "ArecTable ON CACUST = RECCUST AND CACHECKNO = RECITEM " _
   '        & "WHERE (CACASHACCT = '" & sCurAcct & "') AND (CACLEAR IS NULL) AND (CACANCELED = 0) " _
   '        & "and carcdate <= '" & txtDte & "' ORDER BY CARCDATE"
   
   sSql = "SELECT DISTINCT CARCDATE,CACKAMT,CACHECKNO,CUNICKNAME,RECITEM,CACUST, " & vbCrLf _
          & "(SELECT ISNULL(SUM(DCDEBIT),0) FROM JritTable " & vbCrLf _
          & "WHERE DCCHECKNO = CACHECKNO AND DCCUST = CACUST AND DCDESC = 'TRAN FEE') AS FEE " & vbCrLf _
          & "FROM CashTable JOIN CustTable ON CACUST = CUREF " & vbCrLf _
          & "LEFT OUTER JOIN ArecTable ON CACUST = RECCUST AND CACHECKNO = RECITEM " & vbCrLf _
          & "WHERE (CACASHACCT = '" & sCurAcct & "') AND (CACLEAR IS NULL) AND (CACANCELED = 0) " & vbCrLf _
          & "AND CARCDATE <= '" & txtDte & "' " & vbCrLf _
          & "ORDER BY CARCDATE"
   
   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoDep, ES_FORWARD)
   i = 1
   
   'show alternate dates in alternate colors
   Dim lColor1 As Long
   Dim lColor2 As Long
   Dim lColor As Long
   Dim sPriorDate As String
   lColor1 = RGB(255, 255, 255) 'white
   'lColor2 = RGB(224, 224, 224)     'light grey
   lColor2 = &HFFFFC0 'light blue
   lColor = lColor2
   sPriorDate = "nomatch"
   
   If bSqlRows Then
      With rdoDep
         While Not .EOF
            
            '                ' trans fee
            '                sSql = "SELECT SUM(DCDEBIT) FROM JritTable WHERE DCCHECKNO = '" _
            '                    & Trim(!CACHECKNO) & "' AND DCCUST = '" & Trim(!CACUST) _
            '                    & "' AND DCDESC = 'TRANS FEE'"
            '                bSqlRows = clsAdoCon.GetDataSet(sSql,rdoFee)
            '                Debug.Print sSql
            '                Set rdoFee = Nothing
            '                sItem = Chr(9) & Format(!CARCDATE, "mm/dd/yy") _
            '                    & Chr(9) & " " & !CACHECKNO _
            '                    & Chr(9) & " " & !CUNICKNAME _
            '                    & Chr(9) & Format(rdoFee.Fields(0), CURRENCYMASK) _
            '                    & Chr(9) & Format(.Fields(1), CURRENCYMASK)
            sItem = Chr(9) & Format(!CARCDATE, "mm/dd/yy") _
                    & Chr(9) & " " & !CACHECKNO _
                    & Chr(9) & " " & !CUNICKNAME _
                    & Chr(9) & Format(!FEE, CURRENCYMASK) _
                    & Chr(9) & Format(!CACKAMT, CURRENCYMASK)
            Grid2.AddItem sItem
            Grid2.Row = i
            Grid2.Col = 0
            Grid2.CellPictureAlignment = flexAlignCenterCenter
            
            If IsNull(!RECITEM) Then
               Set Grid2.CellPicture = imgdInc
            Else
               Set Grid2.CellPicture = imgInc
            End If
            
            'swap colors if date change
            If sPriorDate <> CStr(!CARCDATE) Then
               sPriorDate = CStr(!CARCDATE)
               If lColor = lColor1 Then
                  lColor = lColor2
               Else
                  lColor = lColor1
               End If
            End If
            Dim j As Integer
            For j = 0 To 5
               Grid2.Col = j
               Grid2.CellBackColor = lColor
            Next
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
          & "CHKNUMBER = RECITEM AND RECCUST = CHKVENDOR" & vbCrLf _
          & "WHERE (CHKCLEARDATE IS NULL) AND (CHKPRINTED = 1) AND ISNUMERIC(CHKNUMBER) = 1 " _
          & "AND (CHKVOID = 0) AND (CHKACCT = '" & sCurAcct & "') and chkpostdate  <='" & txtDte _
          & "' OR (CHKCLEARDATE IS NULL) " _
          & "AND (CHKACCT = '" & sCurAcct & "') AND " _
          & "ISNUMERIC(CHKNUMBER) = 1 and chkvoid = 0 and chkpostdate <= '" & txtDte & "' " _
          & "order by CAST(CHKNUMBER AS decimal)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      FillCheckGrid RdoChk
   End If
   Set RdoChk = Nothing
   
   ' Alpha/Numeric check numbers
   sSql = "SELECT CHKNUMBER,CHKAMOUNT,CHKMEMO,CHKPRINTDATE,CHKPOSTDATE,VENICKNAME," _
          & "RECITEM FROM VndrTable INNER JOIN ChksTable ON VndrTable.VEREF = " _
          & "CHKVENDOR LEFT OUTER JOIN ArecTable ON CHKACCT = RECACCOUNT AND " _
          & "CHKNUMBER = RECITEM AND RECCUST = CHKVENDOR" & vbCrLf _
          & "WHERE (CHKCLEARDATE IS NULL) AND (CHKPRINTED = 1) " _
          & "AND (CHKVOID = 0) AND (CHKACCT = '" & sCurAcct & "') and ISNUMERIC(CHKNUMBER) = 0 " _
          & "and chkpostdate  <='" & txtDte & "' OR (CHKCLEARDATE IS NULL) " _
          & "AND (CHKACCT = '" & sCurAcct & "') AND ISNUMERIC(CHKNUMBER) = 0 " _
          & "and chkvoid = 0 and chkpostdate  <='" & txtDte & "' order by chknumber "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      FillCheckGrid RdoChk
   End If
   Set RdoChk = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCheckGrid(RdoChk As ADODB.Recordset)
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
                 & Chr(9) & " " & Trim(!CHKNUMBER) _
                 & Chr(9) & " " & Trim(!VENICKNAME) _
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
   If Grid1.MouseCol = 0 Then
      GridClick Grid1, False
      Highlight Grid1
      UpdateTotals
   End If
End Sub

Private Sub FillAccounts()
   ' Fill account combo
   ' Need to add account descriptions
   Dim rdoAct As ADODB.Recordset
   Dim b As Byte
   On Error GoTo DiaErr1
   
   sProcName = "getcashacco"
   b = GetCashAccounts()
   If b = 3 Then
      MouseCursor 0
      MsgBox "One Or More Cash Accounts Are Not Active." & vbCr _
         & "Please Set All Cash Accounts In The " & vbCr _
         & "System Setup, Administration Section.", _
         vbInformation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   
   ' Accounts
   sProcName = "fillacounts"
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
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
   On Error Resume Next
   cmbAct.SetFocus
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Local Errors - from FillCombo
'Flash with 1 = no accounts nec, 2 = OK, 3 = Not enough accounts

Public Function GetCashAccounts() As Byte
   Dim rdoCsh As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   sSql = "SELECT COGLVERIFY,COCRCASHACCT,COCRDISCACCT,COSJARACCT," _
          & "COCRCOMMACCT,COCRREVACCT,COCREXPACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCsh, ES_FORWARD)
   sProcName = "getcashacct"
   If bSqlRows Then
      With rdoCsh
         For i = 1 To 6
            If "" & Trim(.Fields(i)) = "" Then
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
   Set rdoCsh = Nothing
End Function

Private Sub UpdateTotals()
   Dim i As Integer
   Dim iTotalDep As Integer
   Dim cTotalDep As Currency
   
   Dim iTotalFee As Integer
   Dim cTotalFee As Currency
   
   Dim iTotalChk As Integer
   Dim cTotalChk As Currency
   Dim iTotalGL As Integer
   Dim cTotalGl As Currency
   Dim cBalance As Currency
   Dim cSnI As Currency
   Dim iTemp As Integer
   
   ' sum checks
   On Error GoTo DiaErr1
   With Grid1
      iTemp = .Row
      .Col = 0
      For i = 1 To .Rows - 1
         .Row = i
         If .CellPicture = imgInc Then
            .Col = 5
            cBalance = cBalance - CCur(.Text)
            cTotalChk = cTotalChk + CCur(.Text)
            iTotalChk = iTotalChk + 1
            Debug.Print "Payment " & iTotalChk & " + " & .Text & " = " & cBalance
            .Col = 1
            Debug.Print "Date " & .Text
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
            cBalance = cBalance + CCur(.Text)
            iTotalDep = iTotalDep + 1
            cTotalDep = cTotalDep + CCur(.Text)
            Debug.Print "Deposit " & iTotalDep & " + " & .Text & " = " & cBalance
            
            .Col = 4
            iTotalFee = iTotalFee + 1
            cTotalFee = cTotalFee + CCur(.Text)
            Debug.Print "Fee " & iTotalFee & " + " & .Text & " = " & cTotalFee
            
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
               cTotalGl = cTotalGl + .Text
            Else
               .Col = 5
               cBalance = cBalance - .Text
               cTotalGl = cTotalGl - .Text
            End If
            iTotalGL = iTotalGL + 1
            .Col = 0
         End If
      Next
      .Row = iTemp
   End With
   
   '    ' Update totals
   '    cSnI = CCur(txtInt) - CCur(txtSer)
   '
   '    lblEndBal = Format(txtend, CURRENCYMASK)
   '    lblClearBal = Format(cBalance + cSnI, CURRENCYMASK)
   '    lblDiffBal = Format((cBalance + cSnI) + CCur(txtBeg & "0") - _
   '        CCur(txtend & "0"), CURRENCYMASK)
   '
   '    lblTotalChk = iTotalChk
   '    lblTotalDep = iTotalDep
   '
   
   'add service charge and interest
   cTotalDep = cTotalDep + CCur(txtInt)
   cTotalChk = cTotalChk + CCur(txtSer)
   
   ' Update totals
   'cSnI = CCur(txtInt) - CCur(txtSer)
   
   If txtBeg = "" Then
      lblStartBal = "0"
   Else
      lblStartBal = Format(txtBeg, CURRENCYMASK)
   End If
   lblDeposits = Format(cTotalDep, CURRENCYMASK)
   lblPayments = Format(-(cTotalChk + cTotalFee), CURRENCYMASK)
   lblGl = Format(cTotalGl, CURRENCYMASK)
   lblEndBal = Format(CCur(lblStartBal) + CCur(lblDeposits) + CCur(lblPayments) _
               + CCur(lblGl), CURRENCYMASK)
   lblDiffBal = Format(CCur(txtend) - CCur(lblEndBal), CURRENCYMASK)
   
   lblTotalChk = iTotalChk
   lblTotalDep = iTotalDep
   lblTotalGL = iTotalGL
   
   
   ' Highlight reconcile button if differnce = 0
   With cmdRec
      If Val(lblDiffBal) = 0 Then
         .enabled = True
      Else
         .enabled = False
      End If
   End With
   Exit Sub
   
DiaErr1:
   sProcName = "updatetotals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Or KeyAscii = 13 Then
      Grid1_Click
   End If
End Sub

Private Sub Grid2_Click()
   If Grid2.MouseCol = 0 Then
      GridClick Grid2, True
      UpdateTotals
   End If
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Or KeyAscii = 13 Then
      Grid2_Click
   End If
End Sub

Private Sub Grid2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   'sum total cash receipts for day
   
   DoEvents
   With Grid2
      If .MouseCol <> 1 Or .MouseRow = 0 Then
         .ToolTipText = ""
         Exit Sub
      End If
      
      '.ToolTipText = "Column " & .MouseCol & " Row = " & .MouseRow
      .Row = .MouseRow
      .Col = 1
      Dim dt As String
      dt = .Text
      
      'sum the amounts of all columns for this date
      Dim i As Integer
      Dim total As Currency
      For i = 0 To .Rows - 1
         .Row = i
         .Col = 1
         If .Text = dt Then
            .Col = 4
            total = total + CCur("0" + .Text) 'fee
            .Col = 5
            total = total + CCur("0" + .Text) 'amount
         End If
      Next
      .ToolTipText = dt & " total: " & Format(total, "0.00")
      DoEvents
   End With
End Sub

Private Sub Grid3_Click()
   If Grid3.MouseCol = 0 Then
      GridClick Grid3, False
      Highlight Grid3
      UpdateTotals
   End If
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
         Grid1.SetFocus
         optDis.enabled = False
      Case 1
         Grid2.SetFocus
         optDis.enabled = True
      Case 2
         Grid3.SetFocus
         optDis.enabled = True
   End Select
End Sub

Private Sub txtBeg_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = Format(txtBeg, CURRENCYMASK)
   If Trim(txtBeg) = "" Then txtBeg = "0.00"
   cOpenBal = CCur(txtBeg & "0")
   UpdateTotals
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub

Private Sub txtEnd_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtEnd_LostFocus()
   On Error Resume Next
   txtend = Format(txtend, CURRENCYMASK)
   If Trim(txtend) = "" Then txtend = "0.00"
   cEndBal = CCur(txtend)
   If Err Then
      MsgBox "Invalid ending amount: " & txtend
      txtend.SetFocus
      Exit Sub
   End If
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

Private Sub txtSer_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtSer_LostFocus()
   txtSer = CheckLen(txtSer, 11)
   txtSer = Format(txtSer, CURRENCYMASK)
   If Trim(txtSer) = "" Then txtSer = "0.00"
   UpdateTotals
End Sub


Private Sub GridClick(Whichgrid As MSFlexGrid, CheckAllRowsForDate As Boolean)
   ' bType doc
   ' 0 = Checks
   ' 1 = Cash Receipt
   ' 2 = GL
   ' if CheckAllRowsForDate and bType = 1 then check or uncheck all items for the given date
   
   Dim bType As Byte
   Dim sCust As String
   Dim iTran As Integer
   Dim iRef As Integer
   On Error GoTo DiaErr1
   bType = CByte(Right(Whichgrid.Name, 1)) - 1
   With Whichgrid
      'If .Rows > 0 And .Row > 0 And .MouseCol = 0 Then   'mousecol checked in callers
      If .Rows > 0 And .Row > 0 Then
         '            If bType = 1 Then
         '                .Col = 3
         '                sCust = Compress(.Text)
         '            End If
         .Col = 0
         If .CellPicture = imgdInc Then
            
            'if no cr items previously selected for this date,
            'check all rows for this date
            Dim nMinRow As Integer
            Dim nMaxRow As Integer
            Dim nStartRow As Integer
            nMinRow = -1
            nMaxRow = -1
            nStartRow = .Row
            
            Dim j As Integer
            If bType = 1 And CheckAllRowsForDate Then
               Dim dt As String
               .Col = 1
               dt = .Text
               For j = 1 To .Rows - 1
                  .Row = j
                  
                  .Col = 1
                  If .Text = dt Then
                     
                     'if rows previously selected for this date, just select the one row
                     .Col = 0
                     If .CellPicture = imgInc Then
                        nMinRow = nStartRow
                        nMaxRow = nStartRow
                        Exit For
                     End If
                     If nMinRow = -1 Then
                        nMinRow = .Row
                     End If
                     nMaxRow = .Row
                  End If
               Next
            Else
               nMinRow = .Row
               nMaxRow = .Row
            End If
            
            'check all selected rows
            For j = nMinRow To nMaxRow
               .Row = j
               'for check (0) get vendor, for deposit (1) get customer
               If bType = 1 Or bType = 0 Then
                  .Col = 3
                  sCust = Compress(.Text)
               ElseIf bType = 2 Then
                  iTran = iGLTran(.Row)
                  iRef = iGLRef(.Row)
               End If
               .Col = 2
               Debug.Print "row " & j & " " & .Text & " " & sCurAcct & " " & sCust & " " & iTran & "-" & iRef
               sSql = "INSERT INTO ArecTable(RECITEM,RECITEMTYPE," _
                      & "RECACCOUNT,RECBY,RECCUST,RECTRAN,RECREF) VALUES ('" _
                      & Trim(.Text) & "'," & bType _
                      & ",'" & sCurAcct & "','" & sInitials _
                      & "','" & sCust & "'," & iTran & "," & iRef & ")"
               clsADOCon.ExecuteSQL sSql
               .Col = 0
               Set .CellPicture = imgInc
            Next
            .Row = nStartRow
         Else
            .Col = 2
            sSql = "DELETE FROM ArecTable WHERE " _
                   & "RECITEM = '" & Trim(.Text) & "' AND " _
                   & "RECITEMTYPE = " & bType & " AND " _
                   & "RECACCOUNT = '" & sCurAcct & "'"
            If bType = 1 Then
               .Col = 3
               sCust = Compress(.Text)
               sSql = sSql & " AND RECCUST = '" & sCust & "'"
            ElseIf bType = 2 Then
               sSql = sSql & " AND RECTRAN = " & iGLTran(.Row) _
                      & " AND RECREF = " & iGLRef(.Row)
            End If
            Debug.Print sSql
            clsADOCon.ExecuteSQL sSql
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
   Dim RdoBal As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sTemp = Compress(cmbAct)
   If sTemp <> sCurAcct Then
      sCurAcct = sTemp
      lblDsc(0) = UpdateActDesc(cmbAct)
      ' Get last reconciled balance
      sSql = "SELECT GLRECBAL,GLRECDATE FROM GlacTable WHERE GLACCTREF = '" _
             & sCurAcct & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBal)
      If bSqlRows Then
         With RdoBal
            txtBeg = Format(.Fields(0), CURRENCYMASK)
            lblLst = Format(.Fields(1), "mm/dd/yy")
            .Cancel
         End With
      Else
         txtBeg = "0.00"
      End If
      Set RdoBal = Nothing
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
   clsADOCon.ExecuteSQL sSql
   Select Case SSTab1.Tab
      Case 0
         With Grid1
            .Col = 0
            For i = 1 To .Rows - 1
               .Row = i
               Set .CellPicture = imgdInc
               If bType Then GridClick Grid1, False
            Next
         End With
      Case 1
         With Grid2
            .Col = 0
            For i = 1 To .Rows - 1
               .Row = i
               Set .CellPicture = imgdInc
               If bType Then GridClick Grid2, False
            Next
         End With
      Case 2
         With Grid3
            .Col = 0
            For i = 1 To .Rows - 1
               .Row = i
               Set .CellPicture = imgdInc
               If bType Then GridClick Grid3, False
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
   On Error Resume Next
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
