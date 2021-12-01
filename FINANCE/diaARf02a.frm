VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form diaARf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Cash Receipt"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -6600
      TabIndex        =   23
      Top             =   -960
      Width           =   6015
   End
   Begin ResizeLibCtl.ReSize ReSize2 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5535
      FormDesignWidth =   6135
   End
   Begin VB.ComboBox cmbChk 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Tag             =   "1"
      Text            =   "cmbChk"
      ToolTipText     =   "Checks Available (Not Posted)"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "C&ancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   "Cancel Selected Cash Receipt"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "cmbCst"
      ToolTipText     =   "Contains Customers With Cash Receipts "
      Top             =   360
      Width           =   1440
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaARf02a.frx":0000
      PictureDn       =   "diaARf02a.frx":0146
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Click On Check Number To Select A Check"
      Top             =   3120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   285
      Index           =   8
      Left            =   2280
      TabIndex        =   22
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   21
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblClo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   20
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblAct 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   19
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Closed"
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   17
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label lblJrn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal"
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   15
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label lblNot 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Applied"
      Height          =   405
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   12
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   11
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1065
   End
End
Attribute VB_Name = "diaARf02a"
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
' diaARf02a - Cancel A Cash Receipt
'
' Created: (cjs)
' Revisons:
' 05/01/03 (nth) Fixed "subscript out of range" error per JLH and JEVCO.
' 02/04/04 (nth) Added check amount.
' 07/07/04 (nth) Show detail for cash receipt.
' 07/30/04 (nth) To show advance payment invoices and correctly cancel.
' 10/04/04 (nth) Do not allow canceling of CA invoices if applied to payment.
'
'*************************************************************************************

Option Explicit


Private Const CB_FINDSTRING = &H14C
Private Const CB_SHOWDROPDOWN = &H14F
Private Const LB_FINDSTRING = &H18F
Private Const CB_ERR = (-1)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
  hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  
Dim bFirstTime As Byte
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim sCust As String
Dim sMsg As String
Dim sCheck As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbChk_Click()
   If sCheck <> cmbChk Then
      GetThisCheck
   End If
End Sub

Private Sub cmbChk_GotFocus()
    SendMessage cmbChk.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End Sub

Private Sub ComboBoxKeyPress(cmb As ComboBox, KeyAscii As Integer)
    Dim cb As Long
    Dim FindString As String
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    If cmb.SelLength = 0 Then
        FindString = cmb.Text & Chr$(KeyAscii)
    Else
        FindString = Left(cmb.Text, cmb.SelStart) & Chr$(KeyAscii)
    End If
    SendMessage cmb.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&
    cb = SendMessage(cmb.hWnd, CB_FINDSTRING, 1, ByVal FindString)
    If cb <> CB_ERR Then
        cmb.ListIndex = cb
        cmb.SelStart = Len(FindString)
        cmb.SelLength = Len(cmb.Text) - cmb.SelStart
    End If
    KeyAscii = 0
End Sub

Private Sub cmbChk_KeyPress(KeyAscii As Integer)
    ComboBoxKeyPress cmbChk, KeyAscii
    If KeyAscii >= 32 And KeyAscii <= 127 Then KeyAscii = 0
   
End Sub

Private Sub cmbChk_LostFocus()
   If Not bCancel Then
      If sCheck <> cmbChk Then
         GetThisCheck
      End If
   End If
End Sub

Private Sub cmbCst_Click()
    If bOnLoad Then Exit Sub

   FindCustomer Me, cmbCst
   sCust = Compress(cmbCst)
   GetChecks
End Sub

Private Sub cmbCst_GotFocus()
    If Not bFirstTime Then SendMessage cmbCst.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&
    bFirstTime = False
End Sub

Private Sub cmbCst_KeyPress(KeyAscii As Integer)
    ComboBoxKeyPress cmbCst, KeyAscii
    If KeyAscii >= 32 And KeyAscii <= 127 Then KeyAscii = 0
End Sub

Private Sub cmbCst_LostFocus()
   If Not bCancel Then
      cmbCst = CheckLen(cmbCst, 10)
      If Len(cmbCst) Then
         sCust = Compress(cmbCst)
         
         FindCustomer Me, cmbCst
         GetChecks
      Else
         sCust = ""
         lblNme = ""
         cmbChk.Clear
         cmbChk = ""
      End If
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdDel_Click()
   CancelReceipt
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Cancel a Cash Receipt"
      cmdHlp = False
      MouseCursor 0
   End If
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
   bFirstTime = True
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   IniGrid
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaARf02a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   Dim rdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sProcName = "fillcombo"
   cmbChk.Clear
   cmbCst.Clear
   lblNme = ""
   'cmbChk = ""
   sSql = "SELECT DISTINCT CACUST,CUREF,CUNICKNAME " _
          & "FROM CashTable,CustTable WHERE CACUST=CUREF AND CACANCELED=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         Do Until .EOF
            AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set rdoCst = Nothing
   If cmbCst.ListCount > 0 Then
      If sCust <> "" Then
         cmbCst = sCust
      Else
         cmbCst = cmbCst.List(0)
      End If
      sCust = Compress(cmbCst)
      FindCustomer Me, cmbCst
      GetChecks
   End If
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetChecks()
   Dim RdoChk As ADODB.Recordset
   Dim sCust As String
   cmbChk.Clear
   On Error GoTo DiaErr1
   sCust = Compress(cmbCst)
   sSql = "SELECT DISTINCT CACUST,CACHECKNO " _
          & "FROM CashTable WHERE CACUST = '" & sCust & "' order by CACHECKNO "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         Do Until .EOF
            AddComboStr cmbChk.hWnd, "" & Trim(!CACHECKNO)
            .MoveNext
         Loop
      End With
   End If
   Set RdoChk = Nothing
   If cmbChk.ListCount > 0 Then
      cmbChk = cmbChk.List(0)
      GetThisCheck
   End If
   Exit Sub
DiaErr1:
   sProcName = "getchecks"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetThisCheck()
   Dim RdoChk As ADODB.Recordset
   Dim RdoInv As ADODB.Recordset
   Dim rdoNot As ADODB.Recordset
   Dim sItem As String
   Dim sCheck As String
   Dim cAmount As Currency
   
   On Error GoTo DiaErr1
   IniGrid
   lblNot = ""
   lblAct = ""
   lblDsc = ""
   sCheck = Trim(cmbChk)
   
   sSql = "SELECT DISTINCT DCHEAD,MJCLOSED,CARCDATE,CACKAMT,CATYPE " _
          & "FROM JritTable INNER JOIN CashTable ON DCCHECKNO = CACHECKNO " _
          & "AND DCCUST = CACUST INNER JOIN JrhdTable ON DCHEAD = MJGLJRNL " _
          & "WHERE CashTable.CACHECKNO = '" & sCheck & "' AND CACUST = '" _
          & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      ' Check
      With RdoChk
         lblAmt = "" & Format(!CACKAMT, CURRENCYMASK)
         lblDte = "" & Format(!CARCDATE, "mm/dd/yy")
         lblJrn = "" & Trim(!DCHEAD)
         lblClo = "" & Format(!MJCLOSED, "mm/dd/yy")
         lblTyp = GetCrType(!CATYPE)
         If lblClo = "" Then
            cmdDel.enabled = True
         Else
            cmdDel.enabled = False
         End If
         .Cancel
      End With
      
      ' Invoices
      sSql = "SELECT DISTINCT INVNO,INVPRE,INVDATE,SUM(DCDEBIT) AS Amount," _
             & "INVTYPE FROM JritTable INNER JOIN CihdTable ON DCINVNO = INVNO " _
             & "WHERE DCCHECKNO = '" & sCheck & "' AND DCCUST = '" & sCust & "' " _
             & "GROUP BY INVNO,INVPRE,INVDATE,INVTYPE "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
      If bSqlRows Then
         With RdoInv
            While Not .EOF
               cAmount = !Amount
               If !INVTYPE = "CM" Or !INVTYPE = "CA" Then
                  cAmount = cAmount * -1
               End If
               sItem = " " & Trim(!INVPRE) & Format(!InvNo, "000000") _
                       & vbTab & " " & !INVTYPE & vbTab & " " _
                       & Format(!INVDATE, "mm/dd/yy") & vbTab _
                       & Format(cAmount, CURRENCYMASK)
               Grid1.AddItem sItem
               .MoveNext
            Wend
            .Cancel
         End With
      End If
      Set RdoInv = Nothing
      
      ' Not Applied
      sSql = "SELECT DISTINCT DCDEBIT,GLACCTNO,GLDESCR FROM JritTable INNER JOIN " _
             & "CashTable ON DCCHECKNO = CACHECKNO AND DCCUST = CACUST INNER JOIN " _
             & "GlacTable ON DCACCTNO = GLACCTNO WHERE DCCREDIT = 0 AND CAINVNO = 0 " _
             & "AND CACHECKNO = '" & sCheck & "' AND CACUST = '" & sCust & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoNot)
      If bSqlRows Then
         With rdoNot
            lblNot = Format(.Fields(0), CURRENCYMASK)
            lblAct = "" & Trim(.Fields(1))
            lblDsc = "" & Trim(.Fields(2))
            .Cancel
         End With
      End If
      Set rdoNot = Nothing
   End If
   Set RdoChk = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getthisch"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub lblDte_Change()
   If lblDte = "None" Then
      lblDte.ForeColor = ES_RED
   Else
      lblDte.ForeColor = vbBlack
   End If
End Sub

Private Sub CancelReceipt()
   Dim RdoInv As ADODB.Recordset
   Dim RdoSum As ADODB.Recordset
   Dim RdoUse As ADODB.Recordset
   Dim bResponse As Byte
   Dim cPay As Currency
   Dim sInvoice As String
   
   sMsg = "This Will Permanently Remove The Cash Receipt." & vbCrLf _
          & "Do You Really Want To Cancel This Cash Receipt?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      Exit Sub
   End If
   
   'On Error Resume Next
   On Error GoTo whoops
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sCheck = Trim(cmbChk)
   
'
'   sSql = "SELECT DISTINCT DCINVNO, DCDEBIT, DCCREDIT, INVPRE, INVTYPE, INVPAY, INVORIGIN" & vbCrLf _
'      & "FROM JritTable " & vbCrLf _
'      & "JOIN CihdTable ON DCINVNO = INVNO" & vbCrLf _
'      & "WHERE DCCHECKNO = '" & sCheck & "' AND DCCUST = '" & sCust & "'"
   
   sSql = "SELECT DISTINCT DCINVNO, SUM(DCDEBIT) DCDEBIT, SUM(DCCREDIT) DCCREDIT," & vbCrLf _
            & "INVPRE , INVTYPE, INVPAY, INVORIGIN " & vbCrLf _
            & "From JritTable " & vbCrLf _
                  & "JOIN CihdTable ON DCINVNO = INVNO " & vbCrLf _
            & "WHERE DCCHECKNO = '" & sCheck & "' AND DCCUST = '" & sCust & "' AND DCDEBIT = 0 " & vbCrLf _
            & "GROUP BY DCINVNO, INVPRE, INVTYPE, INVPAY, INVORIGIN " & vbCrLf _
            & "UNION " & vbCrLf _
            & "SELECT DISTINCT DCINVNO, SUM(DCDEBIT) DCDEBIT, SUM(DCCREDIT) DCCREDIT, " & vbCrLf _
            & "   INVPRE , INVTYPE, INVPAY, INVORIGIN " & vbCrLf _
            & "From JritTable " & vbCrLf _
                  & "JOIN CihdTable ON DCINVNO = INVNO " & vbCrLf _
            & "WHERE DCCHECKNO = '" & sCheck & "' AND DCCUST = '" & sCust & "' AND DCCREDIT = 0 " & vbCrLf _
            & "GROUP BY DCINVNO, INVPRE, INVTYPE, INVPAY, INVORIGIN "
   
   Debug.Print sSql
   
   If clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC) Then
      With RdoInv
         Do Until .EOF
            If Not IsNull(!DCINVNO) Then
               sInvoice = "" & Trim(!INVPRE) & Format(!DCINVNO, "000000")
               
               'if applied elsewhere can't cancel
               If !INVTYPE = "CM" Or !INVTYPE = "CA" Then
                  If sCheck = "" & Trim(!INVORIGIN) Then
                     ' see if advance payment is applied elsewhere
                     sSql = "SELECT COUNT(DCCHECKNO) FROM JritTable WHERE DCINVNO = " _
                            & !DCINVNO & " AND DCCHECKNO <> '" & !INVORIGIN & "'"
                     bSqlRows = clsADOCon.GetDataSet(sSql, RdoUse)
                     If RdoUse.Fields(0) > 0 Then
                        clsADOCon.RollbackTrans
                        sMsg = "Cannot Cancel Cash Receipt." & vbCrLf _
                               & sInvoice & " Is Applied On Other Receipts."
                        MsgBox sMsg, vbInformation, Caption
                        'RdoUse.Cancel
                        'RdoSum.Cancel
                        'RdoInv.Cancel
                        'Set RdoUse = Nothing
                        'Set RdoSum = Nothing
                        'Set RdoInv = Nothing
                        Exit Sub
                     End If
                     If (Not RdoUse Is Nothing) Then RdoUse.Cancel
                     Set RdoUse = Nothing
                  End If
               End If
               
                      
'               sSql = "SELECT SUM(DCDEBIT),SUM(DCCREDIT)" & vbCrLf _
'                  & "FROM JritTable" & vbCrLf _
'                  & "INNER JOIN CashTable ON DCCHECKNO = CACHECKNO AND DCINVNO = CAINVNO" & vbCrLf _
'                  & "WHERE DCCHECKNO = '" & sCheck & "'" & vbCrLf _
'                  & "AND " & "DCCUST = '" & sCust & "' AND DCINVNO = " & !DCINVNO
               
               ' Not need to use SUM as the discount creates new Jrit record
               ' need to use SUM as the Top query is sum which includes discounts
               sSql = "SELECT SUM(DCDEBIT), SUM(DCCREDIT) " & vbCrLf _
                  & "FROM JritTable " & vbCrLf _
                  & "INNER JOIN CashTable ON DCCHECKNO = CACHECKNO AND DCINVNO = CAINVNO" & vbCrLf _
                  & "WHERE DCCHECKNO = '" & sCheck & "'" & vbCrLf _
                  & "AND " & "DCCUST = '" & sCust & "' AND DCINVNO = " & !DCINVNO
               
               '1/9/08 - REQUIRE DCDEBIT > 0 TO AVOID DUPLICATES
               If clsADOCon.GetDataSet(sSql, RdoSum, ES_STATIC) And !DCDEBIT > 0 Then
                  sSql = ""
                  
                  'if PS, SO, or DM:  Subtract total credit
                  If !INVTYPE = "PS" Or !INVTYPE = "SO" Or !INVTYPE = "DM" Then
                     sSql = "UPDATE CihdTable SET " _
                        & "INVCHECKNO=''," _
                        & "INVCHECKDATE = NULL," _
                        & "INVPAY=INVPAY - " & RdoSum.Fields(1) & "," _
                        & "INVPIF=0," _
                        & "INVARDISC=0," _
                        & "INVDAYS=0 " _
                        & "WHERE INVNO=" & !DCINVNO & " "
                  
'                  'if cash advance, zero INVPAY
'                  ElseIf !INVTYPE = "CA" Then
'                     sSql = "UPDATE CihdTable SET " _
'                        & "INVCHECKNO=''," _
'                        & "INVPAY=0," _
'                        & "INVPIF=0," _
'                        & "INVARDISC=0," _
'                        & "INVDAYS=0," _
'                        & "INVCANCELED=1 " _
'                        & "WHERE INVNO=" & !DCINVNO & " "

                  'if cash advance, adjust INVPAY
                  ElseIf !INVTYPE = "CA" Then
                  
                     'if original cash advance
                     'If !DCDEBIT > 0 Then
                        If Trim(!INVORIGIN) = Trim(sCheck) Then
                           sSql = "UPDATE CihdTable SET " _
                              & "INVCHECKNO=''," _
                              & "INVCHECKDATE = NULL," _
                              & "INVPAY=0," _
                              & "INVPIF=0," _
                              & "INVARDISC=0," _
                              & "INVDAYS=0," _
                              & "INVCANCELED=1 " _
                              & "WHERE INVNO=" & !DCINVNO & " "
                           
                        'if application of original cash advance
                        'there really ought to be a different INVTYPE,
                        'or an entirely different table
                        Else
                           sSql = "UPDATE CihdTable SET " _
                              & "INVCHECKNO=''," _
                              & "INVCHECKDATE = NULL," _
                              & "INVPAY = INVPAY + " & !DCDEBIT & "," _
                              & "INVPIF=0," _
                              & "INVARDISC=0," _
                              & "INVDAYS=0" & vbCrLf _
                              & "WHERE INVNO=" & !DCINVNO & " "
                        End If
                     'End If
                  'for others (CM): Add total debit
                  Else
                     sSql = "UPDATE CihdTable SET " _
                        & "INVCHECKNO=''," _
                        & "INVCHECKDATE = NULL," _
                        & "INVPAY=INVPAY + " & RdoSum.Fields(0) & "," _
                        & "INVPIF=0," _
                        & "INVARDISC=0," _
                        & "INVDAYS=0 " & vbCrLf _
                        & "WHERE INVNO=" & !DCINVNO & " "
                  End If
                  
                  If sSql <> "" Then
Debug.Print sSql
                     clsADOCon.ExecuteSql sSql
                  End If
               End If
               If (Not RdoSum Is Nothing) Then RdoSum.Cancel
               Set RdoSum = Nothing
            End If
            .MoveNext
         Loop
         .Cancel
      End With
   End If

   Set RdoInv = Nothing
   
   sSql = "DELETE FROM CashTable WHERE CACHECKNO='" _
          & cmbChk & "' AND CACUST='" & sCust & "' "
   clsADOCon.ExecuteSql sSql
   
   sSql = "DELETE FROM JritTable WHERE " _
          & "DCCUST = '" & sCust & "' AND " _
          & "DCCHECKNO = '" & cmbChk & "'"
   clsADOCon.ExecuteSql sSql
   
  ' 45738:Cancelling a cash reciept doubles the invoice amount (w/discount)
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "Successfully Canceled.", True
      sCust = cmbCst
      FillCombo
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Could Not Cancel Cash Receipt.", _
         vbExclamation, Caption
   End If
   
   MouseCursor 0
   Exit Sub
'DiaErr1:
'   sProcName = "cancelrece"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me

whoops:
   ProcessError "CancelReceipt"
End Sub

Private Sub IniGrid()
   With Grid1
      .Clear
      .Rows = 1
      .Cols = 4
      .Row = 0
      
      .Col = 0
      .ColWidth(0) = 1500
      .Text = "Invoice #"
      
      .Col = 1
      .ColWidth(1) = 800
      .Text = "Inv Type"
      
      .Col = 2
      .ColWidth(2) = 1500
      .Text = "Inv Date"
      
      .Col = 3
      .ColWidth(3) = 1500
      .Text = "Applied"
   End With
End Sub

Private Function GetCrType(b As Byte) As String
   Select Case b
      Case 0
         GetCrType = "Cash"
      Case 1
         GetCrType = "Wire"
      Case 2
         GetCrType = "Check"
      Case Else
         GetCrType = "None"
   End Select
End Function
