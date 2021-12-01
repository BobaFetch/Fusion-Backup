VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPe11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "External Check (No Invoice)"
   ClientHeight    =   6615
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6615
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtExp 
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   16
      Tag             =   "1"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   14
      Tag             =   "1"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Tag             =   "1"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Tag             =   "1"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Tag             =   "1"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Tag             =   "1"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox cmbExp 
      Height          =   315
      Index           =   5
      Left            =   1560
      TabIndex        =   17
      ToolTipText     =   "Expense Account To Debit"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox cmbExp 
      Height          =   315
      Index           =   4
      Left            =   1560
      TabIndex        =   15
      ToolTipText     =   "Expense Account To Debit"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.ComboBox cmbExp 
      Height          =   315
      Index           =   3
      Left            =   1560
      TabIndex        =   13
      ToolTipText     =   "Expense Account To Debit"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox cmbExp 
      Height          =   315
      Index           =   2
      Left            =   1560
      TabIndex        =   11
      ToolTipText     =   "Expense Account To Debit"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox cmbExp 
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      ToolTipText     =   "Expense Account To Debit"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox cmbExp 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      ToolTipText     =   "Expense Account To Debit"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Height          =   315
      Left            =   4920
      TabIndex        =   19
      ToolTipText     =   "Post This Check"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Vendors"
      Top             =   1200
      Width           =   1555
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Tag             =   "4"
      ToolTipText     =   "Check Date"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtChk 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Next Check Number For Checking Account"
      Top             =   1920
      Width           =   1185
   End
   Begin VB.TextBox txtAmt 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Check Amount"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtMemo 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Memo (Optional)"
      Top             =   3000
      Width           =   3135
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Checking Account To Credit"
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
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
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPe11a.frx":0000
      PictureDn       =   "diaAPe11a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4080
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6615
      FormDesignWidth =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Account"
      Height          =   195
      Left            =   1560
      TabIndex        =   38
      Top             =   3660
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Amount"
      Height          =   195
      Left            =   420
      TabIndex        =   37
      Top             =   3660
      Width           =   1035
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5760
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   360
      TabIndex        =   36
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   35
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   34
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   33
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   32
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   31
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   30
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Distribution Accounts:"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   5745
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   28
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Amount"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   2640
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Memo"
      Height          =   285
      Index           =   18
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Checking Account"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   21
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "diaAPe11a"
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
' diaAPe11a - External Check No Invoice
'
' Notes:
'
' Created: 01/08/04 (nth)
' Revisions:
' 01/23/04 (nth) Clear boxs after check is posted.
' 05/11/04 (nth) Added multiple account distributions.
' 08/31/04 (nth) Added ablility to have a negative expense ie negative check per JEVINT.
' 01/04/05 (nth) Set check type to 3 (no invoice).
' 03/17/05 cjs Corrected Qry_FillLowAccounts (casing)
' 03/17/05 cjs Corrected empty Combo
'*********************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodVendor As Byte
Dim sXcAcct As String
Dim sApAcct As String
Dim sJournalID As String
Dim sMsg As String

Dim sDefaultAccount As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbAct_Click()
   lbldsc = UpdateActDesc(cmbAct, lbldsc, True)
   
End Sub

Private Sub cmbAct_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbAct_LostFocus()
   Dim lCheck As Double
   lbldsc = UpdateActDesc(cmbAct, lbldsc, True)
   sXcAcct = Compress(cmbAct)
   lCheck = GetNextCheck(sXcAcct)
   txtChk = lCheck
End Sub

Private Sub cmbExp_Click(Index As Integer)
   If Trim(txtExp(Index)) = "" Then
      lblExp(Index) = ""
      cmbExp(Index) = ""
   Else
      If CCur(txtExp(Index)) = 0 Then
         lblExp(Index) = ""
         cmbExp(Index) = ""
      Else
         lblExp(Index) = UpdateActDesc(cmbExp(Index), lblExp(Index), True)
      End If
   End If
End Sub

Private Sub cmbExp_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub cmbExp_LostFocus(Index As Integer)
   If Trim(txtExp(Index)) = "" Then
      lblExp(Index) = ""
      cmbExp(Index) = ""
   Else
      If IsNumeric(txtExp(Index).Text) Then
         If CCur(txtExp(Index)) = 0 Then
            lblExp(Index) = ""
            cmbExp(Index) = ""
         Else
            lblExp(Index) = UpdateActDesc(cmbExp(Index), lblExp(Index), True)
         End If
      End If
   End If
End Sub

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendorAndAccount
   If (bGoodVendor) Then
      cmbExp(0) = sDefaultAccount
   End If
End Sub

Private Sub cmbVnd_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbVnd_LostFocus()
   bGoodVendor = FindVendorAndAccount
   If (bGoodVendor) Then
      cmbExp(0) = sDefaultAccount
   End If
End Sub
   
Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdPst_Click()
   PostCheck
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      sDefaultAccount = ""
      CurrentJournal "XC", ES_SYSDATE, sJournalID
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   txtAmt = "0.00"
   bOnLoad = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodVendor Then
      cUR.CurrentVendor = cmbVnd
      SaveCurrentSelections
   End If
   FormUnload
   Set diaAPe11a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   Dim rdoAct As ADODB.Recordset
   Dim RdoExp As ADODB.Recordset
   Dim b As Byte
   'On Error GoTo DiaErr1
   On Error GoTo 0
   FillVendors Me
   
   If (FindVendorAndAccount) Then
      If (cmbExp(0) = "") Then cmbExp(0) = sDefaultAccount
   End If
   
   ' Cash Accounts
   cmbAct.Clear
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      cmbAct.ListIndex = 0
   End If
   Set rdoAct = Nothing
   lbldsc = UpdateActDesc(cmbAct)
   
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoExp)
   If bSqlRows Then
      With RdoExp
         While Not .EOF
            For b = 0 To 5
               AddComboStr cmbExp(b).hWnd, "" & Trim(!GLACCTNO)
            Next
            .MoveNext
         Wend
      End With
   End If
   Set RdoExp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtAmt_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtAmt_LostFocus()
   If Trim(txtAmt) = "" Then
      txtAmt = "0.00"
   Else
      txtAmt = Format(txtAmt, CURRENCYMASK)
   End If
End Sub

Private Sub txtChk_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtChk_LostFocus()
   If Not bCancel Then
      If Not ValidateCheck(txtChk, cmbAct) Then
         sMsg = "Check " & txtChk & " Exists For Account " & cmbAct & "."
         MsgBox sMsg, vbInformation, Caption
         txtChk.SetFocus
      End If
   End If
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDte_LostFocus()
   If Not bCancel Then
      txtDte = CheckDate(txtDte)
      ' Check if journal exist
      sJournalID = GetOpenJournal("XC", Format(txtDte, "mm/dd/yy"))
      If sJournalID = "" Then
         sMsg = "There Is No Open External Check Cash Disbursements" _
                & vbCrLf & "Journal For " & txtDte
         MsgBox sMsg, vbInformation, Caption
         txtDte.SetFocus
      End If
   End If
End Sub

Private Sub txtExp_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtExp_LostFocus(Index As Integer)
   If Trim(txtExp(Index).Text) = "" Then
      Exit Sub
   End If
   
   If Not IsNumeric(txtExp(Index).Text) Then
      MsgBox "Amounts must be numeric"
      txtExp(Index).SetFocus
      Exit Sub
   End If
   If Trim(txtExp(Index)) <> "" Then
      txtExp(Index) = Format(txtExp(Index), CURRENCYMASK)
      'MsgBox "txtExp(index)=" & txtExp(Index)
      
      If CCur(txtExp(Index)) = 0 Then
         cmbExp(Index) = ""
         lblExp(Index) = ""
      Else
        If cmbExp(Index) = "" Then cmbExp(Index) = sDefaultAccount
      End If
   Else
      cmbExp(Index) = ""
      lblExp(Index) = ""
   End If
   UpdateTotals
End Sub

Private Sub txtMemo_GotFocus()
   SelectFormat Me
End Sub

Private Sub PostCheck()
   Dim bResponse As String
   Dim lTrans As Long
   Dim iRef As Integer
   Dim sApAcct As String
   Dim sXcAcct As String
   Dim cTotal As Currency
   Dim sVendor As String
   Dim sNow As String
   Dim sCheck As String
   Dim b As Byte
   Dim cExpense As Currency
   Dim sColumn As String
   
   On Error GoTo DiaErr1
   
   If lblTot <> txtAmt Then
      sMsg = "Distribution Total Does Not Match Check Amount."
      MsgBox sMsg, vbInformation, Caption
      txtExp(0).SetFocus
      Exit Sub
   End If
   For b = 0 To 5
      If Trim(txtExp(b)) <> "" Then
         If Trim(cmbExp(b)) = "" And CCur(txtExp(b)) > 0 Then
            sMsg = "One Or More Distributions Are Missing An Account."
            MsgBox sMsg, vbInformation, Caption
            cmbExp(b).SetFocus
            Exit Sub
         End If
      End If
   Next
   sMsg = "Are You Ready To Post This Payment?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      lTrans = GetNextTransaction(sJournalID)
      cTotal = CCur("0" & txtAmt)
      sVendor = Compress(cmbVnd)
      sNow = Format(ES_SYSDATE, "mm/dd/yy")
      sCheck = Trim(txtChk)
      sXcAcct = Compress(cmbAct)
      sApAcct = Compress(cmbExp)
      Err = 0
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      ' Debit\Credit Expense
      For b = 0 To 5
         cExpense = CCur(txtExp(b) & "0")
         If cExpense <> 0 Then
            If cExpense < 0 Then
               sColumn = "DCCREDIT"
            Else
               sColumn = "DCDEBIT"
            End If
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sColumn & "," _
                   & "DCACCTNO,DCDATE,DCCHECKNO,DCCHKACCT,DCVENDOR) " _
                   & "VALUES('" _
                   & sJournalID & "'," _
                   & lTrans & "," _
                   & iRef & "," _
                   & Abs(cExpense) & ",'" _
                   & Compress(cmbExp(b)) & "','" _
                   & txtDte & "','" _
                   & sCheck & "','" _
                   & sXcAcct & "','" _
                   & sVendor & "')"
            clsADOCon.ExecuteSql sSql
         End If
      Next
      ' Credit Checking
      If cTotal > 0 Then
         iRef = iRef + 1
         sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT,DCACCTNO," _
                & "DCDATE,DCCHECKNO,DCCHKACCT,DCVENDOR) " _
                & "VALUES('" _
                & sJournalID & "'," _
                & lTrans & "," _
                & iRef & "," _
                & cTotal & ",'" _
                & sXcAcct & "','" _
                & txtDte & "','" _
                & sCheck & "','" _
                & sXcAcct & "','" _
                & sVendor & "')"
         clsADOCon.ExecuteSql sSql
      End If
      ' Add Check Record
      sSql = "INSERT INTO ChksTable (CHKNUMBER,CHKVENDOR,CHKAMOUNT," _
             & "CHKPOSTDATE,CHKACTUALDATE,CHKBY,CHKMEMO,CHKVOID,CHKPRINTED,CHKTYPE,CHKACCT)" _
             & "VALUES('" _
             & sCheck & "','" _
             & sVendor & "'," _
             & cTotal & ",'" _
             & txtDte & "','" _
             & txtDte & "','" _
             & Secure.UserInitials & "','" _
             & Trim(txtMemo) & "'," _
             & "0,0,3,'" & sXcAcct & "')"
      clsADOCon.ExecuteSql sSql
      
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         sMsg = "Successfully Posted Check."
         SysMsg sMsg, True, Me
         SaveLastCheck sCheck, sXcAcct
         txtDte = Format(ES_SYSDATE, "mm/dd/yy")
         txtAmt = "0.00"
         txtMemo = ""
         txtChk = ""
         cmbVnd.ListIndex = 0
         cmbAct.SetFocus
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         sMsg = "Cannot Post Check."
         MsgBox sMsg, vbExclamation, Caption
      End If
   End If
   Exit Sub
DiaErr1:
   sProcName = "postcheck"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtMemo_LostFocus()
   txtMemo = CheckLen(txtMemo, 40)
   CheckComments txtMemo
End Sub

Private Sub UpdateTotals()
   Dim b As Byte
   Dim cTotal As Currency
   On Error Resume Next
   For b = 0 To 5
      If Trim(txtExp(b)) <> "" Then
         cTotal = cTotal + CCur(txtExp(b))
      End If
   Next
   lblTot = Format(cTotal, CURRENCYMASK)
   If lblTot <> txtAmt Then
      lblTot.ForeColor = ES_RED
   Else
      lblTot.ForeColor = diaAPe11a.ForeColor
   End If
   
End Sub


Private Function FindVendorAndAccount() As Byte
   Dim RdoVnd As ADODB.Recordset
   Dim sVendRef As String
   
   FindVendorAndAccount = 0
   sVendRef = Compress(cmbVnd)
   If Len(sVendRef) = 0 Then Exit Function
   
   sSql = "SELECT VEREF, VENICKNAME, VEBNAME, VEACCOUNT FROM VndrTable WHERE VEREF='" & sVendRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
   On Error Resume Next
   If bSqlRows Then
      With RdoVnd
         cmbVnd = "" & Trim(!VENICKNAME)
         lblNme = "" & Trim(!VEBNAME)
         sDefaultAccount = "" & Trim(!VEACCOUNT)
         FindVendorAndAccount = 1
         .Cancel
      End With
   Else
      cmbVnd = ""
      lblNme = ""
      sDefaultAccount = ""
   End If
   Set RdoVnd = Nothing
   Exit Function
End Function

