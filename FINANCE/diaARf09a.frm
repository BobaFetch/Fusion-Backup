VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form diaARf09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Cash Receipt from Excel Sheet"
   ClientHeight    =   8025
   ClientLeft      =   1845
   ClientTop       =   1080
   ClientWidth     =   12660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNoInv 
      Cancel          =   -1  'True
      Caption         =   "Show Invoice not Found"
      Height          =   360
      Left            =   6720
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2145
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Customer Nicknames"
      Top             =   960
      Width           =   1555
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Tag             =   "8"
      ToolTipText     =   "Cash Account"
      Top             =   480
      Width           =   1555
   End
   Begin VB.CommandButton cmdCR 
      Caption         =   "Create CR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10440
      TabIndex        =   9
      ToolTipText     =   " Create Cash Receipts"
      Top             =   2880
      Width           =   1920
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10440
      TabIndex        =   8
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   3600
      Width           =   1920
   End
   Begin VB.TextBox txtXLFilePath 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   1680
      Width           =   4695
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Excel AR data"
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2145
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "diaARf09a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8025
      FormDesignWidth =   12660
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   7200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4935
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   8
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   14
      Top             =   990
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   195
      Index           =   16
      Left            =   720
      TabIndex        =   13
      Top             =   520
      Width           =   705
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   12
      Top             =   480
      Width           =   2775
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7320
      Picture         =   "diaARf09a.frx":07AE
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7320
      Picture         =   "diaARf09a.frx":0B38
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1305
   End
End
Attribute VB_Name = "diaARf09a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Added ITINVOICE

Option Explicit
Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean

Dim sCrCashAcct As String
Dim sCrDiscAcct As String
Dim sSJARAcct As String
Dim sCrCommAcct As String
Dim sCrRevAcct As String
Dim sCrExpAcct As String
Dim sTransFeeAcct As String
Dim strOffsetAcct As String

Dim sAccount As String
Dim sMsg As String

Private txtKeyPress As New EsiKeyBd


Private Sub cmdCan_Click()
   Unload Me

End Sub

Private Sub cmdCR_Click()

   Dim iList As Long
   Dim strCusName As String
   Dim strChkNum As String
   
   
   For iList = 1 To Grd.rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
         
         Grd.Col = 1
         strChkNum = Trim(Grd.Text)
         If Not CheckExists(strChkNum) And strChkNum <> "" Then
            strCusName = cmbCst
            PostARCheck strChkNum, strCusName
         End If
      End If
   Next
   
   ' Refresh the grid
   cmdImport_Click
End Sub

Private Function CheckExists(sCheck As String) As Byte
   Dim RdoChk As ADODB.Recordset
   
   On Error GoTo DiaErr1
   CheckExists = False
   sSql = "SELECT COUNT(CACHECKNO) FROM CashTable WHERE CACHECKNO = '" _
          & sCheck & "' AND CACUST = '" & Compress(cmbCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         If .Fields(0) > 0 Then
            CheckExists = True
            sMsg = "Check # " & sCheck & " All Ready Exists " & vbCrLf _
                   & "For Customer " & cmbCst
            MsgBox sMsg, vbInformation, Caption
         End If
         .Cancel
      End With
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkexists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function

Private Function GetInvCustName(iInvNum As Double) As String
   Dim RdoInv As ADODB.Recordset
   Dim strInvCust As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT INVCUST FROM CihdTable WHERE INVNO = " & iInvNum
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         strInvCust = !INVCUST
      End With
   Else
      strInvCust = ""
   End If
   Set RdoInv = Nothing
   
   GetInvCustName = strInvCust
   
   Exit Function
   
DiaErr1:
   sProcName = "checkexists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function
Private Function PostARCheck(strChkNum As String, strCusName As String)
   Dim i As Integer
   Dim lTrans As Long
   Dim iRef As Integer
   'Dim cAmount   As Currency
   Dim strCDate As String
   Dim strRDate As String
   Dim strInvNum As String
   Dim pif As Integer
   '    Dim sCredit   As String
   '    Dim sDebit    As String
   Dim lNewInv As Long
   Dim rdoAct As ADODB.Recordset
   Dim TotalReceipt As Currency
   Dim cInvAmt As Currency
   Dim cInvDisc As Currency
   Dim strPONum As String
   
   On Error GoTo DiaErr1
   
   strCDate = Format(ES_SYSDATE, "mm/dd/yy")
'   sAdvAct = Compress(cmbAdvAct)
'   sCrCashAcct = Compress(cmbAct)
'   sCheck = Trim(txtChk)
'   TotalReceipt = CCur("0" & txtAmt)
   
   
   ' Allow posting to journals other than that of the current month
   sJournalID = GetOpenJournal("CR", strCDate)
   If sJournalID <> "" Then
      'lTrans = GetNextTransaction(sJournalID)
   Else
      sMsg = "No Open Cash Recipts Journal Found For " & strRDate & "."
      MsgBox sMsg, vbInformation, Caption
      Exit Function
   End If
   
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0

   Dim gl As GLTransaction
   Set gl = New GLTransaction
   gl.JournalID = sJournalID 'automatically sets next transaction
   
   
   Dim RdoInv As ADODB.Recordset
'   sSql = "SELECT CHECKNUMBER,INVNUMBER, INVAMOUNT,INVDISAMT, " _
'            & "PONUMBER,INVRECDATE FROM ImpBIVSInv " _
'            & "   Where CHECKNUMBER = '" & strChkNum & "'" _
'            & "order by INVNUMBER"
   
   sSql = "SELECT CHECKNUMBER,INVNUMBER, INVAMOUNT,INVDISAMT, PONUMBER,INVRECDATE " _
            & "From ImpBIVSInv, Cihdtable " _
            & "Where IsNumeric(INVNUMBER) = 1 " _
            & "   AND invno = CONVERT(bigint, INVNUMBER) " _
            & "   AND INVPIF = 0 " _
            & "    AND INVCHECKNO = '' " _
            & "    AND CHECKNUMBER <> '' " _
            & "    AND INVCUST IN (SELECT CPPAYER FROM dbo.CpayTable WHERE CPCUST = '" & strCusName & "') " _
            & "   AND checknumber = '" & strChkNum & "' " _
            & "Union " _
            & "SELECT CHECKNUMBER,INVNUMBER, INVAMOUNT,INVDISAMT, PONUMBER,INVRECDATE " _
            & "FROM ImpBIVSInv Where CHECKNUMBER = '" & strChkNum & "' AND PONUMBER = 'NO PO INVOICE' " _
            & "   AND INVNUMBER NOT IN (SELECT CACHECKNO FROM CashTable WHERE CACASHACCT = '" & strOffsetAcct & "')" _
            & " ORDER BY INVNUMBER"
   
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC)
   
   If bSqlRows Then
      With RdoInv
      While Not .EOF
        
         strChkNum = !CHECKNUMBER
         strInvNum = !INVNUMBER
         cInvAmt = !INVAMOUNT
         cInvDisc = !INVDISAMT
         strPONum = !PONumber
         strRDate = Format(Now, "mm/dd/yy")
         pif = 1
         
         If cInvAmt < 0 Then
            Debug.Print cInvAmt
         End If
         
         gl.InvoiceDate = strRDate
      
         'if pif, also calculate discount
         Dim amountPaid As Currency
         Dim debitCash As Currency
         Dim creditAR As Currency
         Dim creditOther As Currency
         Dim discountAmount As Currency
         Dim commissionAmount As Currency
         Dim revenueAmount As Currency
         Dim expenseAmount As Currency
         Dim DisAcct As String
         Dim bCaType As Byte
         Dim strInvCust As String
         Dim iInvNum As Double
         

         creditAR = CCur(cInvAmt)
         discountAmount = -CCur(cInvDisc)
         amountPaid = CCur(cInvAmt) - CCur(cInvDisc)
         TotalReceipt = TotalReceipt + amountPaid
         
         commissionAmount = 0
         revenueAmount = 0
         expenseAmount = 0
         ' CA type is 2 for Check
         bCaType = 2
         
         
         If (strPONum <> "NO PO INVOICE") Then
            
            
            If (IsNumeric(strInvNum)) Then
               gl.InvoiceNumber = Val(strInvNum)
               iInvNum = Val(strInvNum)
               strInvCust = GetInvCustName(iInvNum)
            Else
               gl.InvoiceNumber = 0
            End If
            ' set the Cr/Db disc account
            DisAcct = sCrDiscAcct
            
            If (strInvCust = "") Then strInvCust = strCusName
            'create debit or credit
            'note: AddDebitCredit turns a - credit into a debit
            If Len(DisAcct) > 0 Then
               gl.AddDebitCredit 0, discountAmount, DisAcct, "", 0, 0, "", strInvCust, CStr(Val(strChkNum))
            End If
            
            If (IsNumeric(strInvNum)) Then
               sSql = "UPDATE CihdTable SET " _
                      & "INVCHECKNO='" & Val(strChkNum) & "'," _
                      & "INVPAY=" & creditAR & "," _
                      & "INVPIF=" & pif & "," _
                      & "INVCHECKDATE='" & strRDate & "' " _
                      & "WHERE INVNO=" & iInvNum
               clsADOCon.ExecuteSql sSql
            End If
            ' Add a cash receipt record
            gl.AddCashReceipt CStr(Val(strChkNum)), strCusName, bCaType, strRDate, amountPaid, _
               creditAR, discountAmount, commissionAmount, expenseAmount, revenueAmount, strRDate, sCrCashAcct
            
            ' Add Journal entries
            gl.AddDebitCredit 0, creditAR, sSJARAcct, "", 0, 0, "", strCusName, CStr(Val(strChkNum))
            gl.AddDebitCredit amountPaid, 0, sCrCashAcct, "", 0, 0, "", strCusName, CStr(Val(strChkNum))
         
         Else
            ' Add a cash receipt record
            gl.AddCashReceipt strInvNum, strCusName, bCaType, strRDate, Abs(amountPaid), _
               Abs(creditAR), Abs(discountAmount), commissionAmount, expenseAmount, revenueAmount, strRDate, strOffsetAcct
            
            ' get Offset account
            'Debit Offset account (82000000)
            ' Credit SJ AR account
            gl.InvoiceNumber = 0
            
            gl.AddDebitCredit 0, Abs(creditAR), sSJARAcct, "", 0, 0, "", strCusName, strInvNum
            gl.AddDebitCredit Abs(amountPaid), 0, strOffsetAcct, "", 0, 0, "", strCusName, strInvNum
         End If
         
      .MoveNext
      Wend
      .Close
      End With
      
      ' Not Update the Total Amount on the check.
      sSql = "UPDATE CashTable SET CACKAMT = " & TotalReceipt & " WHERE CACHECKNO = '" & CStr(Val(strChkNum)) & "'"
      clsADOCon.ExecuteSql sSql '
      
   End If
   
   Set RdoInv = Nothing
   
   Dim bResponse As Byte
   bResponse = 0
   
   If clsADOCon.ADOErrNum = 0 And gl.Commit Then
      clsADOCon.CommitTrans
      SysMsg "Cash Receipt Successfully Posted", 1
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      sMsg = "Cash Receipt Not Successfully Posted." _
             & vbCrLf & " Transaction Canceled."
      MsgBox sMsg, vbExclamation, Caption
   End If
   
   Exit Function
DiaErr1:
   sProcName = "sPostCheck"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Private Sub cmdHlp_Click()
    If cmdHlp Then
        MouseCursor (13)
        OpenHelpContext (2150)
        MouseCursor (0)
        cmdHlp = False
    End If

End Sub

Private Sub cmdImport_Click()
   Dim strWindows As String
   Dim strAccFileName As String
   Dim strFilePath As String
   
   On Error GoTo DiaErr1
   strFilePath = txtXLFilePath.Text
   
   If (Trim(strFilePath) = "") Then
      MsgBox "Please select a Excel file to create Cash Receipt.", _
            vbInformation, Caption
      Exit Sub
   End If

   MouseCursor 13
   DeleteOldData ("ImpBIVSInv")
   ParseInvoiceDetail (strFilePath)
   
   
   Dim strCust As String
   strCust = cmbCst.Text
   
   sSql = "SELECT CHECKNUMBER,INVNUMBER, INVAMOUNT,INVDISAMT," _
            & "   PONumber , INVRECDATE, INVSTATUS,INVCUST FROM ImpBIVSInv, Cihdtable" _
            & "      Where IsNumeric(INVNUMBER) = 1" _
            & "            AND invno = CONVERT(bigint, INVNUMBER)" _
            & "            AND INVPIF = 0" _
            & "            AND INVCHECKNO = ''" _
            & "             AND CHECKNUMBER <> ''" _
            & "            AND INVCUST IN (SELECT CPPAYER FROM dbo.CpayTable WHERE CPCUST = '" & strCust & "')" _
            & "      Union" _
            & "      SELECT CHECKNUMBER,INVNUMBER, INVAMOUNT,INVDISAMT," _
            & "         PONumber , INVRECDATE, INVSTATUS,'" & strCust & "'" _
            & "      FROM ImpBIVSInv Where PONUMBER = 'NO PO INVOICE'" _
            & "         AND CHECKNUMBER <> ''" _
            & "     AND  INVNUMBER NOT IN (SELECT CACHECKNO FROM CashTable WHERE CACASHACCT = '" & strOffsetAcct & "')" _
            & "      order by CHECKNUMBER"

'   sSql = "select CHECKNUMBER,INVNUMBER, INVAMOUNT,INVDISAMT, " _
'            & "PONumber , INVRECDATE, INVSTATUS,INVCUST " _
'            & "FROM ImpBIVSInv WHERE CHECKNUMBER IN ( " _
'            & "SELECT DISTINCT CHECKNUMBER FROM ImpBIVSInv, Cihdtable " _
'            & "   Where IsNumeric(INVNUMBER) = 1 " _
'            & "AND invno = CONVERT(int, INVNUMBER) " _
'            & "   AND INVPIF = 0 AND CHECKNUMBER <> '' " _
'            & "AND INVCUST IN (SELECT CPPAYER FROM dbo.CpayTable WHERE CPCUST = '" & strCust & "'))" _
'            & " ORDER by CHECKNUMBER "

   
   FillGrid (sSql)
   cmdCR.Visible = True
   
   MouseCursor 0
   
   Exit Sub
DiaErr1:
   MouseCursor 0
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub cmdNoInv_Click()

   Dim strCust As String
   strCust = cmbCst.Text

   sSql = "select CHECKNUMBER,INVNUMBER, INVAMOUNT,INVDISAMT, PONumber,INVRECDATE," _
         & "  INVSTATUS,'" & strCust & "' as INVCUST FROM ImpBIVSInv WHERE INVSTATUS = 'paid'" _
         & "  AND CONVERT(varchar(24), INVNUMBER) NOT IN " _
         & "  (SELECT DISTINCT convert(varchar(24), INVNO) FROM cihdtable) AND " _
               & " PONumber <> 'NO PO INVOICE' order by CHECKNUMBER"

'               & "  --WHERE invcheckno = CHECKNUMBER) AND " _

   FillGrid (sSql)

   cmdCR.Visible = False
   
End Sub

Private Sub cmdOpenDia_Click()
   fileDlg.Filter = "Excel Files (*.xls) | *.xls|"
   
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
       txtXLFilePath.Text = ""
   Else
       txtXLFilePath.Text = fileDlg.filename
   End If
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   
   If bOnLoad Then
      CurrentJournal "CR", ES_SYSDATE, sJournalID
      b = GetCashAccounts()
      If b = 3 Then
         MouseCursor 0
         MsgBox "One Or More Receivable Accounts Are Not Active." & vbCr _
            & "Please Set All Cash Accounts In The " & vbCr _
            & "System Setup, Administration Section.", _
            vbInformation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
      MouseCursor 13
      FillCombo
      If cUR.CurrentCustomer <> "" Then
         cmbCst = cUR.CurrentCustomer
      End If
      bOnLoad = 0
   End If
    
   MouseCursor (0)

End Sub

Public Function GetCashAccounts() As Byte
   Dim rdoCsh As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT COGLVERIFY,COCRCASHACCT,COCRDISCACCT,COSJARACCT," _
          & "COCRCOMMACCT,COCRREVACCT,COCREXPACCT,COTRANSFEEACCT,COBIVSOFFSETACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCsh, ES_FORWARD)
   sProcName = "getcashacct"
   If bSqlRows Then
      With rdoCsh
         For i = 1 To 7
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
         sTransFeeAcct = "" & Trim(!COTRANSFEEACCT)
         strOffsetAcct = IIf(IsNull(!COBIVSOFFSETACCT), "82000000", Trim(!COBIVSOFFSETACCT))
         If (strOffsetAcct = "") Then strOffsetAcct = "82000000"
         
         .Cancel
         If b = 1 Then GetCashAccounts = 3 Else GetCashAccounts = 2
      End With
   Else
      GetCashAccounts = 0
   End If
   Set rdoCsh = Nothing
   
   ' set the Offset account
   ' TODO: company setting
   'strOffsetAcct = "82000000"

   Exit Function
   
DiaErr1:
   sProcName = "getcashacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub FillCombo()
   Dim rdoCst As ADODB.Recordset
   Dim rdoAct As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   Dim lCheck As Long
   Dim sOptions As String
   
   
   On Error GoTo DiaErr1
   
   ' Customer
   sSql = "SELECT DISTINCT CUNICKNAME FROM CustTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         cmbCst = "" & Trim(!CUNICKNAME)
         While Not .EOF
            AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
            .MoveNext
         Wend
      End With
   End If
   Set rdoCst = Nothing
   
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      
      SetComboBox cmbAct, sAccount
   Else
      ' Multiple cash accounts not found so use
      ' the default cash account
'      cmbAct.enabled = False
'      cmbAct = sCrCashAcct
      AddComboStr cmbAct.hWnd, sCrCashAcct
      cmbAct.ListIndex = 0
      cmbAct.enabled = False
   
   End If
   lblDsc(1) = UpdateActDesc(cmbAct)
   
   Set rdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' make sure that you release the Hook
   Call WheelUnHook(Me.hWnd)
   
End Sub
Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
   
      .rows = 1
      .Row = 0
      .Col = 0
      .Text = "Apply"
      .Col = 1
      .Text = "CheckNumber"
      .Col = 2
      .Text = "Cust Name"
      .Col = 3
      .Text = "InvoiceNumber"
      .Col = 4
      .Text = "InvoiceAmount"
      .Col = 5
      .Text = "InvoiceDisc"
      .Col = 6
      .Text = "PO Number"
      .Col = 7
      .Text = "ReceivedDate"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1500
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
      .ColWidth(6) = 1350
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   

   Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub

Function FillGrid(sSql As String) As Integer
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   Dim iItem  As Integer
   Dim strInvNum As String
   Dim cInvAmt As Double
   Dim cInvDisc As Double
   Dim strPONum As String
   Dim strChkNum As String
   Dim strPrevChkNum As String
   Dim strRecDate As String
   Dim strInvStatus As String
   Dim strDuedate As String
   Dim strDisDate As String
   Dim strInvCust As String
   

   Debug.Print sSql
   
   Dim RdoExo As ADODB.Recordset
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoExo, ES_STATIC)
   
   Grd.rows = 1
   strPrevChkNum = ""
   If bSqlRows Then
      With RdoExo
      While Not .EOF
         strChkNum = !CHECKNUMBER
         strInvNum = !INVNUMBER
         cInvAmt = !INVAMOUNT
         cInvDisc = !INVDISAMT
         strPONum = !PONumber
         strRecDate = !INVRECDATE
         strInvCust = !INVCUST
         
         Grd.rows = Grd.rows + 1
         Grd.Row = Grd.rows - 1
         
         If (strChkNum <> strPrevChkNum) Then
            
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            
            Grd.Col = 1
            Grd.Text = Trim(strChkNum)
            
            strPrevChkNum = strChkNum
         End If
         
         Grd.Col = 2
         Grd.Text = Trim(strInvCust)
            
         Grd.Col = 3
         Grd.Text = Trim(strInvNum)
         Grd.Col = 4
         Grd.Text = Trim(CStr(cInvAmt))
         Grd.Col = 5
         Grd.Text = Trim(cInvDisc)
         Grd.Col = 6
         Grd.Text = Trim(strPONum)
         Grd.Col = 7
         Grd.Text = Trim(strRecDate)
         .MoveNext
         
      Wend
      .Close
      End With
   End If
   
   Set RdoExo = Nothing
   MouseCursor ccArrow
       
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormUnload
    Set diaARf09a = Nothing
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 0
      If Grd.Row >= 1 Then
         If Grd.Row = 0 Then Grd.Row = 1
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
      End If
    End If
   

End Sub


Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd.rows - 1
        Grd.Col = 0
        Grd.Row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
        End If
    Next
End Sub


Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grd.Col = 0
   If Grd.Row >= 1 Then
      If Grd.Row = 0 Then Grd.Row = 1
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
      Else
         Set Grd.CellPicture = Chkyes.Picture
      End If
   End If
End Sub


Private Function DeleteOldData(strTableName As String)

   If (strTableName <> "") Then
      sSql = "DELETE FROM " & strTableName
      clsADOCon.ExecuteSql sSql
   End If

End Function

Private Function ParseInvoiceDetail(strFullPath As String)

   Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim strInvNum As String
   Dim strInvAmt As String
   Dim strInvDisc As String
   Dim strPONum As String
   Dim strChkNum As String
   Dim strRecDate As String
   Dim strInvStatus As String
   Dim strDuedate As String
   Dim strDisDate As String
   Dim bContinue As Boolean
   Dim iIndex As Integer
   
   On Error GoTo DiaErr1
   
   If (strFullPath <> "") Then
      Set xlApp = New Excel.Application
   
      Set wb = xlApp.Workbooks.Open(strFullPath)
   
      Set ws = wb.Worksheets(1) 'Specify your worksheet name
      
      bContinue = True
      iIndex = 10
      While (bContinue)
'         strInvNum = ws.Cells(iIndex, 5)
'         If (IsNumeric(strInvNum)) Then
'            strInvNum = CStr(CDbl(strInvNum))
'         End If
'
'         strInvAmt = ws.Cells(iIndex, 9)
'         strInvDisc = ws.Cells(iIndex, 8)
'         strPONum = ws.Cells(iIndex, 10)
'         strChkNum = ws.Cells(iIndex, 11)
'         If (IsNumeric(strChkNum)) Then
'            strChkNum = CStr(CDbl(strChkNum))
'         End If
'         strRecDate = ws.Cells(iIndex, 6)
'         strInvStatus = ws.Cells(iIndex, 7)
'         strDuedate = ws.Cells(iIndex, 12)
'         strDisDate = ws.Cells(iIndex, 13)

' 2nd format
'
'         strInvNum = ws.Cells(iIndex, 5)
'         If (IsNumeric(strInvNum)) Then
'            strInvNum = CStr(CDbl(strInvNum))
'         End If
'
'         strInvAmt = ws.Cells(iIndex, 10)  ' 9
'         strInvDisc = ws.Cells(iIndex, 9) '8
'         strPONum = ws.Cells(iIndex, 12) '10
'         strChkNum = ws.Cells(iIndex, 14) '11
'         If (IsNumeric(strChkNum)) Then
'            strChkNum = CStr(CDbl(strChkNum))
'         End If
'         strRecDate = ws.Cells(iIndex, 6)
'         strInvStatus = ws.Cells(iIndex, 7)
'         strDuedate = ws.Cells(iIndex, 16) '12
'         strDisDate = ws.Cells(iIndex, 17) ' 13
         

         strInvNum = ws.Cells(iIndex, 13)
         If (IsNumeric(strInvNum)) Then
            strInvNum = CStr(CDbl(strInvNum))
         End If

         strInvAmt = ws.Cells(iIndex, 22)  ' 9
         strInvDisc = ws.Cells(iIndex, 23) '8
         strPONum = ws.Cells(iIndex, 10) '10
         strChkNum = ws.Cells(iIndex, 5) '11
         If (IsNumeric(strChkNum)) Then
            strChkNum = CStr(CDbl(strChkNum))
         End If
         If (Trim(ws.Cells(iIndex, 17)) <> "") Then
            strRecDate = ws.Cells(iIndex, 17)
         Else
            strRecDate = ""
         End If
         strInvStatus = ws.Cells(iIndex, 7)
         strDuedate = Trim(ws.Cells(iIndex, 6)) '12
         strDisDate = Trim(ws.Cells(iIndex, 6)) ' 13
         
         
         If (strInvNum <> "") And (strInvAmt <> "") Then
         
            sSql = "INSERT INTO ImpBIVSInv (Index_ID, INVNUMBER, INVAMOUNT,INVDISAMT, " _
                  & "PONUMBER,CHECKNUMBER,INVRECDATE,INVSTATUS,INVDUEDATE,INVDUEDISDATE) " _
               & "VALUES('" & CStr(iIndex) & "','" & strInvNum & "','" _
                     & strInvAmt & "','" & strInvDisc & "','" _
                     & strPONum & "','" & strChkNum & "','" _
                     & strRecDate & "','" & strInvStatus & "','" _
                     & strDuedate & "','" & strDisDate & "')"
            Debug.Print sSql
            
            clsADOCon.ExecuteSql sSql '
         
         End If
         
         
         If (strInvNum = "") And (strInvAmt = "") Then
            bContinue = False
         End If
         
         strInvNum = ""
         strInvAmt = ""
         strInvDisc = ""
         strPONum = ""
         strChkNum = ""
         strRecDate = ""
         strInvStatus = ""
         strDuedate = ""
         strDisDate = ""
         iIndex = iIndex + 1
      Wend
      
      wb.Close
   
      xlApp.Quit
      Set ws = Nothing
      Set wb = Nothing
      Set xlApp = Nothing
   End If
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
