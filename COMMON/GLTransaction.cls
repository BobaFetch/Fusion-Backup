VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GLTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sJournalID As String
Private iTrans As Integer
Private iNextRef As Integer 'reference number = 1...
Private nInvoiceNumber As Long
Private dtInvoiceDate As Date
Private debits As Currency
Private credits As Currency
Private message As String
Private trace As String

'Sales journal accounts
Private sCOSjARAcct As String
Private sCOSjINVAcct As String
Private sCOSjNFRTAcct As String
Private sCOSjTFRTAcct As String
Private sCOSjTaxAcct As String
Private sCOSjFedTaxAcct As String

Public Enum eDebCredType
   DCTYPE_AP_ChkAPAcct = 11      'debit to AP when check cut, or credit when voided
   DCTYPE_AP_ChkChkAcct = 12     'credit to checking account when check cut, or debit when voided
   DCTYPE_AP_ChkDiscAcct = 13    'credit to AP discount when check cut, or debit when voided
End Enum

Public Property Get SJARAccount() As String
   If sCOSjARAcct = "" Then
      GetSJAccounts
   End If
   SJARAccount = sCOSjARAcct
End Property

Public Property Get SJTaxAccount() As String
   If sCOSjTaxAcct = "" Then
      GetSJAccounts
   End If
   SJTaxAccount = sCOSjTaxAcct
End Property

Public Property Get SJFedTaxAccount() As String
   If sCOSjFedTaxAcct = "" Then
      GetSJAccounts
   End If
   SJFedTaxAccount = sCOSjFedTaxAcct
End Property

Public Property Get SJFreightAccount() As String
   If sCOSjNFRTAcct = "" Then
      GetSJAccounts
   End If
   SJFreightAccount = sCOSjNFRTAcct
End Property

Public Property Get JournalID() As String
   JournalID = sJournalID
End Property

Public Property Let JournalID(NewValue As String)
   sJournalID = Trim(NewValue)
   iTrans = GetNextTransaction(sJournalID)
End Property

Public Property Get TransactionNumber() As Integer
   TransactionNumber = iTrans
End Property

Public Property Get InvoiceNumber() As Long
   InvoiceNumber = nInvoiceNumber
End Property

Public Property Let InvoiceNumber(NewValue As Long)
   nInvoiceNumber = NewValue
End Property

Public Property Get InvoiceDate() As Date
   InvoiceDate = dtInvoiceDate
End Property

Public Property Let InvoiceDate(NewValue As Date)
   dtInvoiceDate = NewValue
End Property

Private Sub Class_Initialize()
   iNextRef = 1
End Sub

Public Sub AddCashReceipt(sCheckNo As String, sCust As String, _
                          bCaType As Byte, ReceiptDate As String, CheckAmt As Currency, _
                          AppliedAmt As Currency, DiscountAmt As Currency, _
                          CommissionAmt As Currency, ExpenseAmt As Currency, _
                          RevenueAmt As Currency, _
                          CheckDate As String, CashAcct As String)
   
   'CheckAmt       = CACKAMT = total cash receipt (multiple invoices)
   'AppliedAmt     = CACRAMT = amount applied to invoice if invoiced
   '                 or uninvoiced amount if no invoice #
   'DiscountAmt    = CADISC = discount amount this invoice
   'CommissionAmt  = CACOMM = commission amount
   'ExpenseAmt     = CAEXP = misc expense
   'RevenueAmt     = CAREV = misc revenue
   
   ' Cash Recipt Types (CATYPE)
   ' 0 = Cash
   ' 1 = Wire / Credit Card
   ' 2 = Check
   
   'Invoice number = gl.InvoiceNumber
   
   'MCS Dictionary
   '    INTEGER CACODE          1       ;RECORD TYPE CODE 1=START OF A DEP-BATCH
   '                                    ;                 2=START OF NEXT PAYMENT
   '                                    ;                 3=NEXT INVOICE OF A PAYMENT
   '    POINTER CANXINV         2       ;POINTER TO NEXT C/R RECORD ON INVOICE CHAIN
   '    POINTER CANXBATCH       2       ;POINTER TO NEXT C/R RECORD ON BATCH CHAIN
   '    POINTER CABPTR          2       ;POINTER TO BATCH RECORD
   '    POINTER CAINVPTR        2       ;POINTER TO INVOICE
   '    POINTER CACUSTPTR       2       ;POINTER TO CUSTOMER RECORD
   '    MONEY   CAAMOUNT        2       ;AMOUNT OF CASH REC FOR THIS TYPE & THIS INVOIC
   '    INTEGER CATYPE          1       ;TYPE OF PAYMENT 1= CASH        3=BANK XFER
   '                                    ;                2= CHECK       4=OTHER
   '    DATE    CARCDATE        2       ;DATE RECEIVED
   '    ALPHA   CACKNO          3       ;6 CHARACTER CHECK NUMBER
   '    MONEY   CACKAMT         2       ;TOTAL CHECK/PAYMENT AMOUNT(COULD BE SEV.INV'S)
   '    DATE    CACKDATE        2       ;CHECK DATE OR DATE OF XFER
   '    POINTER CAENTRY         2       ;POINTER TO EMPL REC OF OPERATOR MAKING ENTRY
   '    MONEY   CADISC          2       ;DISCOUNT AMOUNT
   '    MONEY   CACOMM          2       ;COMMISSION AMOUNT
   '    MONEY   CAEXP           2       ;MISC. EXPENSE
   '    MONEY   CAREV           2       ;MISC. REVENUE
   '    MONEY   CAWOFF          2       ;WRITEOFF AMOUNT
   '    SPARE   OLDCAACCOUNT    4       ;GL ACCOUNT NUMBER   (AS OF 12/30/86)
   '    DATE    CACDATE         2       ;CREATION DATE OF THIS RECORD
   '    YESNO   CAPOSTED        1       ;C/R RECORD POSTED FLAG
   '    INTEGER CADPB           1       ;CASH REC. DEPOSIT-BATCH NUMBER
   '    MONEY   CACRAMT         2       ;CREDIT AMT FOR "OTHER" TYPE 4 RECORDS
   '    ALPHA   CAACCOUNT       6       ;GL ACCOUNT NUMBER
   '    SPARE                   13      ;NOT YET USED
   
   
   On Error Resume Next
   
   Debug.Print "Cash Chk=" & CheckAmt & " Applied=" & AppliedAmt _
      & " Disc=" & DiscountAmt & " Commis=" & CommissionAmt _
      & " Exp=" & ExpenseAmt & " Rev=" & RevenueAmt
   
   sSql = "INSERT INTO CashTable (" _
          & "CAINVNO,CACHECKNO,CACUST,CATYPE," _
          & "CARCDATE,CACKAMT,CACRAMT,CADISC,CACOMM,CAEXP,CAREV,CACDATE," _
          & "CACASHACCT,CAENTRY) VALUES(" _
          & CStr(nInvoiceNumber) & ",'" _
          & sCheckNo & "','" _
          & sCust & "'," _
          & bCaType & ",'" _
          & Format(ReceiptDate, "mm/dd/yyyy") & "'," _
          & CheckAmt & "," _
          & AppliedAmt & "," _
          & DiscountAmt & "," _
          & CommissionAmt & "," _
          & ExpenseAmt & "," _
          & RevenueAmt & ",'" _
          & CheckDate & "','" _
          & CashAcct & "','" _
          & Secure.UserInitials & "')"
   
   Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql
   If Err Then
      Debug.Print Err.Description
   End If
End Sub

Public Function AddDebitCredit(debit As Currency, credit As Currency, acctno As String, _
                          PartNo As String, sono As Long, itno As Integer, itrev As String, cust As String, _
                          Optional sCheckNo As String = "", Optional bIgnrZeroDdCr As Boolean = False) As Boolean
   
   'returns True if successful
   Dim success As Boolean
   success = True
   AddDebitCredit = True
   
   'if debit and credit are both zero, don't create an entry
   If debit = 0 And credit = 0 And bIgnrZeroDdCr = False Then
      Exit Function
   End If
   
   'if debit or credit is negative, reverse debit with credit
   Dim deb As Currency, cred As Currency
   If debit < 0 Or credit < 0 Then
      deb = -credit
      cred = -debit
   Else
      deb = debit
      cred = credit
   End If
   
   'add a debit or a credit to a journal
   'rdoCon is used so that you may include any GL entries created here
   'in a larger transaction.
   
   On Error Resume Next
   
   '    trace = "-" & iNextRef & ": " _
   '        & FixedLeft(acctno, 8) & " " & FixedRight(deb, 12) & " " & FixedRight(cred, 12) _
   '        & "  inv#=" & nInvoiceNumber
   
   trace = "-" & iNextRef & ": " _
           & acctno & " " & Chr(9) & deb & " " & Chr(9) & cred _
           & " " & Chr(9) & "inv#=" & nInvoiceNumber
   
   'takes too many characters - 1024 msgbox limit
   '    If Debugging Then
   '        trace = trace & " " & GetAccountName(acctno)
   '    End If
   
   Debug.Print trace
   message = message & trace & vbCrLf
   
   sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT,DCCREDIT," _
          & "DCACCTNO,DCPARTNO,DCSONUMBER,DCSOITNUMBER,DCSOITREV," _
          & "DCCUST,DCDATE,DCINVNO,DCCHECKNO) " & vbCrLf _
          & "VALUES('" & sJournalID & "'," _
          & iTrans & "," _
          & iNextRef & "," _
          & deb & "," _
          & cred & ",'" _
          & Compress((acctno)) & "','" _
          & Trim(PartNo) & "'," _
          & sono & "," _
          & itno & ",'" _
          & itrev & "','" _
          & cust & "','" _
          & Format(dtInvoiceDate, "mm/dd/yyyy") & "'," _
          & CStr(nInvoiceNumber) & ",'" _
          & sCheckNo & "'" _
          & ")"
   iNextRef = iNextRef + 1
   success = clsADOCon.ExecuteSql(sSql)
   
   If success Then
      AddDebitCredit = True
      debits = debits + deb
      credits = credits + cred
  Else
      AddDebitCredit = False
      'sProcName = "GlTransaction.AddDebitCredit"
      'CurrError.Number = Err.Number
      'CurrError.Description = Err.Description
      'Dim frm As New ClassErrorForm
      'DoModuleErrors frm
  End If
End Function

Public Function Commit()
   'We're not actually doing a single transaction yet,
   'but we can check that the debits and credits balance for now
   
   If debits <> credits Then
      trace = sJournalID & "-" & iTrans _
              & " OUT OF BALANCE debits= " & debits & " credits=" & credits
      message = message & vbCrLf & trace & vbCrLf
      MsgBox message
      Commit = False
      
   Else
      trace = sJournalID & "-" & iTrans _
              & " IN BALANCE debits= " & debits & " credits=" & credits
      message = message & vbCrLf & trace & vbCrLf
      If Debugging Then
         Select Case MsgBox(message & "Click Cancel to cancel the transaction", vbOKCancel, "Check GL Transaction Balance")
            Case vbCancel
               Commit = False
               Exit Function
         End Select
      End If
      Commit = True
   End If
End Function

Public Sub ShowDebitsAndCredits()
   MsgBox message
End Sub

Public Function GetAccountName(acctno As String) As String
   Dim rdoGL As ADODB.Recordset
   
   sSql = "select GLDESCR from GlacTable where GLACCTREF = '" & acctno & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoGL, ES_FORWARD)
   If bSqlRows Then
      GetAccountName = Trim(rdoGL.Fields(0))
   Else
      GetAccountName = "UNDEFINED"
   End If
   Set rdoGL = Nothing
End Function

Public Function IsValidRevenueAccount(sAcct As String)
   Dim rdo As ADODB.Recordset
   On Error GoTo whoops
   Dim S As String
   IsValidRevenueAccount = False
   S = Compress(sAcct)
   '    sSql = "select * from GlacTable where GLACCTREF = '" & s & "'" & vbCrLf _
   '        & "and GLACCTREF not in (select GLMASTER from GlacTable)" & vbCrLf _
   '        & "and GLTYPE = 4"
   sSql = "select * from GlacTable where GLACCTREF = '" & S & "'" & vbCrLf _
          & "AND GlacTable.GLINACTIVE = 0 and GLACCTREF not in (select GLMASTER from GlacTable)"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   Set rdo = Nothing
   If bSqlRows Then
      IsValidRevenueAccount = True
   Else
      Dim sMsg As String
      sMsg = sAcct & " is not a valid revenue account"
      MsgBox sMsg, vbInformation, "Invalid revenue account"
   End If
   Exit Function
   
whoops:
   sProcName = "IsValidRevenueAccount"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Function


'See if the Accounts are there

Public Sub GetSJAccounts()
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   
   '    On Error GoTo DiaErr1
   On Error Resume Next
   sSql = "SELECT COREF,COSJARACCT,COSJINVACCT,COSJNFRTACCT," _
          & "COSJTFRTACCT,COSJTAXACCT, COFEDTAXACCT FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         sCOSjARAcct = "" & Trim(.Fields(1))
         sCOSjINVAcct = "" & Trim(.Fields(2))
         sCOSjNFRTAcct = "" & Trim(.Fields(3))
         sCOSjTFRTAcct = "" & Trim(.Fields(4))
         sCOSjTaxAcct = "" & Trim(.Fields(5))
         sCOSjFedTaxAcct = "" & Trim(.Fields(6))
         .Cancel
      End With
   End If
   Set rdoJrn = Nothing
   '    Exit Sub
   '
   'DiaErr1:
   '    sProcName = "getsjacco"
   '    CurrError.Number = Err.Number
   '    CurrError.description = Err.description
   '    DoModuleErrors Me
   
End Sub

Public Sub FillComboWithAccounts(cbo As ComboBox, _
                                 Optional cbo2 As ComboBox = Nothing, _
                                 Optional cbo3 As ComboBox = Nothing, _
                                 Optional cbo4 As ComboBox = Nothing, _
                                 Optional cbo5 As ComboBox = Nothing, _
                                 Optional cbo6 As ComboBox = Nothing, _
                                 Optional cbo7 As ComboBox = Nothing, _
                                 Optional cbo8 As ComboBox = Nothing, _
                                 Optional cbo9 As ComboBox = Nothing, _
                                 Optional cbo10 As ComboBox = Nothing _
                                 )
   
   Dim rdo As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT GLACCTNO FROM GlacTable" & vbCrLf _
          & "where GLACCTREF NOT IN (select GLMASTER from GlacTable)" & vbCrLf _
          & " AND GlacTable.GLINACTIVE = 0 " & vbCrLf _
          & "order by GLACCTNO"
   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   While Not rdo.EOF
      cbo.AddItem Trim(rdo!GLACCTNO)
      If Not cbo2 Is Nothing Then cbo2.AddItem Trim(rdo!GLACCTNO)
      If Not cbo3 Is Nothing Then cbo3.AddItem Trim(rdo!GLACCTNO)
      If Not cbo4 Is Nothing Then cbo4.AddItem Trim(rdo!GLACCTNO)
      If Not cbo5 Is Nothing Then cbo5.AddItem Trim(rdo!GLACCTNO)
      If Not cbo6 Is Nothing Then cbo6.AddItem Trim(rdo!GLACCTNO)
      If Not cbo7 Is Nothing Then cbo7.AddItem Trim(rdo!GLACCTNO)
      If Not cbo8 Is Nothing Then cbo8.AddItem Trim(rdo!GLACCTNO)
      If Not cbo9 Is Nothing Then cbo9.AddItem Trim(rdo!GLACCTNO)
      If Not cbo10 Is Nothing Then cbo10.AddItem Trim(rdo!GLACCTNO)
      rdo.MoveNext
   Wend
   cbo.ListIndex = 0
   Set rdo = Nothing
   Exit Sub
DiaErr1:
   sProcName = "FillComboWithAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Sub

Private Function FixedLeft(inputstring As String, desiredSize As Integer) As String
   If desiredSize > Len(Trim(inputstring)) Then
      FixedLeft = Trim(inputstring) & Space(desiredSize - Len(Trim(inputstring)))
   Else
      FixedLeft = inputstring
   End If
End Function

Private Function FixedRight(Amount As Currency, desiredSize As Integer) As String
   Dim inputstring As String
   inputstring = Format(Amount, "###,###.00")
   If desiredSize > Len(inputstring) Then
      FixedRight = Space(desiredSize - Len(inputstring)) & inputstring
   Else
      FixedRight = inputstring
   End If
End Function
