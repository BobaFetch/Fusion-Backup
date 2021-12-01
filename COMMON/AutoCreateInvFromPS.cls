VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoCreateInvFromPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private RdoQry1 As rdoQuery
'Private RdoQry2 As rdoQuery
Dim cmdObj1 As ADODB.Command
Dim cmdObj2 As ADODB.Command

Private bOnLoad As Byte
Private bCancel As Byte
Private bGoodPs As Boolean
Private bInvUsed As Boolean
Private bGoodAct As Boolean
Private bInvPS As Boolean
Private sCusNickName As String
Private sCusName As String

Private iTotalItems As Integer
Private lNewInv As Long
Private cFREIGHT As Currency
Private cTax As Currency
Private cItm As Currency
Private cTot As Currency
Private sAccount As String
Private sInvPre As String
Private sPsStadr As String
Private sPsCust As String
Private lSo As Long
Private sMsg As String

' Sales Tax
Private sTaxCode As String
Private sTaxState As String
Private sTaxAccount As String
Private nTaxRate As Currency

' Sales journal
Private sCOSjARAcct As String
Private sCOSjINVAcct As String
Private sCOSjNFRTAcct As String
Private sCOSjTFRTAcct As String
Private sCOSjTaxAcct As String

Private sPackSlip  As String

Private vItems(300, 12) As Variant

Public Function Init(ByVal PackSlip As String)
   
   ' initialize the packslip
   sPackSlip = PackSlip
   
   ' check for open journal via posting date
   sJournalID = GetOpenJournal("SJ", Format(Now(), "mm/dd/yy"))
   If sJournalID = "" Then
      sMsg = "There Is No Open Journal For The Posting Date."
      MsgBox sMsg, vbInformation ',"Journals Open"
      Exit Function
   End If
   
   sAccount = GetSetting("Esi2000", "EsiFina", "LastRAccount", sAccount)
   sInvPre = GetSetting("Esi2000", "EsiFina", "LastInvPref", sInvPre)
   If Len(Trim(sInvPre)) = 0 Then sInvPre = "I"
   
   On Error Resume Next
   sSql = "SELECT DISTINCT PSNUMBER,PSCUST,PSTERMS,PSSTNAME,PSSTADR," _
          & "PSFREIGHT,CUREF,CUNICKNAME,CUNAME,PIPACKSLIP,PSPRIMARYSO,CUTYPE FROM PshdTable," _
          & "CustTable,PsitTable WHERE CUREF=PSCUST AND (PSTYPE=1 AND " _
          & "PSSHIPPRINT=1 AND PSINVOICE=0) AND PSNUMBER=PIPACKSLIP AND PSNUMBER= ? "
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   'RdoQry1.MaxRows = 1
   
   Set cmdObj1 = New ADODB.Command
   cmdObj1.CommandText = sSql
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   'RdoQry1.MaxRows = 1
   Dim prmObj1 As ADODB.Parameter
   Set prmObj1 = New ADODB.Parameter
   prmObj1.Type = adChar
   prmObj1.Size = 8
   cmdObj1.parameters.Append prmObj1
   
   
'   sSql = "SELECT PIPACKSLIP,PIQTY,PIPART,PISONUMBER,PISOITEM," _
'          & "PISOREV,PISELLPRICE,PARTREF,PARTNUM,PALEVEL,PAPRODCODE,PATAXEXEMPT," _
'          & "PASTDCOST FROM PsitTable,PartTable WHERE PIPART=" _
'          & "PARTREF AND PIPACKSLIP= ? ORDER BY PISONUMBER,PISOITEM"
          
   sSql = "SELECT PIPACKSLIP,PIQTY,PIPART,PISONUMBER,PISOITEM,PISOREV,PISELLPRICE," & vbCrLf _
          & "PARTREF,PARTNUM,PALEVEL,PAPRODCODE,PATAXEXEMPT,PASTDCOST,SOTAXABLE" & vbCrLf _
          & "FROM PsitTable" & vbCrLf _
          & "JOIN PartTable ON PIPART=PARTREF" & vbCrLf _
          & "JOIN  SohdTable on PISONUMBER = SONUMBER" & vbCrLf _
          & "WHERE PIPACKSLIP= ?" & vbCrLf _
          & "ORDER BY PISONUMBER,PISOITEM"

   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   
   Set cmdObj2 = New ADODB.Command
   cmdObj2.CommandText = sSql
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   'RdoQry1.MaxRows = 1
   Dim prmObj2 As ADODB.Parameter
   Set prmObj2 = New ADODB.Parameter
   prmObj2.Type = adChar
   prmObj2.Size = 8
   cmdObj2.parameters.Append prmObj2

End Function


Private Sub GetSJAccounts()
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   
   If sJournalID = "" Then
      bGoodAct = True
      Exit Sub
   End If
   
   On Error GoTo DiaErr1
   sSql = "SELECT COREF,COSJARACCT,COSJNFRTACCT," _
          & "COSJTFRTACCT,COSJTAXACCT FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         ' A/R
         sCOSjARAcct = "" & Trim(.Fields(1))
         If sCOSjARAcct = "" Then b = 1
         ' NonTaxable freight
         sCOSjNFRTAcct = "" & Trim(.Fields(2))
         If sCOSjNFRTAcct = "" Then b = 1
         ' Taxable freight
         sCOSjTFRTAcct = "" & Trim(.Fields(3))
         If sCOSjTFRTAcct = "" Then b = 1
         ' Sales tax
         sCOSjTaxAcct = "" & Trim(.Fields(4))
         If sCOSjTaxAcct = "" Then b = 1
         .Cancel
      End With
   End If
   If b = 1 Then
      bGoodAct = False
   Else
      bGoodAct = True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getsjacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillPackslipDetail()
   bGoodPs = GetPackslip()
   If bGoodPs Then
      GetPSItems
   End If
End Sub

Public Function AddNewInvoice()
   
   If (sPackSlip <> "") Then
      
      GetSJAccounts
      If bGoodAct = False Then
         MsgBox "One Or More Journal Accounts Are Not Registered." & vbCr _
            & "Please Install All Accounts In the Company Setup.", _
            vbInformation ', "Journal Account"
         Exit Function
      End If
      
      CheckInvAsPS
      FillPackslipDetail
      
      If bGoodPs Then
         
         GetNextInvoice
         If (lNewInv > 0) Then
            AddInvoice (sPackSlip)
            UpdateInvoice lNewInv
         End If
      Else
         MsgBox "Not a valid Packslip." & vbCr _
            & "Couldn't create a new Invoice for the Packslip.", vbExclamation
         Exit Function
      End If
      
   End If
End Function


Public Function GetPackslip() As Boolean
   Dim RdoPsl As ADODB.Recordset
   
   On Error GoTo DiaErr1
   Erase vItems
   sTaxCode = ""
   sTaxState = ""
   sTaxAccount = ""
   nTaxRate = 0
   'RdoQry1(0) = Trim(sPackSlip)
   'bSqlRows = GetQuerySet(RdoPsl, RdoQry1)
   cmdObj1.parameters(0).Value = sPackSlip
   bSqlRows = clsADOCon.GetQuerySet(RdoPsl, cmdObj1, ES_KEYSET, True)
   If bSqlRows Then
      With RdoPsl
         'cmbPsl = "" & Trim(!PsNumber)
         sCusNickName = "" & Trim(!CUNICKNAME)
         sCusName = "" & Trim(!CUNAME)
         sPsCust = "" & Trim(!CUREF)
         sPsStadr = "" & Trim(!PSSTNAME) & vbCrLf _
                    & Trim(!PSSTADR)
         cFREIGHT = Format(!PSFREIGHT, "#####0.00")
         lSo = !PSPRIMARYSO
         .Cancel
      End With
      
      GetSalesTaxInfo Compress(sCusName), nTaxRate, sTaxCode, sTaxState, sTaxAccount
      sPsStadr = CheckComments(sPsStadr)
      GetPackslip = True
   Else
      
      cFREIGHT = 0
      GetPackslip = False
   End If
   Set RdoPsl = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpackslip"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   DoModuleErrors Me
   
End Function

Public Function GetNextInvoice()
   Dim inv As ClassARInvoice
   Dim bDup As Boolean
   
   Set inv = New ClassARInvoice
   lNewInv = inv.GetNextInvoiceNumber
   
   If (bInvPS = True) Then
      Dim strPSNum As String
      strPSNum = Mid$(CStr(sPackSlip), 3, Len(sPackSlip))
      If (strPSNum <> "") Then
         lNewInv = Val(strPSNum)
      End If
   
   End If
   
   ' Validate the Invoice number
   If (lNewInv > 0) Then
      bDup = inv.DuplicateInvNumber(lNewInv)
      
      If (bDup = True) Then
         MsgBox "Invoice number exists.", vbInformation, "Invoice Exists"
         Exit Function
      End If
   End If
   
   
End Function

Public Sub AddInvoice(sPS As String)
   Dim RdoInvc As ADODB.Recordset
   On Error GoTo DiaErr1
   ' Reserve a record in case the invoice number is changed.
   ' Use TM so that it won't show and can be safely deleted.
   
   sSql = "SELECT * FROM CihdTable WHERE INVNO = " & Val(lNewInv) & " AND INVTYPE = 'TM'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInvc, ES_KEYSET)
   If Not bSqlRows Then
      sSql = "INSERT INTO CihdTable (INVNO,INVTYPE,INVSO,INVCANCELED) " _
             & "VALUES(" & lNewInv & ",'TM'," _
             & Val(sPS) & ",0)"
      clsADOCon.ExecuteSql sSql ',  rdExecDirect
   End If
   Set RdoInvc = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "addinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Public Sub UpdateInvoice(lCurrInvoice As Long)
   Dim bByte As Byte
   Dim bResponse As Byte
   
   Dim a As Integer
   Dim i As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   
   Dim nLDollars As Currency
   Dim nTdollars As Currency
   
   Dim sPart As String
   Dim iLevel As Integer
   Dim cCost As Currency
   Dim sProd As String
   
   ' Accounts most of which are not used in this transaction yet.
   Dim sREVAccount As String
   Dim sCGSAccount As String
   
   ' BnO Taxes
   Dim nRate As Currency
   Dim sType As String
   Dim sState As String
   Dim sCode As String
   
   Dim sPost As String
   
   Dim sTemp As String
   
   ' Make sure user really wants to proceed.
   sMsg = "Post The Selected Invoice With Packing Slip?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, "Invoice with Packslip")
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If
   
   ' Look For Accounts ?
   If sJournalID <> "" Then
      iTrans = GetNextTransaction(sJournalID)
   End If
   If iTrans > 0 Then
      bByte = True
      For i = 1 To iTotalItems
         If Val(vItems(i, 8)) > 0 Then
            sPart = vItems(i, 3)
            iLevel = Val(vItems(i, 7))
            sProd = vItems(i, 8)
            
         
            bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
                    sREVAccount, sCGSAccount)
         End If
         If bByte = False Then Exit For
      Next
   End If
   
   sPost = Format(Now(), "mm/dd/yyyy")
   
   On Error Resume Next
   'RdoCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "UPDATE PshdTable SET PSINVOICE=" & lCurrInvoice & " " _
          & "WHERE PSNUMBER='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql
   
   ' Update lot record
   sSql = "UPDATE LoitTable SET LOICUSTINVNO=" & lCurrInvoice _
          & ",LOICUST='" & sPsCust & "' WHERE LOIPSNUMBER = '" & sPackSlip & "'"
   clsADOCon.ExecuteSql sSql
   
   For i = 1 To iTotalItems
      
      ' Running invoice total
      nTdollars = nTdollars + (Val(vItems(i, 4)) * Val(vItems(i, 11)))
      
      ' Part accounts
      sPart = vItems(i, 3)
      sProd = vItems(i, 8)
      iLevel = Val(vItems(i, 7))
      
      bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
              sREVAccount, sCGSAccount)
      
      ' BnO tax
      sCode = ""
      nRate = 0
      sState = ""
      sType = ""
      GetPartBnO vItems(i, 3), nRate, sCode, sState, sType
      If sCode = "" Then
         GetCustBnO sPsCust, nRate, sCode, sState, sType
      End If
      
      ' Update the sales order item
      sSql = "UPDATE SoitTable SET " _
             & "ITINVOICE=" & lCurrInvoice & "," _
             & "ITREVACCT='" & sREVAccount & "'," _
             & "ITCGSACCT='" & sCOSjARAcct & "'," _
             & "ITBOSTATE='" & sState & "'," _
             & "ITBOCODE='" & sCode & "'," _
             & "ITSLSTXACCT='" & sTaxAccount & "'," _
             & "ITTAXCODE='" & sTaxCode & "'," _
             & "ITSTATE='" & sTaxState & "'," _
             & "ITTAXRATE=" & nTaxRate & "," _
             & "ITTAXAMT=" & CCur((nTaxRate / 100) * (Val(vItems(i, 4)) * Val(vItems(i, 11)))) & " " _
             & "WHERE ITSO=" & Val(vItems(i, 0)) & " AND " _
             & "ITNUMBER=" & Val(vItems(i, 1)) & " AND " _
             & "ITREV='" & vItems(i, 2) & "' "
      clsADOCon.ExecuteSql sSql
      
      'Journal entries
      nLDollars = (Val(vItems(i, 4)) * Val(vItems(i, 11)))
      cCost = (Val(vItems(i, 4)) * Val(vItems(i, 9)))
      
      
      Dim gl As New GLTransaction
      gl.JournalID = Trim(sJournalID)
      gl.InvoiceDate = CDate(Now())
      gl.InvoiceNumber = lCurrInvoice
      
      gl.AddDebitCredit CCur(nLDollars), 0, Compress(sCOSjARAcct), sPart, _
                             CLng(vItems(i, 0)), CInt(vItems(i, 1)), CStr(vItems(i, 2)), sPsCust, "", True
      ' Credit Revenue (-)
      gl.AddDebitCredit 0, CCur(nLDollars), sREVAccount, _
                        sPart, CLng(vItems(i, 0)), CInt(vItems(i, 1)), CStr(vItems(i, 2)), sPsCust, "", True
      
      Set gl = Nothing
   Next
   
   If cFREIGHT > 0 Then
      
      ' Debit A/R Freight
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cFREIGHT & ",'" _
             & sCOSjARAcct & "','" _
             & sPsCust & "','" _
             & Format(Now(), "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
      
      ' Credit Freight
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cFREIGHT & ",'" _
             & sCOSjNFRTAcct & "','" _
             & sPsCust & "','" _
             & Format(Now(), "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
   End If
   
   If cTax > 0 Then
      ' Debit A/R Taxes
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cTax & ",'" _
             & sCOSjARAcct & "','" _
             & sPsCust & "','" _
             & Format(Now(), "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
      
      ' Credit Taxes
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cTax & ",'" _
             & sCOSjTaxAcct & "','" _
             & sPsCust & "','" _
             & Format(Now(), "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
   End If
   
   ' Change TM invoice to PS
   sSql = "UPDATE CihdTable SET INVNO=" & lCurrInvoice & "," _
          & "INVPRE='" & sInvPre & "',INVSTADR='" & sPsStadr & "'," _
          & "INVTYPE='PS',INVSO=0," _
          & "INVCUST='" & sPsCust & "' WHERE " _
          & "INVNO=" & lNewInv & " AND INVTYPE='TM'"
   clsADOCon.ExecuteSql sSql
   
   If (bInvPS = False) Then
      Dim inv As New ClassARInvoice
      inv.SaveLastInvoiceNumber lNewInv
   End If
   ' Add freight and tax to invoice total
   nTdollars = nTdollars + (cTax + cFREIGHT)
   
   'MM added CANCELED flag to 0
   ' Then post the total to the invoice
   sSql = "UPDATE CihdTable SET INVTOTAL=" & nTdollars & "," _
          & "INVFREIGHT=" & cFREIGHT & "," _
          & "INVTAX=" & cTax & "," _
          & "INVSHIPDATE='" & sPost & "'," _
          & "INVDATE='" & sPost & "'," _
          & "INVCANCELED=0," _
          & "INVCOMMENTS='" & sPsStadr & "'," _
          & "INVPACKSLIP='" & sPackSlip & "' " _
          & "WHERE INVNO=" & lCurrInvoice & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE PshdTable SET PSFREIGHT=" & cFREIGHT _
          & " WHERE PSNUMBER='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql
   
   MouseCursor 0
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      sMsg = "Successfully Posted Customer Invoice number - " & CStr(lCurrInvoice)
      MsgBox sMsg, vbInformation, "Error"
   Else
      clsADOCon.RollbackTrans
      sMsg = "Couldn't Post The Invoice."
      MsgBox sMsg, vbInformation, "Error"
      Exit Sub
   End If
   
   ' We don't want to reset the date
   'txtDte = Format(GetServerDateTime, "mm/dd/yy")
   Erase vItems
   Exit Sub
   
DiaErr1:
   sProcName = "updateinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub GetPSItems()
   Dim i As Integer
   Dim RdoPsi As ADODB.Recordset
   'Dim cTax As Currency
   
   MouseCursor 13
   On Error Resume Next
   cTax = 0
   
   
   ' Update the packslip dollars from sales order
   sSql = "UPDATE PsitTable SET PISELLPRICE=" _
          & "ITDOLLARS FROM PsitTable,SoitTable WHERE " _
          & "(PISONUMBER=ITSO AND PISOITEM=ITNUMBER " _
          & "AND PISOREV=ITREV) AND PIPACKSLIP='" & Trim(sPackSlip) & "'"
   clsADOCon.ExecuteSql sSql ',  rdExecDirect
   
   On Error GoTo DiaErr1
   
   ' Return packslip items
   'RdoQry2(0) = Trim(sPackSlip)
   'bSqlRows = GetQuerySet(RdoPsi, RdoQry2)
   cmdObj2.parameters(0).Value = sPackSlip
   bSqlRows = clsADOCon.GetQuerySet(RdoPsi, cmdObj2, ES_KEYSET, True)
   
   
   If bSqlRows Then
      With RdoPsi
         Do Until .EOF
            i = i + 1
            vItems(i, 0) = Format(!PISONUMBER, SO_NUM_FORMAT)
            vItems(i, 1) = Format(!PISOITEM, "##0")
            vItems(i, 2) = Trim(!PISOREV)
            vItems(i, 3) = "" & Trim(!PartRef)
            'vItems(i, 4) = Format(!PIQTY, "#####0.000")
            vItems(i, 4) = !PIQTY
            vItems(i, 5) = "" & sAccount
            vItems(i, 6) = "" & !PIPART
            vItems(i, 7) = "" & !PALEVEL
            vItems(i, 8) = "" & Trim(!PAPRODCODE)
            'vItems(i, 9) = Format(!PASTDCOST, "#####0.000")
            vItems(i, 9) = !PASTDCOST
            vItems(i, 10) = 1
            'vItems(i, 11) = Format(!PISELLPRICE, "#####0.000")
            vItems(i, 11) = !PISELLPRICE
            vItems(i, 12) = !PATAXEXEMPT
            If vItems(i, 12) = 0 And !SOTAXABLE = 1 Then
               cTax = cTax + ((vItems(i, 11) * vItems(i, 4)) * (nTaxRate / 100))
            End If
            cItm = cItm + (vItems(i, 11) * vItems(i, 4))
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoPsi = Nothing
   iTotalItems = i
   cTot = CCur(cItm) + CCur(cTax) + CCur(cFREIGHT)
   'UpdateTotals
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getpsitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   DoModuleErrors Me
End Sub

'Private Sub UpdateTotals()
'   Dim cTot As Currency
'   cTot = CCur(lblItm) + CCur(txtTax) + CCur(txtFrt)
'   'lblTot = Format(cTot, CURRENCYMASK)
'End Sub

Private Function DeleteOldTmpInv(strPreInv As String)
   On Error Resume Next
   'dump any left over dummies
   sSql = "DELETE FROM CihdTable WHERE INVNO=" & Val(strPreInv) _
          & " AND INVTYPE='TM'"
   clsADOCon.ExecuteSql sSql ',  rdExecDirect
End Function

Private Function CheckInvAsPS()
   Dim RdoInv As ADODB.Recordset
   
   sSql = "SELECT * FROM ComnTable WHERE COALLOWINVNUMPS = 1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_KEYSET)
   If bSqlRows Then
      bInvPS = True
      ClearResultSet RdoInv
   Else
      'chkInvPS.enabled = False
      bInvPS = False
   End If
   Set RdoInv = Nothing
   
End Function

Private Function GetSalesTaxInfo(sCust As String, _
                           nRate As Currency, _
                           sCode As String, _
                           sState As String, _
                           sAccount As String)
   
   On Error GoTo DiaErr1
   
   ' Load tax from customer.
   Dim RdoTax As ADODB.Recordset
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE,TAXACCT FROM CustTable INNER JOIN " _
          & "TxcdTable ON CustTable.CUTAXCODE = TxcdTable.TAXREF " _
          & "WHERE CUREF = '" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTax)
   If bSqlRows Then
      With RdoTax
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sAccount = "" & Trim(!TAXACCT)
         .Cancel
      End With
   End If
   Set RdoTax = Nothing
   Exit Function
DiaErr1:
   sProcName = "getsaletaxinfo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Private Sub GetCustBnO(sCust, nRate, sCode, sState, sType)
   ' Get B&O tax codes from customer
   ' Retail takes precidence over wholesale
   
   Dim rdoTx1 As ADODB.Recordset
   Dim rdoTx2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,CustTable " _
          & "WHERE CUBORTAXCODE = TAXREF AND CUREF = '" & sCust _
          & "' AND TAXTYPE = 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx1)
   If bSqlRows Then
      With rdoTx1
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sType = "R"
         .Cancel
      End With
   Else
      sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,CustTable " _
             & "WHERE CUBORTAXCODE = TAXREF AND CUREF = '" & sCust _
             & "' AND TAXTYPE = 0"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx2)
      If bSqlRows Then
         With rdoTx2
            nRate = !TAXRATE
            sCode = "" & Trim(!taxCode)
            sState = "" & Trim(!taxState)
            sType = "W"
            .Cancel
         End With
      End If
   End If
   
   Set rdoTx1 = Nothing
   Set rdoTx2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcustbno"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Private Sub GetPartBnO(sPart, nRate, sCode, sState, sType)
   
   Dim rdoTx1 As ADODB.Recordset
   Dim rdoTx2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,PartTable WHERE " _
          & "PABORTAX = TAXREF AND TAXTYPE = 0 AND PARTREF = '" & sPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx1)
   If bSqlRows Then
      With rdoTx1
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sType = "R"
         .Cancel
      End With
   Else
      sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,PartTable WHERE " _
             & "PABOWTAX = TAXREF AND TAXTYPE = 0 AND PARTREF = '" & sPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx2)
      If bSqlRows Then
         With rdoTx2
            nRate = !TAXRATE
            sCode = "" & Trim(!taxCode)
            sState = "" & Trim(!taxState)
            sType = "W"
            .Cancel
         End With
      End If
   End If
   
   Set rdoTx1 = Nothing
   Set rdoTx2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpartbno"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub

