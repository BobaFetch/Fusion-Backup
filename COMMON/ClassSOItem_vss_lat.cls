VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Sub UpdateCommission(SoNumber As Long, SoItem As Integer, SoRev As String, Salesman As String)
'
'   'if item not comissionable, remove commissions
'   Dim rdo As rdoResultset
'   Dim part As String, commissionable As Boolean, extPrice As Currency, percent As Currency
'   Dim commission As Currency
'   sSql = "select PACOMMISSION, PARTREF," & vbCrLf _
'      & "cast(ITQTY * ITDOLLARS as decimal(12,2))as ExtendedPrice" & vbCrLf _
'      & "from PartTable" & vbCrLf _
'      & "join SoItTable on ITPART = PARTREF" & vbCrLf _
'      & "where ITSO = " & SoNumber & vbCrLf _
'      & "and ITNUMBER = " & SoItem & vbCrLf _
'      & "and ITREV = '" & SoRev & "'"
'   If GetDataSet(rdo) Then
'      part = rdo!PartRef
'      commissionable = IIf(rdo!PACOMMISSION = 1, True, False)
'      extPrice = rdo!ExtendedPrice
'   End If
'
'   'if the part is not commissionable, delete any commssions for this item
'   If Not commissionable Then
'      sSql = "delete from SpcoTable" & vbCrLf _
'         & "where SMCOSO = " & SoNumber & vbCrLf _
'         & "and SMCOSOIT = " & SoItem & " and SMCOITREV = '" & SoRev & "'"
'      RdoCon.Execute sSql
'      Exit Sub
'
'   Else
'      'calculate commission rate for this item for this salesman
'      'for now, all customers just use a fixed rate.
'      sSql = "select SPPERC1 as Rate from SprsTable" & vbCrLf _
'         & "where SPNUMBER = '" & Salesman & "'"
'      If GetDataSet(rdo) Then
'         percent = rdo!Rate
'      End If
'   End If
'
'   'attempt to insert a commission item
'   'if it already exists, we will update it
'   commission = percent * extPrice / 100
'
'   On Error Resume Next
'   sSql = "INSERT INTO SpcoTable (SMCOSO,SMCOSOIT,SMCOITREV," & vbCrLf _
'      & "SMCOSM,SMCOUSER,SMCOAMT,SMCOPCT)" & vbCrLf _
'      & "VALUES (" & SoNumber & "," & SoItem _
'      & ",'" & SoRev & "','" & Salesman & "','" _
'      & sInitials & "'," & commission & "," & percent & ")"
'   RdoCon.Execute sSql, rdExecDirect
'   'On Error GoTo 0
'
'   'if insert failed, record already exists.  update it.
'   If Err > 0 Then
'      On Error GoTo 0
'      Err.Clear
'      sSql = "UPDATE SpcoTable SET SMCOAMT = " & commission _
'         & ",SMCOPCT = " & percent & vbCrLf _
'         & " WHERE SMCOSO = " & SoNumber _
'         & " AND SMCOSOIT = " & SoItem _
'         & " AND SMCOITREV = '" & SoRev _
'         & "' AND SMCOSM = '" & Salesman & "'"
'      RdoCon.Execute sSql, rdExecDirect
'   End If
'End Sub
'

Public Sub InsertCommission(SoNumber As Long, SoItem As Integer, SoRev As String, Salesman As String)
   'insert a new commission record for a sales order item
   'if a record already exists for that salesman for that item, no insertion occurs
   'do this as part of a transaction, inserting the new commission record(s) first with InsertCommission()
   'and then updating the commissions for the item with UpdateCommissions.
   'it's OK to use zero commission amounts as this will be updated by UpdateCommissions
   
   '12/13/08 revised to do selects first to avoid unnecessary locks
   'there is some kind of deadlock situation.  maybe this will help.
   
   'if part is not commissionable, do nothing
   Dim rdo As rdoResultset
   Dim cExtPrice As Currency
   
   sSql = "select PACOMMISSION, " & vbCrLf _
      & " cast(ITQTY * ITDOLLARS as decimal(12,2))as ExtendedPrice " & vbCrLf _
      & " from PartTable" & vbCrLf _
      & " join SoItTable on ITPART = PARTREF" & vbCrLf _
      & " where ITSO = " & SoNumber & vbCrLf _
      & " and ITNUMBER = " & SoItem & vbCrLf _
      & " and ITREV = '" & SoRev & "'"
   If GetDataSet(rdo) Then
      If rdo!PACOMMISSION <> 1 Then
         Exit Sub
      Else
        cExtPrice = rdo!ExtendedPrice
      End If
      
   Else
      Exit Sub
   End If
   
   sSql = "select SMCOSO from SpcoTable" & vbCrLf _
      & "where SMCOSO = " & SoNumber & vbCrLf _
      & "and SMCOSOIT = " & SoItem & vbCrLf _
      & "and SMCOITREV = '" & SoRev & "'" & vbCrLf _
      & "and SMCOSM = '" & Salesman & "'"
   If GetDataSet(rdo) Then
      Exit Sub
   End If

    Dim cDefFltBase As Currency, cDefComPer As Currency
    ' Get default commission
    GetDefaultComm SoNumber, SoItem, _
                SoRev, Salesman, cExtPrice, _
                cDefFltBase, cDefComPer
   
   sSql = "INSERT INTO SpcoTable (SMCOSO,SMCOSOIT,SMCOITREV," & vbCrLf _
      & "SMCOSM,SMCOUSER,SMCOAMT,SMCOPCT)" & vbCrLf _
      & "VALUES (" & SoNumber & "," & SoItem _
      & ",'" & SoRev & "','" & Salesman & "','" _
      & sInitials & "', '" & CStr(cDefFltBase) & "','" & CStr(cDefComPer) & "')"
   RdoCon.Execute sSql, rdExecDirect
   
End Sub


Public Sub UpdateCommissions(SoNumber As Long, SoItem As Integer, SoRev As String, Optional bUserMsg As Boolean = True)

   'update commissions for all salesmen for a sales order item
   'do this as part of a transaction, inserting new commission records first with InsertCommission()
   
   '12/13/08 revised to do selects first to avoid unnecessary locks
   'there is some kind of deadlock situation.  maybe this will help.
   
   'if there are no commission records, there is nothing to do
   Dim rdoSales As rdoResultset, rdo As rdoResultset
   Dim cCurComPer As Currency, cComPer As Currency
   Dim cfixedAmount As Currency, cPerAmount As Currency
   Dim strSalesman As String
   
   sSql = "select SMCOSM,SMCOPCT,SMCOAMT from SpcoTable" & vbCrLf _
      & " where SMCOSO = " & SoNumber & vbCrLf _
      & " and SMCOSOIT = " & SoItem & vbCrLf _
      & " and SMCOITREV = '" & SoRev & "'"
      '& " AND SMCOSM = '" & strSalesman & "'"
   If Not GetDataSet(rdoSales) Then
      Exit Sub
   Else
      With rdoSales
         Do Until .EOF
            strSalesman = !SMCOSM
            cCurComPer = !SMCOPCT
            cfixedAmount = !SMCOAMT
   
            'get parameters for part
            Dim part As String, commissionable As Boolean, cExtPrice As Currency, percent As Currency
            
            sSql = "select PACOMMISSION, PARTREF," & vbCrLf _
               & "cast(ITQTY * ITDOLLARS as decimal(12,2))as ExtendedPrice" & vbCrLf _
               & "from PartTable" & vbCrLf _
               & "join SoItTable on ITPART = PARTREF" & vbCrLf _
               & "where ITSO = " & SoNumber & vbCrLf _
               & "and ITNUMBER = " & SoItem & vbCrLf _
               & "and ITREV = '" & SoRev & "'"
            
            If GetDataSet(rdo) Then
               part = rdo!PartRef
               commissionable = IIf(rdo!PACOMMISSION = 1, True, False)
               cExtPrice = rdo!ExtendedPrice
               cPerAmount = Format(cfixedAmount + (cExtPrice * cCurComPer / 100), CURRENCYMASK)
               ClearResultSet rdo
               Set rdo = Nothing
            End If
   
            'if the part is not commissionable, delete any commssions for this item
            If Not commissionable Then
               sSql = "delete from SpcoTable" & vbCrLf _
                  & "where SMCOSO = " & SoNumber & vbCrLf _
                  & "and SMCOSOIT = " & SoItem & " and SMCOITREV = '" & SoRev & "'"
               RdoCon.Execute sSql
            
            'otherwise, update commissions for this item where commission has changed
            Else
               
                 Dim cDefFltBase As Currency, cDefComPer As Currency
                 ' Get default commission
                 GetDefaultComm SoNumber, SoItem, _
                             SoRev, strSalesman, cExtPrice, _
                             cDefFltBase, cDefComPer
                 
                 Dim bResponse As Byte
                 Dim sMsg As String
                 
                 If cCurComPer <> cDefComPer And bUserMsg = True Then
               
                      sMsg = "The default commission for Sales man (" & strSalesman & ") is different" & vbCrLf _
                             & " Default value : " & CStr(cDefComPer) & vbCrLf _
                             & " Existing value : " & CStr(cCurComPer) & vbCrLf _
                             & " Do you want to set to default?"
                     
                      bResponse = MsgBox(sMsg, ES_YESQUESTION, "Information")
                      If bResponse = vbYes Then
                         cComPer = cDefComPer
                      Else
                         cComPer = cCurComPer
                      End If
                 Else
                     cComPer = cCurComPer
                 End If
                 
                 sSql = "UPDATE SPCOTable SET SMCOAMT = '" & cfixedAmount & "'," & vbCrLf _
                     & " SMCOPCT = '" & cComPer & "'," & vbCrLf _
                     & "SMCOREVISED = '" & Format(ES_SYSDATE, "mm/dd/yy") & "'" & vbCrLf _
                     & " FROM SpcoTable WHERE " & vbCrLf _
                     & " SMCOSO = " & SoNumber & vbCrLf _
                       & " and SMCOSOIT = " & SoItem & vbCrLf _
                       & " and SMCOITREV = '" & SoRev & "'" & vbCrLf _
                       & " AND SMCOSM = '" & strSalesman & "'"
                 
                 RdoCon.Execute sSql, rdExecDirect
                 
            End If
            .MoveNext
         Loop ' loop through all the commission sales man
         ClearResultSet rdoSales
      End With

   End If
End Sub

Public Sub GetDefaultComm(ByVal SoNumber As Long, ByVal SoItem As Integer, _
                    ByVal SoRev As String, ByVal Salesman As String, _
                    ByVal cAmt As Currency, ByRef cFltBase As Currency, _
                    ByRef cComPer As Currency)
   
   Dim RdoRte As rdoResultset
   Dim iList As Integer
   Dim cTotal As Currency
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT SPFROM1,SPTHRU1,SPBASE1,SPPERC1,SPFROM2,SPTHRU2," _
          & "SPBASE2,SPPERC2,SPFROM3,SPTHRU3,SPBASE3,SPPERC3,SPFROM4," _
          & "SPTHRU4,SPBASE4,SPPERC4,SPFROM5,SPTHRU5,SPBASE5,SPPERC5," _
          & "SPFROM6,SPTHRU6,SPBASE6,SPPERC6,SPFROM7,SPTHRU7,SPBASE7," _
          & "SPPERC7,SPFROM8,SPTHRU8,SPBASE8,SPPERC8,SPFROM9,SPTHRU9," _
          & "SPBASE9,SPPERC9,SPFROM10,SPTHRU10,SPBASE10,SPPERC10 " _
          & "FROM SprsTable WHERE SPNUMBER = '" & Salesman & "'"
   
   bSqlRows = GetDataSet(RdoRte)
   If bSqlRows Then
      With RdoRte
         For iList = 0 To 39 Step 4
            If CCur(.rdoColumns(iList)) <= cTotal And CCur(.rdoColumns(iList + 1)) >= cTotal Then
                cFltBase = .rdoColumns(iList + 2)
                cComPer = .rdoColumns(iList + 3)
                Exit For
            Else
                cFltBase = 0
                cComPer = 0
            End If
         Next
      End With
   End If
   Set RdoRte = Nothing
   Exit Sub
DiaErr1:
   sProcName = "GetDefaultComm"
End Sub


